import pandas as pd
import math
from datetime import datetime, timedelta

def process_extra_file(df_extra, service_name):
    result_pdt = {}
    result_ind = {}
    result_off = {}
    if df_extra is None or len(df_extra) == 0:
        return result_pdt, result_ind, result_off
    df_e = df_extra.copy()
    if '项目状态' in df_e.columns:
        df_e = df_e[df_e['项目状态'].isin(['已出票未月结', '已出票已月结'])]
    if len(df_e) == 0:
        return result_pdt, result_ind, result_off
    if '年月' not in df_e.columns or '系统部' not in df_e.columns:
        return result_pdt, result_ind, result_off
    raw_months = df_e['年月'].astype(str).str.strip()
    months_extracted = raw_months.str.extract(r'/(\d{2})$', expand=False)
    months_conv = months_extracted.dropna().astype(int).apply(lambda x: f'{x}月')
    industries = df_e['系统部'].astype(str).str.strip()
    if '办事处' in df_e.columns:
        office_col = '办事处'
    elif '代表处' in df_e.columns:
        office_col = '代表处'
    elif '地市' in df_e.columns:
        office_col = '地市'
    else:
        return result_pdt, result_ind, result_off
    offices_raw = df_e[office_col].astype(str).str.strip()
    aliases = {'西宁地区': '西宁'}
    office_to_region = {
        '天津': '天津市', '黑龙江': '黑龙江', '吉林': '吉林省', '辽宁': '辽宁省',
        '内蒙古': '内蒙古', '山东': '山东省', '河南': '河南省', '山西': '山西省',
        '河北': '河北省', '湖北': '湖北省', '陕西': '陕西省', '甘肃': '甘肃省',
        '宁夏': '宁夏省', '青海': '青海省', '新疆': '新疆省', '江苏': '江苏省',
        '福建': '福建省', '浙江': '浙江省', '上海': '上海市', '江西': '江西省',
        '安徽': '安徽省', '云南': '云南省', '广西': '广西壮族', '西藏': '西藏省',
        '四川': '四川省', '湖南': '湖南省', '贵州': '贵州省', '重庆': '重庆市',
        '海南': '海南省', '广州': '广州市', '深圳': '深圳市',
        '香港': '香港', '澳门': '澳门',
    }
    offices_mapped = offices_raw.map(lambda x: office_to_region.get(x, aliases.get(x, x)))
    if '营业收入' in df_e.columns:
        amounts = pd.to_numeric(df_e['营业收入'], errors='coerce').fillna(0) / 10000.0
    else:
        amounts = pd.Series([0.0] * len(df_e), index=df_e.index)

    if months_conv.notna().any():
        for month, amt in amounts.groupby(months_conv).sum().items():
            result_pdt[(service_name, month)] = result_pdt.get((service_name, month), 0) + amt
    for ind, amt in amounts.groupby(industries).sum().items():
        if ind and ind != 'nan':
            result_ind[ind] = {service_name: amt}
    for idx, row in df_e.iterrows():
        off_raw = offices_mapped.loc[idx]
        amt = amounts.loc[idx]
        ind_val = industries.loc[idx]
        if off_raw in ('总部', '北京') and ind_val and ind_val != 'nan':
            office_key = ind_val
        else:
            office_key = off_raw
        if office_key and office_key != 'nan':
            if office_key not in result_off:
                result_off[office_key] = {}
            result_off[office_key][service_name] = result_off[office_key].get(service_name, 0) + amt
    return result_pdt, result_ind, result_off

def process_maoli_order(df, week_start=None, week_end=None,
                       df_mgmt=None, df_oem=None, df_manual=None):
    """订单毛利表处理函数 - 与process_order完全相同的逻辑"""
    column_mapping = {
        '月': ['月', '月份', 'month'],
        '办事处': ['办事处', 'office'],
        '行业': ['行业', 'industry'],
        '服务bom': ['服务bom', 'bom', '服务BOM'],
        '产品族描述': ['产品族描述', 'product_family', '产品族'],
        '文本总金额': ['文本总金额', '金额', '总金额', 'amount'],
        '合同签定日期': ['合同签定日期', '合同签订日期', '日期', 'date'],
        '币种': ['币种', 'currency', '货币']
    }
    rename_dict = {}
    for target, aliases in column_mapping.items():
        for alias in aliases:
            if alias in df.columns:
                rename_dict[alias] = target
                break
    if rename_dict:
        df.rename(columns=rename_dict, inplace=True)
    
    required_cols = ['月', '办事处', '行业', '服务bom', '产品族描述', '文本总金额', '合同签定日期', '币种']
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f'Excel文件中缺少必需列：{col}。实际列名：{df.columns.tolist()}')

    currency_rates = {'人民币': 1.0, 'cny': 1.0, 'usd': 7.0187, 'hkd': 0.90131}
    raw_currencies = df['币种'].astype(str).str.strip().str.lower()
    rates = raw_currencies.map(currency_rates)
    if rates.isna().any():
        unknown = df.loc[rates.isna(), '币种'].unique()
        raise ValueError(f'未知币种：{unknown}')
    df['换算后金额'] = pd.to_numeric(df['文本总金额'], errors='coerce').fillna(0) * rates

    months = df['月'].astype(str).str.strip()
    offices = df['办事处'].astype(str).str.strip()
    office_aliases = {'宁夏': '银川', '西宁地区': '西宁'}
    offices = offices.map(lambda x: office_aliases.get(x, x))
    industries = df['行业'].astype(str).str.strip()
    boms = df['服务bom'].astype(str).str.strip()
    product_families = df['产品族描述'].astype(str).str.strip()
    amounts = df['换算后金额']
    
    dates = pd.to_datetime(df['合同签定日期'], errors='coerce')
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    if week_start is not None and week_end is not None:
        week_mask = (dates >= pd.Timestamp(week_start)) & (dates <= pd.Timestamp(week_end))
    else:
        week_ago = today - timedelta(days=7)
        week_mask = (dates >= week_ago) & (dates <= today)

    pdt1, ind1, off1 = process_extra_file(df_mgmt, '数据安全管理中心')
    pdt2, ind2, off2 = process_extra_file(df_oem, '数据安全OEM IN')

    def sumif(bom_list=None, month=None, product_family=None, industry=None, office=None, use_week_mask=False, hq_industry=None):
        mask = pd.Series([True] * len(df))
        if bom_list is not None:
            mask &= boms.isin(bom_list)
        if month is not None:
            mask &= (months == month)
        if product_family is not None:
            mask &= (product_families == product_family)
        if industry is not None:
            mask &= (industries == industry)
        if office is not None:
            mask &= (offices == office)
        if hq_industry is not None:
            mask &= offices.isin(['北京', '总部'])
            mask &= (industries == hq_industry)
        if use_week_mask:
            mask &= week_mask
        return amounts[mask].sum() / 10000.0

    bom_co_wei = ['8814A12V', '8814A0F7', '8814A0F8', '8814A0F9']
    exclude_boms_op = ['8814A0SB', '8814A12V', '8814A0F7', '8814A0F8', '8814A0F9', '8814A197', '8814A0VX']
    bom_gong_fang = ['8814A0JP', '8814A0BF']
    bom_consult = ['8814A0BJ', '8814A0FA', '8814A0BV']
    bom_tool = ['8814A0SB', '8814A0VX']
    bom_overseas = ['8814A0S5']
    bom_education = ['8814A0H3']
    bom_data_security = ['8814A197']
    bom_integration = ['8814A09W', '8814A0G9', '8814A0LQ']

    months_list = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月']

    pdt_data = {}
    week_data = {}
    categories = {
        '综合服务': ['安全驻场服务', '云端运营服务', '安全攻防服务', '安全咨询服务', '安全服务工具', '海外安全服务'],
        '数据安全': ['数据安全管理中心', '数据安全OEM IN', '数据安全产品转售', '数据安全服务'],
        '产教融合': ['产教融合产品转售', '产教融合服务'],
        '安全集成服务': ['安全集成服务']
    }

    for month in months_list:
        for cat, subcats in categories.items():
            for sub in subcats:
                key = (sub, month)
                if sub == '安全驻场服务':
                    val = sumif(bom_list=bom_co_wei, month=month)
                elif sub == '云端运营服务':
                    total = sumif(product_family='安全运维服务', month=month)
                    excl = sumif(bom_list=exclude_boms_op, month=month)
                    val = total - excl
                elif sub == '安全攻防服务':
                    val = sumif(bom_list=bom_gong_fang, month=month)
                elif sub == '安全咨询服务':
                    val = sumif(bom_list=bom_consult, month=month)
                elif sub == '安全服务工具':
                    val = sumif(bom_list=bom_tool, month=month)
                elif sub == '海外安全服务':
                    val = sumif(bom_list=bom_overseas, month=month)
                elif sub == '产教融合服务':
                    val = sumif(bom_list=bom_education, month=month)
                elif sub == '数据安全服务':
                    val = sumif(bom_list=bom_data_security, month=month)
                elif sub == '安全集成服务':
                    val = sumif(bom_list=bom_integration, month=month)
                else:
                    val = 0.0
                pdt_data[key] = val

    for cat, subcats in categories.items():
        for sub in subcats:
            if sub == '安全驻场服务':
                week_val = sumif(bom_list=bom_co_wei, use_week_mask=True)
            elif sub == '云端运营服务':
                week_total = sumif(product_family='安全运维服务', use_week_mask=True)
                week_excl = sumif(bom_list=exclude_boms_op, use_week_mask=True)
                week_val = week_total - week_excl
            elif sub == '安全攻防服务':
                week_val = sumif(bom_list=bom_gong_fang, use_week_mask=True)
            elif sub == '安全咨询服务':
                week_val = sumif(bom_list=bom_consult, use_week_mask=True)
            elif sub == '安全服务工具':
                week_val = sumif(bom_list=bom_tool, use_week_mask=True)
            elif sub == '海外安全服务':
                week_val = sumif(bom_list=bom_overseas, use_week_mask=True)
            elif sub == '产教融合服务':
                week_val = sumif(bom_list=bom_education, use_week_mask=True)
            elif sub == '数据安全服务':
                week_val = sumif(bom_list=bom_data_security, use_week_mask=True)
            elif sub == '安全集成服务':
                week_val = sumif(bom_list=bom_integration, use_week_mask=True)
            else:
                week_val = 0.0
            week_data[sub] = week_val

    for src in (pdt1, pdt2):
        for key, val in src.items():
            pdt_data[key] = pdt_data.get(key, 0) + val

    pdt_rows = []
    integration_rows_data = []
    targets = {'综合服务': 18000, '数据安全': 7000, '产教融合': 5000}
    for cat, subcats in categories.items():
        cat_rows = []
        for sub in subcats:
            row = {'二级分类': cat, '三级分类': sub}
            month_values = {}
            for month in months_list:
                month_values[month] = pdt_data.get((sub, month), 0)
                row[month] = month_values[month]
            q1 = month_values['1月'] + month_values['2月'] + month_values['3月']
            q2 = month_values['4月'] + month_values['5月'] + month_values['6月']
            q3 = month_values['7月'] + month_values['8月'] + month_values['9月']
            q4 = month_values['10月'] + month_values['11月'] + month_values['12月']
            row['Q1小计'] = q1
            row['Q2小计'] = q2
            row['Q3小计'] = q3
            row['Q4小计'] = q4
            row['合计'] = q1+q2+q3+q4
            row['本周新增'] = week_data.get(sub, 0)
            row['2026订单目标'] = ''
            row['完成率'] = ''
            cat_rows.append(row)
        if cat in targets:
            target = targets[cat]
            cat_total = sum(r['合计'] for r in cat_rows)
            cat_rows[0]['2026订单目标'] = target
            cat_rows[0]['完成率'] = f'{round(cat_total / target * 100, 2)}%'
        if cat == '安全集成服务':
            integration_rows_data = cat_rows
        else:
            pdt_rows.extend(cat_rows)
            if cat in targets:
                subtotal = {'二级分类': cat, '三级分类': cat + '小计'}
                for month in months_list:
                    subtotal[month] = sum(r[month] for r in cat_rows)
                for prefix in ['Q1', 'Q2', 'Q3', 'Q4']:
                    col_name = prefix + '小计'
                    subtotal[col_name] = sum(r[col_name] for r in cat_rows)
                subtotal['合计'] = cat_total
                subtotal['本周新增'] = sum(r['本周新增'] for r in cat_rows)
                subtotal['2026订单目标'] = ''
                subtotal['完成率'] = ''
                pdt_rows.append(subtotal)
    data_rows = [r for r in pdt_rows if r['三级分类'] and not r['三级分类'].endswith('小计') and r['二级分类'] != '安全集成服务']
    total_row = {'二级分类': '服务产品经理业绩总计', '三级分类': ''}
    for col in months_list + ['Q1小计', 'Q2小计', 'Q3小计', 'Q4小计', '合计', '本周新增']:
        total_row[col] = sum(r[col] for r in data_rows)
    total_row['2026订单目标'] = 30000
    total_val = total_row['合计']
    total_row['完成率'] = f'{round(total_val / 30000 * 100, 2)}%' if total_val else '0.00%'
    pdt_rows.append(total_row)
    pdt_rows.extend(integration_rows_data)
    column_order = ['二级分类', '三级分类'] + \
                   ['1月', '2月', '3月', 'Q1小计'] + \
                   ['4月', '5月', '6月', 'Q2小计'] + \
                   ['7月', '8月', '9月', 'Q3小计'] + \
                   ['10月', '11月', '12月', 'Q4小计', '合计', '本周新增', '2026订单目标', '完成率']

    dept_mapping = {
        '安平': ['安平', 'Public Sector', '检法司'],
        '电力能源': ['电网', '能源', '煤炭发电'],
        '互联网': ['互联网系统二部', '互联网系统一部', '代表处互联网', 'Internet'],
        '运营商': ['电信总部', '电信战略客户', '广电网络', '广电媒体', '广电', '广电总部',
                  '联通战略客户', '联通', '联通总部', '移动战略客户', '移动总部'],
        '企业': ['全球大客户', '外资客户', '现代服务', 'Distribution and Services', 'Consumer',
                 '央企与智能制造', 'Enterprises', 'Manufacturing and Resources', '智能制造',
                 '央国企', '央国企系统一部', '央国企系统二部', 'AI与创新科技', '阿里', '其他一部', '其他二部'],
        '数字政府': ['财税民生', 'Finance', '党政', 'Government', '政务', '数字政务'],
        '内部销售': ['内部销售'],
        '分销': ['分销', 'Infrastructure'],
        '金融': ['保险', '银行三部', '银行二部', '银行一部', '证券'],
        '交通': ['轨道', '交通综合', '代理销售-交通-铁路', '民航及水运', 'Carrier'],
        '教育': ['高教', '教育部及职教', '教育科研', '教育', '科研及智慧教育', '部委及职教', '科研及北京教育'],
        '医疗': ['医保及北京销售部', '医疗', '医院及公共卫生'],
        '专网': ['军工', '专网'],
    }
    industry_to_dept = {}
    ordered_industries = []
    for dept, inds in dept_mapping.items():
        for ind in inds:
            industry_to_dept[ind] = dept
            ordered_industries.append(ind)

    numeric_cols = ['安全协维','安全运营','攻防服务','专家服务','安全产品转售',
                    '海外服务','数据安全管理中心','数据安全OEM IN','数据安全产品转售','数据安全服务',
                    '产教融合产品转售','产教融合服务','服务小计','安全集成服务']
    sub_keys = ['安全协维','安全运营','攻防服务','专家服务','安全产品转售','海外服务',
                '数据安全管理中心','数据安全OEM IN','数据安全产品转售','数据安全服务',
                '产教融合产品转售','产教融合服务']

    industry_rows = []
    industry_raws = []
    for ind in ordered_industries:
        dept = industry_to_dept.get(ind, '')
        raw = {}
        raw['安全协维'] = sumif(bom_list=bom_co_wei, industry=ind)
        total_op = sumif(product_family='安全运维服务', industry=ind)
        excl_op = sumif(bom_list=exclude_boms_op, industry=ind)
        raw['安全运营'] = total_op - excl_op
        raw['攻防服务'] = sumif(bom_list=bom_gong_fang, industry=ind)
        raw['专家服务'] = sumif(bom_list=bom_consult, industry=ind)
        raw['安全产品转售'] = sumif(bom_list=bom_tool, industry=ind)
        raw['海外服务'] = sumif(bom_list=bom_overseas, industry=ind)
        raw['数据安全管理中心'] = sum(val for key, val in ind1.get(ind, {}).items()
                               if key == '数据安全管理中心')
        raw['数据安全OEM IN'] = sum(val for key, val in ind2.get(ind, {}).items()
                                if key == '数据安全OEM IN')
        raw['数据安全产品转售'] = 0.0
        raw['数据安全服务'] = sumif(bom_list=bom_data_security, industry=ind)
        raw['产教融合产品转售'] = 0.0
        raw['产教融合服务'] = sumif(bom_list=bom_education, industry=ind)
        raw['服务小计'] = sum(raw[k] for k in sub_keys)
        raw['安全集成服务'] = sumif(bom_list=bom_integration, industry=ind)
        industry_raws.append(raw)
        row = {'事业部': dept, '行业': ind}
        for col in numeric_cols:
            row[col] = raw[col]
        industry_rows.append(row)

    total_row = {'事业部': '', '行业': '总计'}
    for col in numeric_cols:
        total_row[col] = sum(r[col] for r in industry_raws)
    industry_rows.append(total_row)

    df_industry_calc = pd.DataFrame(industry_rows)

    merge_rules = {
        '互联网系统一部': ['互联网系统一部', '代表处互联网', 'Internet'],
        '电信': ['电信总部', '电信战略客户'],
        '广电': ['广电网络', '广电媒体', '广电', '广电总部'],
        '联通': ['联通战略客户', '联通', '联通总部'],
        '移动': ['移动战略客户', '移动总部'],
        '全球大客户': ['全球大客户', '外资客户', '阿里'],
        '智能制造': ['Manufacturing and Resources', '智能制造'],
        '央国企': ['央国企', '央国企系统一部', '央国企系统二部'],
        '其他': ['其他一部', '其他二部', '现代服务', 'Distribution and Services', 'Consumer', '央企与智能制造', 'Enterprises'],
        '财税民生': ['财税民生', 'Finance'],
        '党政': ['党政', 'Government'],
        '安平': ['安平', 'Public Sector', '检法司'],
        '分销': ['分销', 'Infrastructure'],
        '交通综合': ['交通综合', '代理销售-交通-铁路', '民航及水运', 'Carrier'],
        '部委及职教': ['教育部及职教', '部委及职教'],
        '教育': ['教育科研', '教育'],
        '科研及北京教育': ['科研及智慧教育', '科研及北京教育'],
        '医院及公共卫生': ['医疗', '医院及公共卫生'],
    }

    new_dept_mapping = {
        '电力能源': ['电网', '能源', '煤炭发电'],
        '互联网': ['互联网系统二部', '互联网系统一部'],
        '运营商': ['电信', '广电', '联通', '移动'],
        '企业': ['全球大客户', '智能制造', '央国企', 'AI与创新科技', '其他', '中资海外'],
        '数字政府': ['财税民生', '党政', '政务', '数字政务', '安平'],
        '内部销售': ['内部销售'],
        '分销': ['分销'],
        '金融': ['保险', '银行三部', '银行二部', '银行一部', '证券'],
        '交通': ['轨道', '交通综合'],
        '教育': ['高教', '部委及职教', '教育', '科研及北京教育'],
        '医疗': ['医保及北京销售部', '医院及公共卫生'],
        '专网': ['军工', '专网'],
    }

    new_industry_rows = []
    new_industry_raws = []
    for dept, inds in new_dept_mapping.items():
        for ind in inds:
            sources = merge_rules.get(ind, [ind])
            raw = {}
            for col in numeric_cols:
                raw[col] = sum(ir_raw[col] for ir_raw, ir in zip(industry_raws, industry_rows)
                              if ir['行业'] in sources)
            new_industry_raws.append(raw)
            row = {'事业部': dept, '行业': ind}
            for col in numeric_cols:
                row[col] = raw[col]
            new_industry_rows.append(row)

    total_row_new = {'事业部': '', '行业': '总计'}
    for col in numeric_cols:
        total_row_new[col] = sum(r[col] for r in new_industry_raws)
    new_industry_rows.append(total_row_new)

    df_industry_new = pd.DataFrame(new_industry_rows)

    region_mapping = {
        '总部企业': ['全球大客户', '外资客户', '现代服务', '央企与智能制造', '智能制造',
                    '央国企', '央国企系统一部', '央国企系统二部', 'AI与创新科技', '其他一部', '其他二部'],
        '总部运营商': ['电信总部', '电信战略客户', '广电网络', '广电媒体', '广电', '广电总部',
                     '联通战略客户', '联通', '联通总部', '移动战略客户', '移动总部'],
        '总部数字政府': ['财税民生', '党政', '政务', '数字政务'],
        '总部金融': ['保险', '银行三部', '银行二部', '银行一部', '证券'],
        '总部分销': ['分销'],
        '总部内销': ['内部销售'],
        '总部军工': ['军工'],
        '总部安平': ['安平', '检法司'],
        '总部专网': ['专网'],
        '总部电力能源': ['电网', '能源', '煤炭发电'],
        '总部互联网': ['互联网系统二部', '互联网系统一部'],
        '总部交通': ['轨道', '交通综合', '代理销售-交通-铁路', '民航及水运'],
        '总部教育': ['高教', '教育科研', '教育', '教育部及职教', '科研及北京教育', '部委及职教'],
        '总部医疗': ['医保及北京销售部', '医疗', '医院及公共卫生'],
        '天津市': ['天津'],
        '黑龙江': ['哈尔滨'],
        '吉林省': ['长春'],
        '辽宁省': ['沈阳'],
        '内蒙古': ['呼和浩特'],
        '山东省': ['济南'],
        '河南省': ['郑州'],
        '山西省': ['太原'],
        '河北省': ['石家庄'],
        '湖北省': ['武汉'],
        '陕西省': ['西安'],
        '甘肃省': ['兰州'],
        '宁夏省': ['银川'],
        '青海省': ['西宁'],
        '新疆省': ['乌鲁木齐'],
        '江苏省': ['南京'],
        '福建省': ['福州'],
        '浙江省': ['杭州'],
        '上海市': ['上海'],
        '江西省': ['南昌'],
        '安徽省': ['合肥'],
        '云南省': ['昆明'],
        '广西壮族': ['南宁'],
        '西藏省': ['拉萨地区'],
        '四川省': ['成都'],
        '湖南省': ['长沙'],
        '贵州省': ['贵阳'],
        '重庆市': ['重庆'],
        '海南省': ['海南'],
        '广州市': ['广州'],
        '深圳市': ['深圳'],
        '香港': ['香港'],
        '澳门': ['澳门'],
        '欧洲': ['欧洲'],
        '墨西哥': ['Mexico'],
        '俄罗斯': ['Russia'],
        '阿拉伯': ['United Arab Emirates'],
        '巴基斯坦': ['Pakistan'],
        '马来西亚': ['Malaysia'],
        '菲律宾': ['Philippines'],
        '俄罗斯联邦': ['Russian Federation'],
        '印度尼西亚': ['Indonesia'],
        '西班牙': ['Spain'],
        '南非': ['South Africa'],
        '哈萨克斯坦': ['Kazakhstan'],
        '越南': ['Vietnam'],
        '日本': ['Japan'],
        '泰国': ['Thailand'],
        '新加坡': ['Singapore'],
        '孟加拉': ['Bangladesh'],
        '沙特阿拉伯': ['Saudi Arabia'],
    }
    office_to_region = {}
    ordered_offices = []
    for region, offs in region_mapping.items():
        for off in offs:
            office_to_region[off] = region
            ordered_offices.append(off)

    office_numeric_cols = ['安全协维','安全运营','攻防服务','专家服务','安全产品转售',
                           '海外服务','数据安全管理中心','数据安全OEM IN','数据安全产品转售','数据安全服务',
                           '产教融合产品转售','产教融合服务','服务小计','安全集成服务']

    office_rows = []
    office_raws = []
    for off in ordered_offices:
        region = office_to_region.get(off, '')
        is_hq = region.startswith('总部')
        skw = {'hq_industry': off} if is_hq else {'office': off}
        raw = {}
        raw['安全协维'] = sumif(bom_list=bom_co_wei, **skw)
        total_op = sumif(product_family='安全运维服务', **skw)
        excl_op = sumif(bom_list=exclude_boms_op, **skw)
        raw['安全运营'] = total_op - excl_op
        raw['攻防服务'] = sumif(bom_list=bom_gong_fang, **skw)
        raw['专家服务'] = sumif(bom_list=bom_consult, **skw)
        raw['安全产品转售'] = sumif(bom_list=bom_tool, **skw)
        raw['海外服务'] = sumif(bom_list=bom_overseas, **skw)
        if is_hq:
            raw['数据安全管理中心'] = off1.get(off, {}).get('数据安全管理中心', 0)
            raw['数据安全OEM IN'] = off2.get(off, {}).get('数据安全OEM IN', 0)
        else:
            region_short = region.rstrip('省市区')
            mgmt_val = 0.0
            oem_val = 0.0
            for try_key in (off, region, region_short):
                mgmt_val = off1.get(try_key, {}).get('数据安全管理中心', 0)
                if mgmt_val != 0:
                    break
            for try_key in (off, region, region_short):
                oem_val = off2.get(try_key, {}).get('数据安全OEM IN', 0)
                if oem_val != 0:
                    break
            raw['数据安全管理中心'] = mgmt_val
            raw['数据安全OEM IN'] = oem_val
        raw['数据安全产品转售'] = 0.0
        raw['数据安全服务'] = sumif(bom_list=bom_data_security, **skw)
        raw['产教融合产品转售'] = 0.0
        raw['产教融合服务'] = sumif(bom_list=bom_education, **skw)
        raw['服务小计'] = sum(raw[k] for k in sub_keys)
        raw['安全集成服务'] = sumif(bom_list=bom_integration, **skw)
        office_raws.append(raw)
        row = {'区域': region, '办事处': off}
        for col in office_numeric_cols:
            row[col] = raw[col]
        office_rows.append(row)

    total_row_office = {'区域': '', '办事处': '总计'}
    for col in office_numeric_cols:
        total_row_office[col] = sum(r[col] for r in office_raws)
    office_rows.append(total_row_office)

    df_office_calc = pd.DataFrame(office_rows)

    region_raw = {}
    for off_raw, off_row in zip(office_raws, office_rows):
        region_name = off_row['区域']
        if region_name not in region_raw:
            region_raw[region_name] = {col: 0.0 for col in office_numeric_cols}
        for col in office_numeric_cols:
            region_raw[region_name][col] += off_raw[col]

    big_region_mapping = {
        '京津区域': ['总部企业', '总部运营商', '总部数字政府', '总部金融', '总部分销', '总部内销',
                    '总部军工', '总部安平', '总部专网', '总部电力能源', '总部互联网',
                    '总部交通', '总部教育', '总部医疗', '天津市'],
        '华北区域': ['河南省', '山西省', '河北省'],
        '东北区域': ['黑龙江', '吉林省', '辽宁省', '内蒙古', '山东省'],
        '西北区域': ['陕西省', '甘肃省', '宁夏省', '青海省', '新疆省'],
        '华东区域': ['江苏省', '福建省', '浙江省', '上海市', '江西省', '安徽省', '湖北省'],
        '华南区域': ['云南省', '广西壮族', '西藏省', '四川省', '湖南省', '贵州省', '重庆市', '海南省',
                    '广州市', '深圳市', '香港', '澳门'],
    }
    foreign_candidates = ['欧洲', '墨西哥', '俄罗斯', '阿拉伯', '巴基斯坦', '马来西亚', '菲律宾',
                          '俄罗斯联邦', '印度尼西亚', '西班牙', '南非', '哈萨克斯坦', '越南', '日本',
                          '泰国', '新加坡', '孟加拉', '沙特阿拉伯']

    new_off_rows = []
    new_off_raws = []
    big_region_order = ['京津区域', '华北区域', '东北区域', '西北区域', '华东区域', '华南区域']
    for big_region in big_region_order:
        for region_name in big_region_mapping[big_region]:
            raw = region_raw.get(region_name, {col: 0.0 for col in office_numeric_cols})
            new_off_raws.append(raw)
            row = {'区域': big_region, '省办': region_name}
            for col in office_numeric_cols:
                row[col] = raw[col]
            new_off_rows.append(row)

    foreign_added = []
    for region_name in foreign_candidates:
        raw = region_raw.get(region_name)
        if raw is not None and raw['服务小计'] != 0:
            new_off_raws.append(raw)
            row = {'区域': '国外区域', '省办': region_name}
            for col in office_numeric_cols:
                row[col] = raw[col]
            new_off_rows.append(row)
            foreign_added.append(region_name)

    total_row_new_off = {'区域': '', '省办': '总计'}
    for col in office_numeric_cols:
        total_row_new_off[col] = sum(r[col] for r in new_off_raws)
    new_off_rows.append(total_row_new_off)

    df_office_new = pd.DataFrame(new_off_rows)

    if df_manual is not None:
        manual_keys = list(df_manual.keys())
        manual_pdt = df_manual[manual_keys[0]]
        manual_ind = df_manual[manual_keys[1]] if len(manual_keys) > 1 else None
        manual_off = df_manual[manual_keys[2]] if len(manual_keys) > 2 else None

        print(f'[毛利手工数据] sheet数={len(manual_keys)}, sheets={manual_keys}')

        months = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']
        manual_service_cols = ['数据安全管理中心', '数据安全OEM IN', '数据安全产品转售', '产教融合产品转售']
        service_cols = ['安全协维','安全运营','攻防服务','专家服务','安全产品转售',
                        '海外服务','数据安全管理中心','数据安全OEM IN','数据安全产品转售','数据安全服务',
                        '产教融合产品转售','产教融合服务','安全集成服务']
        sub_keys = ['安全协维','安全运营','攻防服务','专家服务','安全产品转售','海外服务',
                    '数据安全管理中心','数据安全OEM IN','数据安全产品转售','数据安全服务',
                    '产教融合产品转售','产教融合服务']

        if manual_pdt is not None:
            manual_months = [m for m in months if m in manual_pdt.columns]
            for i, (mrow, prow) in enumerate(zip(manual_pdt.itertuples(index=False), pdt_rows)):
                if not prow['三级分类'] or (prow['三级分类'] or '').endswith('小计') or prow['二级分类'] == '服务产品经理业绩总计':
                    continue
                for month in manual_months:
                    col_idx = manual_pdt.columns.get_loc(month)
                    try:
                        val = float(mrow[col_idx])
                        if math.isnan(val):
                            val = 0.0
                    except (ValueError, TypeError):
                        val = 0.0
                    if val != 0:
                        old_val = prow.get(month, 0)
                        prow[month] = old_val + val
            for row in pdt_rows:
                if not row['三级分类'] or (row['三级分类'] or '').endswith('小计') or row['二级分类'] == '服务产品经理业绩总计':
                    continue
                q1 = row.get('1月',0)+row.get('2月',0)+row.get('3月',0)
                q2 = row.get('4月',0)+row.get('5月',0)+row.get('6月',0)
                q3 = row.get('7月',0)+row.get('8月',0)+row.get('9月',0)
                q4 = row.get('10月',0)+row.get('11月',0)+row.get('12月',0)
                row['Q1小计'] = q1
                row['Q2小计'] = q2
                row['Q3小计'] = q3
                row['Q4小计'] = q4
                row['合计'] = q1+q2+q3+q4
            for cat in ['综合服务', '数据安全', '产教融合']:
                cat_rows = [r for r in pdt_rows if r['二级分类']==cat and not (r['三级分类'] or '').endswith('小计')]
                sub_row = next((r for r in pdt_rows if r['二级分类']==cat and (r['三级分类'] or '').endswith('小计')), None)
                if sub_row and cat_rows:
                    for col in months + ['Q1小计','Q2小计','Q3小计','Q4小计','合计','本周新增']:
                        sub_row[col] = sum(r.get(col,0) for r in cat_rows)
            tgt_row = next((r for r in pdt_rows if r['二级分类']=='服务产品经理业绩总计'), None)
            if tgt_row:
                data_rows_t = [r for r in pdt_rows if r['三级分类'] and not (r['三级分类'] or '').endswith('小计') and r['二级分类']!='安全集成服务']
                for col in months + ['Q1小计','Q2小计','Q3小计','Q4小计','合计','本周新增']:
                    tgt_row[col] = sum(r.get(col,0) for r in data_rows_t)
            print(f'[毛利手工-PDT] 数据加和完成')

        if manual_ind is not None:
            ind_rows_list = list(df_industry_new.itertuples(index=False))
            for i, (mrow, orow) in enumerate(zip(manual_ind.itertuples(index=False), ind_rows_list)):
                if orow.行业 == '总计':
                    continue
                for col in service_cols:
                    if col in manual_ind.columns:
                        col_idx = manual_ind.columns.get_loc(col)
                        try:
                            val = float(mrow[col_idx])
                            if math.isnan(val):
                                val = 0.0
                        except (ValueError, TypeError):
                            val = 0.0
                        if val != 0:
                            df_industry_new.at[i, col] = df_industry_new.at[i, col] + val
            for i in range(len(df_industry_new)):
                if df_industry_new.at[i, '行业'] == '总计':
                    continue
                df_industry_new.at[i, '服务小计'] = sum(df_industry_new.at[i, c] for c in sub_keys)
            total_mask = df_industry_new['行业'] == '总计'
            if total_mask.any():
                total_idx = df_industry_new[total_mask].index[0]
                for col in service_cols + ['服务小计']:
                    if col in df_industry_new.columns:
                        df_industry_new.at[total_idx, col] = sum(df_industry_new.at[j, col] for j in df_industry_new.index if j != total_idx)
            print(f'[毛利手工-行业报表] 所有列数据加和完成')

        if manual_off is not None:
            off_rows_list = list(df_office_new.itertuples(index=False))
            for i, (mrow, orow) in enumerate(zip(manual_off.itertuples(index=False), off_rows_list)):
                if orow.省办 == '总计':
                    continue
                for col in service_cols:
                    if col in manual_off.columns:
                        col_idx = manual_off.columns.get_loc(col)
                        try:
                            val = float(mrow[col_idx])
                            if math.isnan(val):
                                val = 0.0
                        except (ValueError, TypeError):
                            val = 0.0
                        if val != 0:
                            df_office_new.at[i, col] = df_office_new.at[i, col] + val
            for i in range(len(df_office_new)):
                if df_office_new.at[i, '省办'] == '总计':
                    continue
                df_office_new.at[i, '服务小计'] = sum(df_office_new.at[i, c] for c in sub_keys)
            total_mask = df_office_new['省办'] == '总计'
            if total_mask.any():
                total_idx = df_office_new[total_mask].index[0]
                for col in service_cols + ['服务小计']:
                    if col in df_office_new.columns:
                        df_office_new.at[total_idx, col] = sum(df_office_new.at[j, col] for j in df_office_new.index if j != total_idx)
            print(f'[毛利手工-办事处报表] 所有列数据加和完成')

    pdt_numeric = months_list + ['Q1小计','Q2小计','Q3小计','Q4小计','合计','本周新增']
    for row in pdt_rows:
        for col in pdt_numeric:
            if col in row and isinstance(row[col], (int, float)):
                row[col] = round(row[col])

    df_pdt = pd.DataFrame(pdt_rows, columns=column_order)
    for df_ref in [df_industry_new, df_office_new]:
        for col in numeric_cols:
            if col in df_ref.columns:
                df_ref[col] = pd.to_numeric(df_ref[col], errors='coerce').fillna(0).round().astype(int)

    return {
        '毛利_PDT产品线': df_pdt,
        '毛利_行业计算表': df_industry_calc,
        '毛利_行业报表': df_industry_new,
        '毛利_办事处计算表': df_office_calc,
        '毛利_办事处报表': df_office_new
    }

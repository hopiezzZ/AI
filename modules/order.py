import pandas as pd
import math
import shutil
import tempfile
import os
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (QVBoxLayout, QLabel, QTableWidget, QTableWidgetItem,
                             QPushButton, QHBoxLayout, QHeaderView, QMessageBox,
                             QProgressDialog, QDateEdit, QFileDialog, QFrame, QDialog,
                             QLineEdit, QWidget, QScrollArea)
from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
from modules.base import BaseModule, ReportThread, apply_excel_style
from modules.process_maoli import process_maoli_order


class YearInputDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("输入25年合计数据")
        self.setModal(True)
        self.setMinimumWidth(450)
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QLabel {
                color: #333333;
                font-size: 14px;
            }
            QLineEdit {
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                padding: 8px 12px;
                font-size: 14px;
                background-color: #ffffff;
            }
            QLineEdit:focus {
                border-color: #1890ff;
            }
            QPushButton {
                background-color: #1890ff;
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 10px 24px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #40a9ff;
            }
            QPushButton:pressed {
                background-color: #096dd9;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        title_label = QLabel("请输入25年合计数据（单位：万元）")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #1890ff; margin-bottom: 8px;")
        layout.addWidget(title_label)

        self.input_zonghe = QLineEdit()
        self.input_zonghe.setPlaceholderText("综合服务 25年合计")
        layout.addWidget(QLabel("综合服务 25年合计："))
        layout.addWidget(self.input_zonghe)

        self.input_shuju = QLineEdit()
        self.input_shuju.setPlaceholderText("数据安全 25年合计")
        layout.addWidget(QLabel("数据安全 25年合计："))
        layout.addWidget(self.input_shuju)

        self.input_chanjiao = QLineEdit()
        self.input_chanjiao.setPlaceholderText("产教融合 25年合计")
        layout.addWidget(QLabel("产教融合 25年合计："))
        layout.addWidget(self.input_chanjiao)

        note_label = QLabel("注：总计（W17）将自动计算生成，同比数据将自动计算")
        note_label.setStyleSheet("font-size: 12px; color: #999999; margin-top: 8px;")
        layout.addWidget(note_label)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)
        spacer = QWidget()
        spacer.setMinimumWidth(100)
        btn_layout.addWidget(spacer)

        ok_btn = QPushButton("确认")
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("取消")
        cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #ffffff;
                color: #666666;
                border: 1px solid #d9d9d9;
                border-radius: 6px;
                padding: 10px 24px;
                font-size: 14px;
            }
            QPushButton:hover {
                border-color: #1890ff;
                color: #1890ff;
            }
        """)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

    def get_values(self):
        try:
            zonghe = float(self.input_zonghe.text()) if self.input_zonghe.text().strip() else 0
            shuju = float(self.input_shuju.text()) if self.input_shuju.text().strip() else 0
            chanjiao = float(self.input_chanjiao.text()) if self.input_chanjiao.text().strip() else 0
            return zonghe, shuju, chanjiao
        except ValueError:
            return None


class DropButton(QPushButton):
    file_dropped = pyqtSignal(str)

    def __init__(self, text='', parent=None):
        super().__init__(text, parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(('.xlsx', '.xls')):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(('.xlsx', '.xls')):
                self.file_dropped.emit(path)
                break

# ==================== 订单报表处理逻辑 ====================
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


def process_order(df, week_start=None, week_end=None,
                  df_mgmt=None, df_oem=None, df_pwd_oem=None, df_manual=None, year25_data=None):
    """订单报表处理函数"""
    # 列名别名映射
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

    # 提取数据
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

    # 处理额外数据源
    pdt1, ind1, off1 = process_extra_file(df_mgmt, '数据安全管理中心')
    pdt2, ind2, off2 = process_extra_file(df_oem, '数据安全OEM IN')
    pdt3, ind3, off3 = process_extra_file(df_pwd_oem, '密码安全OEM IN')

    # 构建PDT报表
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
        '综合服务': ['安全驻场服务', '云端运营服务', '安全攻防服务', '安全咨询服务', '密码安全服务', '安全服务工具', '海外安全服务'],
        '数据安全': ['数据安全管理中心', '数据安全OEM IN', '数据安全产品转售', '数据安全服务'],
        '产教融合': ['产教融合产品转售', '产教融合服务'],
        '安全集成服务': ['安全集成服务'],
        '密码安全': ['密码安全OEM IN', '密码转售']
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

    # 融合额外数据源到 PDT pdt_data
    for src in (pdt1, pdt2, pdt3):
        for key, val in src.items():
            pdt_data[key] = pdt_data.get(key, 0) + val

    # 构建PDT报表
    pdt_rows = []
    integration_rows_data = []
    pwd_rows_data = []
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
            row['2026订单目标'] = '-'
            row['完成率'] = '-'
            if cat == '密码安全':
                row['本周新增'] = 0
            cat_rows.append(row)
        if cat in targets:
            target = targets[cat]
            cat_total = sum(r['合计'] for r in cat_rows)
            cat_rows[0]['2026订单目标'] = target
            cat_rows[0]['完成率'] = f'{round(cat_total / target * 100, 2)}%'
        if cat == '安全集成服务':
            integration_rows_data = cat_rows
        elif cat == '密码安全':
            pwd_rows_data = cat_rows
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
    data_rows = [r for r in pdt_rows if r['三级分类'] and not r['三级分类'].endswith('小计') and r['二级分类'] not in ('安全集成服务', '密码安全')]
    total_row = {'二级分类': '服务产品经理业绩总计', '三级分类': ''}
    for col in months_list + ['Q1小计', 'Q2小计', 'Q3小计', 'Q4小计', '合计', '本周新增']:
        total_row[col] = sum(r[col] for r in data_rows)
    total_row['2026订单目标'] = 30000
    total_val = total_row['合计']
    total_row['完成率'] = f'{round(total_val / 30000 * 100, 2)}%' if total_val else '0.00%'
    pdt_rows.append(total_row)
    pdt_rows.extend(integration_rows_data)
    pdt_rows.extend(pwd_rows_data)
    column_order = ['二级分类', '三级分类'] + \
                   ['1月', '2月', '3月', 'Q1小计'] + \
                   ['4月', '5月', '6月', 'Q2小计'] + \
                   ['7月', '8月', '9月', 'Q3小计'] + \
                   ['10月', '11月', '12月', 'Q4小计', '合计', '本周新增', '2026订单目标', '完成率', '25年合计', '25年同比']

    # df_pdt 将在最终统一四舍五入后创建

    # 行业报表
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

    numeric_cols = ['安全协维','安全运营','攻防服务','专家服务','密码安全服务','安全产品转售',
                    '海外服务','数据安全管理中心','数据安全OEM IN','数据安全产品转售','数据安全服务',
                    '产教融合产品转售','产教融合服务','服务小计','安全集成服务','密码安全OEM IN','密码转售']
    sub_keys = ['安全协维','安全运营','攻防服务','专家服务','密码安全服务','安全产品转售','海外服务',
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
        raw['密码安全服务'] = 0.0
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
        raw['密码安全OEM IN'] = sum(val for key, val in ind3.get(ind, {}).items()
                                if key == '密码安全OEM IN')
        raw['密码转售'] = 0.0
        industry_raws.append(raw)
        row = {'事业部': dept, '行业': ind}
        for col in numeric_cols:
            row[col] = raw[col]
        industry_rows.append(row)

    total_row = {'事业部': '', '行业': '总计'}
    for col in numeric_cols:
        if col in ('密码安全OEM IN', '密码转售'):
            total_row[col] = 0
        else:
            total_row[col] = sum(r[col] for r in industry_raws)
    industry_rows.append(total_row)

    df_industry_calc = pd.DataFrame(industry_rows)

    # 新建行业报表（合并版）
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
        if col in ('密码安全OEM IN', '密码转售'):
            total_row_new[col] = 0
        else:
            total_row_new[col] = sum(r[col] for r in new_industry_raws)
    new_industry_rows.append(total_row_new)

    df_industry_new = pd.DataFrame(new_industry_rows)

    # 办事处报表
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

    office_numeric_cols = ['安全协维','安全运营','攻防服务','专家服务','密码安全服务','安全产品转售',
                           '海外服务','数据安全管理中心','数据安全OEM IN','数据安全产品转售','数据安全服务',
                           '产教融合产品转售','产教融合服务','服务小计','安全集成服务','密码安全OEM IN','密码转售']

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
        raw['密码安全服务'] = 0.0
        raw['安全产品转售'] = sumif(bom_list=bom_tool, **skw)
        raw['海外服务'] = sumif(bom_list=bom_overseas, **skw)
        if is_hq:
            raw['数据安全管理中心'] = off1.get(off, {}).get('数据安全管理中心', 0)
            raw['数据安全OEM IN'] = off2.get(off, {}).get('数据安全OEM IN', 0)
            pwd_oem_val = off3.get(off, {}).get('密码安全OEM IN', 0)
            raw['密码安全OEM IN'] = pwd_oem_val
        else:
            region_short = region.rstrip('省市区')
            mgmt_val = 0.0
            oem_val = 0.0
            pwd_val = 0.0
            for try_key in (off, region, region_short):
                mgmt_val = off1.get(try_key, {}).get('数据安全管理中心', 0)
                if mgmt_val != 0:
                    break
            for try_key in (off, region, region_short):
                oem_val = off2.get(try_key, {}).get('数据安全OEM IN', 0)
                if oem_val != 0:
                    break
            for try_key in (off, region, region_short):
                pwd_val = off3.get(try_key, {}).get('密码安全OEM IN', 0)
                if pwd_val != 0:
                    break
            raw['数据安全管理中心'] = mgmt_val
            raw['数据安全OEM IN'] = oem_val
            raw['密码安全OEM IN'] = pwd_val
        raw['数据安全产品转售'] = 0.0
        raw['数据安全服务'] = sumif(bom_list=bom_data_security, **skw)
        raw['产教融合产品转售'] = 0.0
        raw['产教融合服务'] = sumif(bom_list=bom_education, **skw)
        raw['服务小计'] = sum(raw[k] for k in sub_keys)
        raw['安全集成服务'] = sumif(bom_list=bom_integration, **skw)
        raw['密码转售'] = 0.0
        office_raws.append(raw)
        row = {'区域': region, '办事处': off}
        for col in office_numeric_cols:
            row[col] = raw[col]
        office_rows.append(row)

    total_row_office = {'区域': '', '办事处': '总计'}
    for col in office_numeric_cols:
        if col in ('密码安全OEM IN', '密码转售'):
            total_row_office[col] = 0
        else:
            total_row_office[col] = sum(r[col] for r in office_raws)
    office_rows.append(total_row_office)

    df_office_calc = pd.DataFrame(office_rows)

    # 新建办事处报表（大区域版）
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
        if col in ('密码安全OEM IN', '密码转售'):
            total_row_new_off[col] = 0
        else:
            total_row_new_off[col] = sum(r[col] for r in new_off_raws)
    new_off_rows.append(total_row_new_off)

    df_office_new = pd.DataFrame(new_off_rows)

    # 应用手工计入数据 — 按位置一一对应加和
    if df_manual is not None:
        manual_keys = list(df_manual.keys())
        manual_pdt = df_manual[manual_keys[0]]
        manual_ind = df_manual[manual_keys[1]] if len(manual_keys) > 1 else None
        manual_off = df_manual[manual_keys[2]] if len(manual_keys) > 2 else None

        print(f'[手工数据] sheet数={len(manual_keys)}, sheets={manual_keys}')

        months = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']
        service_cols = ['安全协维','安全运营','攻防服务','专家服务','密码安全服务','安全产品转售',
                        '海外服务','数据安全管理中心','数据安全OEM IN','数据安全产品转售','数据安全服务',
                        '产教融合产品转售','产教融合服务','安全集成服务']
        sub_keys = ['安全协维','安全运营','攻防服务','专家服务','密码安全服务','安全产品转售','海外服务',
                    '数据安全管理中心','数据安全OEM IN','数据安全产品转售','数据安全服务',
                    '产教融合产品转售','产教融合服务']

        # --- PDT产品线：按位置对应，叶子行月份加和 ---
        if manual_pdt is not None:
            manual_months = [m for m in months if m in manual_pdt.columns]
            pdt_cat_idx = None
            for idx, c in enumerate(manual_pdt.columns):
                if '二级分类' in str(c):
                    pdt_cat_idx = idx
                    break
            pdt_total_added = 0
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
                        pdt_total_added += round(val)
            print(f'[手工-PDT] 加和总量={pdt_total_added}')

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
                    if cat in {'综合服务': 18000, '数据安全': 7000, '产教融合': 5000}:
                        target = {'综合服务': 18000, '数据安全': 7000, '产教融合': 5000}[cat]
                        cat_rows[0]['完成率'] = f'{round(sub_row["合计"] / target * 100, 2)}%'
            tgt_row = next((r for r in pdt_rows if r['二级分类']=='服务产品经理业绩总计'), None)
            if tgt_row:
                data_rows_t = [r for r in pdt_rows if r['三级分类'] and not (r['三级分类'] or '').endswith('小计') and r['二级分类'] not in ('安全集成服务', '密码安全')]
                for col in months + ['Q1小计','Q2小计','Q3小计','Q4小计','合计','本周新增']:
                    tgt_row[col] = sum(r.get(col,0) for r in data_rows_t)
                tgt_row['完成率'] = f'{round(tgt_row["合计"] / 30000 * 100, 2)}%' if tgt_row['合计'] else '0.00%'

        # --- 行业报表：按位置对应，非总计行服务列加和 ---
        if manual_ind is not None:
            ind_rows_list = list(df_industry_new.itertuples(index=False))
            ind_new_added = 0
            for i, (mrow, orow) in enumerate(zip(manual_ind.itertuples(index=False), ind_rows_list)):
                if orow.行业 == '总计':
                    continue
                row_added = 0
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
                            old_val = getattr(orow, col, 0) if hasattr(orow, col) else df_industry_new.at[i, col]
                            df_industry_new.at[i, col] = old_val + val
                            row_added += round(val)
                if row_added != 0:
                    df_industry_new.at[i, '服务小计'] = sum(
                        df_industry_new.at[i, c] for c in sub_keys)
                    ind_new_added += row_added
            total_mask = df_industry_new['行业'] == '总计'
            if total_mask.any():
                total_idx = df_industry_new[total_mask].index[0]
                for col in service_cols + ['服务小计']:
                    if col in df_industry_new.columns:
                        df_industry_new.at[total_idx, col] = df_industry_new[col].drop(total_idx).sum()
            print(f'[手工-行业] 报表加和={ind_new_added}')

        # --- 办事处报表：按位置对应，非总计行服务列加和 ---
        if manual_off is not None:
            off_new_rows = list(df_office_new.itertuples(index=False))
            off_new_added = 0
            for i, (mrow, orow) in enumerate(zip(manual_off.itertuples(index=False), off_new_rows)):
                if orow.省办 == '总计':
                    continue
                row_added = 0
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
                            old_val = getattr(orow, col, 0) if hasattr(orow, col) else df_office_new.at[i, col]
                            df_office_new.at[i, col] = old_val + val
                            row_added += round(val)
                if row_added != 0:
                    df_office_new.at[i, '服务小计'] = sum(
                        df_office_new.at[i, c] for c in sub_keys)
                    off_new_added += row_added
            total_mask = df_office_new['省办'] == '总计'
            if total_mask.any():
                total_idx = df_office_new[total_mask].index[0]
                for col in service_cols + ['服务小计']:
                    if col in df_office_new.columns:
                        df_office_new.at[total_idx, col] = df_office_new[col].drop(total_idx).sum()
            print(f'[手工-办事处] 报表加和={off_new_added}')

    if year25_data is not None:
        zonghe_25, shuju_25, chanjiao_25 = year25_data
        first_zonghe_row = None
        first_shuju_row = None
        first_chanjiao_row = None
        total_row = None
        zonghe_subtotal_row = None
        shuju_subtotal_row = None
        chanjiao_subtotal_row = None
        
        print(f'[DEBUG] year25_data: 综合={zonghe_25}, 数据={shuju_25}, 产教={chanjiao_25}')
        print(f'[DEBUG] pdt_rows数量: {len(pdt_rows)}')
        
        for row in pdt_rows:
            cat = row.get('二级分类', '')
            sub = row.get('三级分类', '')
            
            if cat == '综合服务' and first_zonghe_row is None:
                first_zonghe_row = row
            elif cat == '数据安全' and first_shuju_row is None:
                first_shuju_row = row
            elif cat == '产教融合' and first_chanjiao_row is None:
                first_chanjiao_row = row
            elif cat == '服务产品经理业绩总计':
                total_row = row
            elif sub == '综合服务小计':
                zonghe_subtotal_row = row
            elif sub == '数据安全小计':
                shuju_subtotal_row = row
            elif sub == '产教融合小计':
                chanjiao_subtotal_row = row
        
        print(f'[DEBUG] 第一个数据行: 综合={first_zonghe_row is not None}, 数据={first_shuju_row is not None}, 产教={first_chanjiao_row is not None}, 总计={total_row is not None}')
        print(f'[DEBUG] 小计行: 综合={zonghe_subtotal_row is not None}, 数据={shuju_subtotal_row is not None}, 产教={chanjiao_subtotal_row is not None}')
        
        if first_zonghe_row and first_shuju_row and first_chanjiao_row and total_row:
            first_zonghe_row['25年合计'] = zonghe_25
            first_shuju_row['25年合计'] = shuju_25
            first_chanjiao_row['25年合计'] = chanjiao_25
            total_row['25年合计'] = zonghe_25 + shuju_25 + chanjiao_25
            
            def calc_yoy(s_val, w_val):
                if w_val and w_val != 0:
                    return f'{round((s_val - w_val) / w_val * 100, 2)}%'
                return ''
            
            zonghe_total = zonghe_subtotal_row['合计'] if zonghe_subtotal_row else first_zonghe_row['合计']
            shuju_total = shuju_subtotal_row['合计'] if shuju_subtotal_row else first_shuju_row['合计']
            chanjiao_total = chanjiao_subtotal_row['合计'] if chanjiao_subtotal_row else first_chanjiao_row['合计']
            
            first_zonghe_row['25年同比'] = calc_yoy(zonghe_total, zonghe_25)
            first_shuju_row['25年同比'] = calc_yoy(shuju_total, shuju_25)
            first_chanjiao_row['25年同比'] = calc_yoy(chanjiao_total, chanjiao_25)
            total_row['25年同比'] = calc_yoy(total_row['合计'], total_row['25年合计'])
            
            print(f'[DEBUG] 数据设置完成: 综合26年合计={zonghe_total}, 25年合计={first_zonghe_row["25年合计"]}, 同比={first_zonghe_row["25年同比"]}')
            print(f'[DEBUG] 数据设置完成: 数据26年合计={shuju_total}, 25年合计={first_shuju_row["25年合计"]}, 同比={first_shuju_row["25年同比"]}')
            print(f'[DEBUG] 数据设置完成: 产教26年合计={chanjiao_total}, 25年合计={first_chanjiao_row["25年合计"]}, 同比={first_chanjiao_row["25年同比"]}')
    else:
        for row in pdt_rows:
            row['25年合计'] = ''
            row['25年同比'] = ''

    # ===== 统一四舍五入：所有数值保留到最后一步才取整 =====
    pdt_numeric = months_list + ['Q1小计','Q2小计','Q3小计','Q4小计','合计','本周新增']
    for row in pdt_rows:
        for col in pdt_numeric:
            if col in row and isinstance(row[col], (int, float)):
                row[col] = round(row[col])
    
    print('[DEBUG] 转换为DataFrame前的验证:')
    for idx, row in enumerate(pdt_rows):
        if row.get('25年合计') is not None:
            print(f'[DEBUG] Excel行{idx+2}: 二级={row.get("二级分类")}, 三级={row.get("三级分类")}, 25年合计={row.get("25年合计")}, 25年同比={row.get("25年同比")}')
    
    df_pdt = pd.DataFrame(pdt_rows, columns=column_order)
    print(f'[DEBUG] DataFrame列: {list(df_pdt.columns)}')
    for df_ref in [df_industry_calc, df_industry_new, df_office_calc, df_office_new]:
        for col in numeric_cols:
            if col in df_ref.columns:
                df_ref[col] = pd.to_numeric(df_ref[col], errors='coerce').fillna(0).round().astype(int)

    return {
        'PDT产品线': df_pdt,
        '行业计算表': df_industry_calc,
        '行业报表': df_industry_new,
        '办事处计算表': df_office_calc,
        '办事处报表': df_office_new
    }


# ==================== 订单模块界面 ====================
class OrderModule(BaseModule):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollArea > QWidget > QWidget {
                background-color: transparent;
            }
        """)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setAlignment(Qt.AlignTop)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        
        content_widget = QWidget()
        self.init_ui_content(content_widget)
        
        scroll.setWidget(content_widget)
        
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(scroll)
    
    def init_ui_content(self, content_widget):
        layout = QVBoxLayout(content_widget)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(14)

        self.current_df_mgmt = None
        self.current_df_oem = None
        self.current_df_pwd_oem = None
        self.current_df_manual = None
        self.year25_data = None
        self.current_df_maoli = None
        self.current_df_maoli_manual = None
        self.last_dir = os.path.expanduser("~")

        upload_card_style = (
            "QPushButton {"
            "  background-color: #ffffff;"
            "  border: 2px dashed #d9d9d9;"
            "  border-radius: 10px;"
            "  padding: 18px 24px;"
            "  font-size: 14px;"
            "  color: #666666;"
            "  text-align: left;"
            "}"
            "QPushButton:hover {"
            "  border-color: #1890ff;"
            "  background-color: #f0f8ff;"
            "  color: #1890ff;"
            "}"
        )
        upload_card_ok_style = (
            "QPushButton {"
            "  background-color: #f6ffed;"
            "  border: 2px solid #b7eb8f;"
            "  border-radius: 10px;"
            "  padding: 18px 24px;"
            "  font-size: 14px;"
            "  color: #389e0d;"
            "  text-align: left;"
            "}"
            "QPushButton:hover {"
            "  border-color: #52c41a;"
            "  background-color: #f0fff0;"
            "}"
        )

        # ---- 上传区域 ----
        section_label_1 = QLabel("📥  数据源文件")
        section_label_1.setStyleSheet("font-size: 15px; font-weight: bold; color: #333333; margin-top: 4px;")
        layout.addWidget(section_label_1)

        self.upload_btn_1 = DropButton("   📂  请上传服务订单报表（必选）\n       点击选择 Excel 文件")
        self.upload_btn_1.setMinimumHeight(72)
        self.upload_btn_1.setCursor(Qt.PointingHandCursor)
        self.upload_btn_1.setStyleSheet(upload_card_style)
        self.upload_btn_1.clicked.connect(self.select_main_file)
        self.upload_btn_1.file_dropped.connect(self._on_drop_file)
        layout.addWidget(self.upload_btn_1)

        self.upload_btn_2 = DropButton("   📂  请上传数据安全管理中心报表（可选）\n       点击选择 Excel 文件")
        self.upload_btn_2.setMinimumHeight(72)
        self.upload_btn_2.setCursor(Qt.PointingHandCursor)
        self.upload_btn_2.setStyleSheet(upload_card_style)
        self.upload_btn_2.clicked.connect(self.select_mgmt_file)
        self.upload_btn_2.file_dropped.connect(lambda p: self._on_drop_file(p, 2))
        layout.addWidget(self.upload_btn_2)

        self.upload_btn_3 = DropButton("   📂  请上传数据安全OEM IN产品报表（可选）\n       点击选择 Excel 文件")
        self.upload_btn_3.setMinimumHeight(72)
        self.upload_btn_3.setCursor(Qt.PointingHandCursor)
        self.upload_btn_3.setStyleSheet(upload_card_style)
        self.upload_btn_3.clicked.connect(self.select_oem_file)
        self.upload_btn_3.file_dropped.connect(lambda p: self._on_drop_file(p, 3))
        layout.addWidget(self.upload_btn_3)

        self.upload_btn_pwd_oem = DropButton("   📂  请上传密码安全OEM IN产品报表（可选）\n       点击选择 Excel 文件")
        self.upload_btn_pwd_oem.setMinimumHeight(72)
        self.upload_btn_pwd_oem.setCursor(Qt.PointingHandCursor)
        self.upload_btn_pwd_oem.setStyleSheet(upload_card_style)
        self.upload_btn_pwd_oem.clicked.connect(self.select_pwd_oem_file)
        self.upload_btn_pwd_oem.file_dropped.connect(lambda p: self._on_drop_file(p, 7))
        layout.addWidget(self.upload_btn_pwd_oem)

        self.upload_btn_4 = DropButton("   📂  请导入手工计入表格（可选）\n       点击选择 Excel 文件")
        self.upload_btn_4.setMinimumHeight(72)
        self.upload_btn_4.setCursor(Qt.PointingHandCursor)
        self.upload_btn_4.setStyleSheet(upload_card_style)
        self.upload_btn_4.clicked.connect(self.select_manual_file)
        self.upload_btn_4.file_dropped.connect(lambda p: self._on_drop_file(p, 4))
        layout.addWidget(self.upload_btn_4)

        # ---- 毛利表数据源区域 ----
        section_label_3 = QLabel("📥  毛利表数据源")
        section_label_3.setStyleSheet("font-size: 15px; font-weight: bold; color: #333333; margin-top: 6px;")
        layout.addWidget(section_label_3)

        self.upload_btn_5 = DropButton("   📂  请上传订单毛利表（可选）\n       点击选择 Excel 文件")
        self.upload_btn_5.setMinimumHeight(72)
        self.upload_btn_5.setCursor(Qt.PointingHandCursor)
        self.upload_btn_5.setStyleSheet(upload_card_style)
        self.upload_btn_5.clicked.connect(self.select_maoli_file)
        self.upload_btn_5.file_dropped.connect(lambda p: self._on_drop_file(p, 5))
        layout.addWidget(self.upload_btn_5)

        self.upload_btn_6 = DropButton("   📂  请上传订单毛利手工表（可选）\n       点击选择 Excel 文件")
        self.upload_btn_6.setMinimumHeight(72)
        self.upload_btn_6.setCursor(Qt.PointingHandCursor)
        self.upload_btn_6.setStyleSheet(upload_card_style)
        self.upload_btn_6.clicked.connect(self.select_maoli_manual_file)
        self.upload_btn_6.file_dropped.connect(lambda p: self._on_drop_file(p, 6))
        layout.addWidget(self.upload_btn_6)

        self.upload_card_style = upload_card_style
        self.upload_card_ok_style = upload_card_ok_style

        # ---- 25年合计数据输入区域 ----
        year25_card = QFrame()
        year25_card.setStyleSheet("QFrame { background-color: #ffffff; border: 1px solid #e8e8e8; border-radius: 8px; }")
        year25_layout = QVBoxLayout(year25_card)
        year25_layout.setContentsMargins(16, 14, 16, 14)
        year25_layout.setSpacing(14)

        year25_label = QLabel("📊  25年合计数据（单位：万元）")
        year25_label.setStyleSheet("font-size: 15px; font-weight: bold; color: #333333;")
        year25_layout.addWidget(year25_label)

        input_row1 = QHBoxLayout()
        input_row1.addWidget(QLabel("综合服务："))
        self.input_zonghe = QLineEdit()
        self.input_zonghe.setPlaceholderText("请输入综合服务25年合计")
        self.input_zonghe.setStyleSheet("border: 1px solid #d9d9d9; border-radius: 4px; padding: 8px;")
        input_row1.addWidget(self.input_zonghe)
        year25_layout.addLayout(input_row1)

        input_row2 = QHBoxLayout()
        input_row2.addWidget(QLabel("数据安全："))
        self.input_shuju = QLineEdit()
        self.input_shuju.setPlaceholderText("请输入数据安全25年合计")
        self.input_shuju.setStyleSheet("border: 1px solid #d9d9d9; border-radius: 4px; padding: 8px;")
        input_row2.addWidget(self.input_shuju)
        year25_layout.addLayout(input_row2)

        input_row3 = QHBoxLayout()
        input_row3.addWidget(QLabel("产教融合："))
        self.input_chanjiao = QLineEdit()
        self.input_chanjiao.setPlaceholderText("请输入产教融合25年合计")
        self.input_chanjiao.setStyleSheet("border: 1px solid #d9d9d9; border-radius: 4px; padding: 8px;")
        input_row3.addWidget(self.input_chanjiao)
        year25_layout.addLayout(input_row3)

        layout.addWidget(year25_card)

        # ---- 数据预览 ----
        section_label_2 = QLabel("📊  数据预览")
        section_label_2.setStyleSheet("font-size: 15px; font-weight: bold; color: #333333; margin-top: 6px;")
        layout.addWidget(section_label_2)

        preview_frame = QFrame()
        preview_frame.setStyleSheet("QFrame { background-color: #ffffff; border: 1px solid #e8e8e8; border-radius: 8px; }")
        preview_layout = QVBoxLayout(preview_frame)
        preview_layout.setContentsMargins(6, 6, 6, 6)
        self.preview_table = QTableWidget()
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.preview_table.setMinimumHeight(180)
        preview_layout.addWidget(self.preview_table)
        layout.addWidget(preview_frame)

        # ---- 日期范围 + 操作按钮 ----
        bottom_card = QFrame()
        bottom_card.setStyleSheet("QFrame { background-color: #ffffff; border: 1px solid #e8e8e8; border-radius: 8px; }")
        bottom_layout = QVBoxLayout(bottom_card)
        bottom_layout.setContentsMargins(16, 14, 16, 14)
        bottom_layout.setSpacing(14)

        week_label = QLabel("📅  本周新增时间范围")
        week_label.setStyleSheet("font-size: 15px; font-weight: bold; color: #333333;")
        bottom_layout.addWidget(week_label)

        week_layout = QHBoxLayout()
        week_layout.setSpacing(10)
        self.week_start_edit = QDateEdit()
        self.week_start_edit.setCalendarPopup(True)
        self.week_start_edit.setDisplayFormat("yyyy-MM-dd")
        self.week_end_edit = QDateEdit()
        self.week_end_edit.setCalendarPopup(True)
        self.week_end_edit.setDisplayFormat("yyyy-MM-dd")
        today = QDate.currentDate()
        dow = today.dayOfWeek()
        monday = today.addDays(-(dow - 1)) if dow <= 6 else today.addDays(-6)
        tuesday = monday.addDays(-6)
        self.week_start_edit.setDate(tuesday)
        self.week_end_edit.setDate(monday)
        week_layout.addWidget(QLabel("  从"))
        week_layout.addWidget(self.week_start_edit)
        week_layout.addWidget(QLabel("至"))
        week_layout.addWidget(self.week_end_edit)
        week_layout.addStretch()
        bottom_layout.addLayout(week_layout)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self.generate_btn = QPushButton("  生成报表")
        self.generate_btn.setStyleSheet(
            "QPushButton { background-color: #1890ff; color: #ffffff; border: none;"
            "  border-radius: 6px; padding: 10px 28px; font-size: 15px; font-weight: bold; }"
            "QPushButton:hover { background-color: #40a9ff; }"
            "QPushButton:pressed { background-color: #096dd9; }"
            "QPushButton:disabled { background-color: #d9d9d9; color: #ffffff; }"
        )
        self.generate_btn.setCursor(Qt.PointingHandCursor)

        self.clear_btn = QPushButton("  清除全部上传")
        self.clear_btn.setStyleSheet(
            "QPushButton { background-color: #ffffff; color: #ff4d4f; border: 1px solid #ff4d4f;"
            "  border-radius: 6px; padding: 10px 28px; font-size: 15px; }"
            "QPushButton:hover { color: #ffffff; background-color: #ff4d4f; }"
        )
        self.clear_btn.setCursor(Qt.PointingHandCursor)

        self.download_btn = QPushButton("  下载报表")
        self.download_btn.setEnabled(False)
        self.download_btn.setStyleSheet(
            "QPushButton { background-color: #ffffff; color: #52c41a; border: 1px solid #52c41a;"
            "  border-radius: 6px; padding: 10px 28px; font-size: 15px; }"
            "QPushButton:hover { color: #ffffff; background-color: #52c41a; }"
            "QPushButton:disabled { background-color: #d9d9d9; color: #ffffff; border: none; }"
        )
        self.download_btn.setCursor(Qt.PointingHandCursor)

        self.download_full_btn = QPushButton("  下载完整报表")
        self.download_full_btn.setEnabled(False)
        self.download_full_btn.setStyleSheet(
            "QPushButton { background-color: #1890ff; color: #ffffff; border: none;"
            "  border-radius: 6px; padding: 10px 28px; font-size: 15px; font-weight: bold; }"
            "QPushButton:hover { background-color: #40a9ff; }"
            "QPushButton:disabled { background-color: #d9d9d9; color: #ffffff; border: none; }"
        )
        self.download_full_btn.setCursor(Qt.PointingHandCursor)

        btn_layout.addWidget(self.generate_btn)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(self.download_btn)
        btn_layout.addWidget(self.download_full_btn)
        btn_layout.addStretch()
        bottom_layout.addLayout(btn_layout)

        layout.addWidget(bottom_card)
        layout.addStretch()

        self.generate_btn.clicked.connect(self.start_generate)
        self.clear_btn.clicked.connect(self.clear_all_uploads)
        self.download_btn.clicked.connect(self.save_report)
        self.download_full_btn.clicked.connect(self.save_report_full)

        self.report_thread = None

    def on_file_loaded(self, df):
        if '币种' in df.columns and '文本总金额' in df.columns:
            currency_rates = {'人民币': 1.0, 'cny': 1.0, 'usd': 7.0187, 'hkd': 0.90131}
            raw_currencies = df['币种'].astype(str).str.strip().str.lower()
            rates = raw_currencies.map(currency_rates).fillna(1.0)
            df['换算后金额(CNY)'] = pd.to_numeric(df['文本总金额'], errors='coerce').fillna(0) * rates
        self.preview_data(df.head(100))
        self.download_btn.setEnabled(False)
        self.download_full_btn.setEnabled(False)
        self.current_report_path = None

    def preview_data(self, df):
        self.preview_table.clear()
        if df.empty:
            return
        self.preview_table.setRowCount(df.shape[0])
        self.preview_table.setColumnCount(df.shape[1])
        self.preview_table.setHorizontalHeaderLabels(df.columns.astype(str))
        for i, row in df.iterrows():
            for j, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                self.preview_table.setItem(i, j, item)
        self.preview_table.resizeColumnsToContents()

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", self.last_dir, "Excel文件 (*.xlsx *.xls)",
            options=QFileDialog.DontUseNativeDialog
        )
        if file_path:
            self.last_dir = os.path.dirname(file_path)
            self.load_file(file_path)

    def select_main_file(self):
        self.select_file()

    def _on_drop_file(self, path, slot=1):
        self.last_dir = os.path.dirname(path)
        if slot == 1:
            self.load_file(path)
        elif slot == 2:
            self.load_mgmt_file(path)
        elif slot == 3:
            self.load_oem_file(path)
        elif slot == 4:
            self.load_manual_file(path)
        elif slot == 5:
            self.load_maoli_file(path)
        elif slot == 6:
            self.load_maoli_manual_file(path)
        elif slot == 7:
            self.load_pwd_oem_file(path)

    def load_file(self, file_path):
        try:
            df = pd.read_excel(file_path, sheet_name=0, engine='openpyxl')
            self.current_df = df
            self.main_file_path = file_path
            fname = os.path.basename(file_path)
            self.upload_btn_1.setText(f"   ✅  服务订单报表已加载：{fname}\n       点击可重新选择")
            self.upload_btn_1.setStyleSheet(self.upload_card_ok_style)
            self.on_file_loaded(df)
        except Exception as e:
            QMessageBox.critical(self, "读取失败", str(e))

    def clear_all_uploads(self):
        self.current_df = None
        self.main_file_path = None
        self.current_df_mgmt = None
        self.current_df_oem = None
        self.current_df_pwd_oem = None
        self.current_df_manual = None
        self.year25_data = None
        self.current_df_maoli = None
        self.current_df_maoli_manual = None
        self.preview_table.clear()
        self.preview_table.setRowCount(0)
        self.preview_table.setColumnCount(0)
        self.download_btn.setEnabled(False)
        self.download_full_btn.setEnabled(False)
        self.current_report_path = None
        self.upload_btn_1.setText("   📂  请上传服务订单报表（必选）\n       点击选择 Excel 文件")
        self.upload_btn_1.setStyleSheet(self.upload_card_style)
        self.upload_btn_2.setText("   📂  请上传数据安全管理中心报表（可选）\n       点击选择 Excel 文件")
        self.upload_btn_2.setStyleSheet(self.upload_card_style)
        self.upload_btn_3.setText("   📂  请上传数据安全OEM IN产品报表（可选）\n       点击选择 Excel 文件")
        self.upload_btn_3.setStyleSheet(self.upload_card_style)
        self.upload_btn_4.setText("   📂  请导入手工计入表格（可选）\n       点击选择 Excel 文件")
        self.upload_btn_4.setStyleSheet(self.upload_card_style)
        self.upload_btn_5.setText("   📂  请上传订单毛利表（可选）\n       点击选择 Excel 文件")
        self.upload_btn_5.setStyleSheet(self.upload_card_style)
        self.upload_btn_6.setText("   📂  请上传订单毛利手工表（可选）\n       点击选择 Excel 文件")
        self.upload_btn_6.setStyleSheet(self.upload_card_style)
        self.upload_btn_pwd_oem.setText("   📂  请上传密码安全OEM IN产品报表（可选）\n       点击选择 Excel 文件")
        self.upload_btn_pwd_oem.setStyleSheet(self.upload_card_style)
        self.input_zonghe.clear()
        self.input_shuju.clear()
        self.input_chanjiao.clear()

    def select_mgmt_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择数据安全管理中心报表", self.last_dir,
            "Excel文件 (*.xlsx *.xls)", options=QFileDialog.DontUseNativeDialog
        )
        if file_path:
            self.last_dir = os.path.dirname(file_path)
            self.load_mgmt_file(file_path)

    def select_oem_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择数据安全OEM IN产品报表", self.last_dir,
            "Excel文件 (*.xlsx *.xls)", options=QFileDialog.DontUseNativeDialog
        )
        if file_path:
            self.last_dir = os.path.dirname(file_path)
            self.load_oem_file(file_path)

    def select_pwd_oem_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择密码安全OEM IN产品报表", self.last_dir,
            "Excel文件 (*.xlsx *.xls)", options=QFileDialog.DontUseNativeDialog
        )
        if file_path:
            self.last_dir = os.path.dirname(file_path)
            self.load_pwd_oem_file(file_path)

    def select_manual_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择手工计入表格", self.last_dir,
            "Excel文件 (*.xlsx *.xls)", options=QFileDialog.DontUseNativeDialog
        )
        if file_path:
            self.last_dir = os.path.dirname(file_path)
            self.load_manual_file(file_path)

    def select_maoli_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择订单毛利表", self.last_dir,
            "Excel文件 (*.xlsx *.xls)", options=QFileDialog.DontUseNativeDialog
        )
        if file_path:
            self.last_dir = os.path.dirname(file_path)
            self.load_maoli_file(file_path)

    def select_maoli_manual_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择订单毛利手工表", self.last_dir,
            "Excel文件 (*.xlsx *.xls)", options=QFileDialog.DontUseNativeDialog
        )
        if file_path:
            self.last_dir = os.path.dirname(file_path)
            self.load_maoli_manual_file(file_path)

    def _validate_extra_columns(self, df, file_type):
        required = ['年月', '系统部', '营业收入']
        missing = [c for c in required if c not in df.columns]
        if missing:
            QMessageBox.warning(self, "列名不匹配",
                f"{file_type} 缺少必要列：{', '.join(missing)}\n\n"
                f"当前列名：{list(df.columns)}\n\n请确认表格格式是否正确（需包含：年月、系统部、营业收入等列）")

    def load_mgmt_file(self, file_path):
        try:
            df = self._read_extra_file(file_path, "数据安全管理中心报表")
            self.current_df_mgmt = df
            fname = os.path.basename(file_path)
            self.upload_btn_2.setText(f"   ✅  数据安全管理中心报表已加载：{fname}\n       点击可重新选择")
            self.upload_btn_2.setStyleSheet(self.upload_card_ok_style)
            self.download_btn.setEnabled(False)
            self.current_report_path = None
        except Exception as e:
            QMessageBox.critical(self, "读取失败", str(e))

    def load_oem_file(self, file_path):
        try:
            df = self._read_extra_file(file_path, "数据安全OEM IN产品报表")
            self.current_df_oem = df
            fname = os.path.basename(file_path)
            self.upload_btn_3.setText(f"   ✅  数据安全OEM IN产品报表已加载：{fname}\n       点击可重新选择")
            self.upload_btn_3.setStyleSheet(self.upload_card_ok_style)
            self.download_btn.setEnabled(False)
            self.current_report_path = None
        except Exception as e:
            QMessageBox.critical(self, "读取失败", str(e))

    def load_pwd_oem_file(self, file_path):
        try:
            df = self._read_extra_file(file_path, "密码安全OEM IN产品报表")
            self.current_df_pwd_oem = df
            fname = os.path.basename(file_path)
            self.upload_btn_pwd_oem.setText(f"   ✅  密码安全OEM IN产品报表已加载：{fname}\n       点击可重新选择")
            self.upload_btn_pwd_oem.setStyleSheet(self.upload_card_ok_style)
            self.download_btn.setEnabled(False)
            self.current_report_path = None
        except Exception as e:
            QMessageBox.critical(self, "读取失败", str(e))

    def _read_extra_file(self, file_path, file_type):
        required = ['年月', '系统部', '营业收入']
        for header_row in [1, 0, 2, 3]:
            df = pd.read_excel(file_path, sheet_name=0, header=header_row, engine='openpyxl')
            if all(c in df.columns for c in required):
                return df
        df = pd.read_excel(file_path, sheet_name=0, header=1, engine='openpyxl')
        self._validate_extra_columns(df, file_type)
        return df

    def load_manual_file(self, file_path):
        try:
            sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            if len(sheets) < 3:
                QMessageBox.warning(self, "格式错误", "手工计入表格需要包含3个sheet：PDT产品线、行业报表、办事处报表")
                return
            self.current_df_manual = sheets
            fname = os.path.basename(file_path)
            sheet_names = list(sheets.keys())
            self.upload_btn_4.setText(f"   ✅  手工计入表格已加载：{fname}\n       （含 {len(sheet_names)} 个sheet，点击可重新选择）")
            self.upload_btn_4.setStyleSheet(self.upload_card_ok_style)
            self.download_btn.setEnabled(False)
            self.current_report_path = None
        except Exception as e:
            QMessageBox.critical(self, "读取失败", str(e))

    def load_maoli_file(self, file_path):
        try:
            df = pd.read_excel(file_path, sheet_name=0, engine='openpyxl')
            self.current_df_maoli = df
            fname = os.path.basename(file_path)
            self.upload_btn_5.setText(f"   ✅  订单毛利表已加载：{fname}\n       点击可重新选择")
            self.upload_btn_5.setStyleSheet(self.upload_card_ok_style)
            self.download_btn.setEnabled(False)
            self.current_report_path = None
        except Exception as e:
            QMessageBox.critical(self, "读取失败", str(e))

    def load_maoli_manual_file(self, file_path):
        try:
            sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            if len(sheets) < 3:
                QMessageBox.warning(self, "格式错误", "订单毛利手工表需要包含3个sheet：毛利_PDT产品线、毛利_行业报表、毛利_办事处报表")
                return
            self.current_df_maoli_manual = sheets
            fname = os.path.basename(file_path)
            sheet_names = list(sheets.keys())
            self.upload_btn_6.setText(f"   ✅  订单毛利手工表已加载：{fname}\n       （含 {len(sheet_names)} 个sheet，点击可重新选择）")
            self.upload_btn_6.setStyleSheet(self.upload_card_ok_style)
            self.download_btn.setEnabled(False)
            self.current_report_path = None
        except Exception as e:
            QMessageBox.critical(self, "读取失败", str(e))

    def start_generate(self):
        if self.current_df is None and self.current_df_maoli is None:
            QMessageBox.warning(self, "无数据", "请至少上传服务订单报表或订单毛利表之一")
            return

        try:
            zonghe_text = self.input_zonghe.text().strip()
            shuju_text = self.input_shuju.text().strip()
            chanjiao_text = self.input_chanjiao.text().strip()
            
            if zonghe_text or shuju_text or chanjiao_text:
                try:
                    zonghe_25 = float(zonghe_text) if zonghe_text else 0
                    shuju_25 = float(shuju_text) if shuju_text else 0
                    chanjiao_25 = float(chanjiao_text) if chanjiao_text else 0
                    self.year25_data = (zonghe_25, shuju_25, chanjiao_25)
                except ValueError:
                    QMessageBox.warning(self, "输入错误", "请输入有效的数字")
                    return
            else:
                self.year25_data = None

            self.generate_btn.setEnabled(False)
            self.download_btn.setEnabled(False)
            self.download_full_btn.setEnabled(False)

            week_start = self.week_start_edit.date().toPyDate()
            week_end = self.week_end_edit.date().toPyDate()
            
            def process_both():
                result = {}
                
                if self.current_df is not None:
                    result = process_order(self.current_df,
                                          week_start=week_start, week_end=week_end,
                                          df_mgmt=self.current_df_mgmt,
                                          df_oem=self.current_df_oem,
                                          df_pwd_oem=self.current_df_pwd_oem,
                                          df_manual=self.current_df_manual,
                                          year25_data=self.year25_data)

                if self.current_df_maoli is not None:
                    maoli_manual = None
                    if self.current_df_maoli_manual is not None:
                        maoli_manual = self.current_df_maoli_manual
                    result2 = process_maoli_order(self.current_df_maoli,
                                                  week_start=week_start, week_end=week_end,
                                                  df_mgmt=None, df_oem=None,
                                                  df_pwd_oem=self.current_df_pwd_oem,
                                                  df_manual=maoli_manual)
                    result.update(result2)
                
                return result
            
            self.report_thread = ReportThread(process_both)
            self.report_thread.finished.connect(self.on_generate_finished)
            self.report_thread.error.connect(self.on_generate_error)
            self.report_thread.start()

            self.progress = QProgressDialog("正在生成报表，请稍候...", None, 0, 0, self)
            self.progress.setWindowModality(Qt.WindowModal)
            self.progress.show()
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "错误", f"生成报表时发生错误：\n{str(e)}\n\n详细信息：\n{traceback.format_exc()}")

    def on_generate_finished(self, result):
        self.progress.close()

        order_totals = {}
        maoli_totals = {}
        
        if 'PDT产品线' in result and result['PDT产品线'] is not None:
            df_pdt = result['PDT产品线']
            pdt_total_mask = df_pdt['二级分类'] == '服务产品经理业绩总计'
            pdt_total = int(df_pdt[pdt_total_mask]['合计'].values[0]) if pdt_total_mask.any() else 0
            order_totals['PDT产品线'] = pdt_total

        if '行业报表' in result and result['行业报表'] is not None:
            df_ind = result['行业报表']
            ind_total_mask = df_ind['行业'] == '总计'
            ind_total = int(df_ind[ind_total_mask]['服务小计'].values[0]) if ind_total_mask.any() else 0
            order_totals['行业报表'] = ind_total

        if '办事处报表' in result and result['办事处报表'] is not None:
            df_off = result['办事处报表']
            off_total_mask = df_off['省办'] == '总计'
            off_total = int(df_off[off_total_mask]['服务小计'].values[0]) if off_total_mask.any() else 0
            order_totals['办事处报表'] = off_total

        if '毛利_PDT产品线' in result and result['毛利_PDT产品线'] is not None:
            df_maoli_pdt = result['毛利_PDT产品线']
            maoli_pdt_mask = df_maoli_pdt['二级分类'] == '服务产品经理业绩总计'
            maoli_pdt_total = int(df_maoli_pdt[maoli_pdt_mask]['合计'].values[0]) if maoli_pdt_mask.any() else 0
            maoli_totals['毛利_PDT产品线'] = maoli_pdt_total

        if '毛利_行业报表' in result and result['毛利_行业报表'] is not None:
            df_maoli_ind = result['毛利_行业报表']
            maoli_ind_mask = df_maoli_ind['行业'] == '总计'
            maoli_ind_total = int(df_maoli_ind[maoli_ind_mask]['服务小计'].values[0]) if maoli_ind_mask.any() else 0
            maoli_totals['毛利_行业报表'] = maoli_ind_total

        if '毛利_办事处报表' in result and result['毛利_办事处报表'] is not None:
            df_maoli_off = result['毛利_办事处报表']
            maoli_off_mask = df_maoli_off['省办'] == '总计'
            maoli_off_total = int(df_maoli_off[maoli_off_mask]['服务小计'].values[0]) if maoli_off_mask.any() else 0
            maoli_totals['毛利_办事处报表'] = maoli_off_total

        validation_parts = []
        
        if order_totals:
            if len(order_totals) > 1:
                vals = list(order_totals.values())
                order_match = all(abs(v - vals[0]) <= 2 for v in vals)
                detail_order = '\n    '.join(f'{k}: {int(v):,} 万元' for k, v in order_totals.items())
                if order_match:
                    validation_parts.append(f'✅ 服务订单部分校验通过\n    {detail_order}')
                else:
                    validation_parts.append(f'⚠️ 服务订单部分校验不一致！\n    {detail_order}')
            else:
                k, v = list(order_totals.items())[0]
                validation_parts.append(f'服务订单部分\n    {k}: {int(v):,} 万元')

        if maoli_totals:
            if len(maoli_totals) > 1:
                vals = list(maoli_totals.values())
                maoli_match = all(abs(v - vals[0]) <= 2 for v in vals)
                detail_maoli = '\n    '.join(f'{k}: {int(v):,} 万元' for k, v in maoli_totals.items())
                if maoli_match:
                    validation_parts.append(f'✅ 毛利部分校验通过\n    {detail_maoli}')
                else:
                    validation_parts.append(f'⚠️ 毛利部分校验不一致！\n    {detail_maoli}')
            else:
                k, v = list(maoli_totals.items())[0]
                validation_parts.append(f'毛利部分\n    {k}: {int(v):,} 万元')

        self._validation_msg = '\n\n'.join(validation_parts) if validation_parts else '报表已生成'

        fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
            for sheet_name, df_sheet in result.items():
                if df_sheet is not None and not df_sheet.empty:
                    df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
            apply_excel_style(writer.book)
        self.current_report_path = temp_path
        self.current_report_data = result
        self.download_btn.setEnabled(True)
        self.download_full_btn.setEnabled(True)
        self.generate_btn.setEnabled(True)
        QMessageBox.information(self, "报表生成完毕",
            f"报表已生成，点击「下载报表」或「下载完整报表」保存文件\n\n{self._validation_msg}")

    def on_generate_error(self, error_msg):
        self.progress.close()
        self.generate_btn.setEnabled(True)
        QMessageBox.critical(self, "生成报表失败", error_msg)

    def save_report(self):
        try:
            if not hasattr(self, 'current_report_data') or not self.current_report_data:
                QMessageBox.warning(self, "无报表", "请先生成报表")
                return
            desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
            default_path = os.path.join(desktop, "报表结果.xlsx")
            save_path, _ = QFileDialog.getSaveFileName(
                self, "保存报表", default_path, "Excel文件 (*.xlsx)",
                options=QFileDialog.DontUseNativeDialog
            )
            if not save_path:
                return
            calc_sheets = ['行业计算表', '办事处计算表', '毛利_行业计算表', '毛利_办事处计算表']
            result = {k: v for k, v in self.current_report_data.items() if k not in calc_sheets}
            fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
            os.close(fd)
            with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
                for sheet_name, df_sheet in result.items():
                    if df_sheet is not None and not df_sheet.empty:
                        df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                apply_excel_style(writer.book)
            shutil.copy2(temp_path, save_path)
            QMessageBox.information(self, "保存成功", f"报表已保存至：{save_path}")
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "保存失败", f"保存报表时发生错误：\n{str(e)}\n\n{traceback.format_exc()}")

    def save_report_full(self):
        try:
            if not hasattr(self, 'current_report_data') or not self.current_report_data:
                QMessageBox.warning(self, "无报表", "请先生成报表")
                return
            desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
            default_path = os.path.join(desktop, "报表完整版.xlsx")
            save_path, _ = QFileDialog.getSaveFileName(
                self, "保存完整报表", default_path, "Excel文件 (*.xlsx)",
                options=QFileDialog.DontUseNativeDialog
            )
            if not save_path:
                return
            result = self.current_report_data
            fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
            os.close(fd)
            with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
                for sheet_name, df_sheet in result.items():
                    if df_sheet is not None and not df_sheet.empty:
                        df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                apply_excel_style(writer.book)
            shutil.copy2(temp_path, save_path)
            QMessageBox.information(self, "保存成功", f"完整报表已保存至：{save_path}")
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "保存失败", f"保存报表时发生错误：\n{str(e)}\n\n{traceback.format_exc()}")
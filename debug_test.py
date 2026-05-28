import pandas as pd

# 模拟pdt_rows的结构
pdt_rows = [
    {'二级分类': '综合服务', '三级分类': '安全驻场服务', '合计': 100},
    {'二级分类': '综合服务', '三级分类': '云端运营服务', '合计': 200},
    {'二级分类': '综合服务', '三级分类': '安全攻防服务', '合计': 150},
    {'二级分类': '综合服务', '三级分类': '安全咨询服务', '合计': 80},
    {'二级分类': '综合服务', '三级分类': '安全服务工具', '合计': 50},
    {'二级分类': '综合服务', '三级分类': '海外安全服务', '合计': 30},
    {'二级分类': '综合服务', '三级分类': '综合服务小计', '合计': 610},
    {'二级分类': '数据安全', '三级分类': '数据安全管理中心', '合计': 120},
    {'二级分类': '数据安全', '三级分类': '数据安全OEM IN', '合计': 80},
    {'二级分类': '数据安全', '三级分类': '数据安全产品转售', '合计': 40},
    {'二级分类': '数据安全', '三级分类': '数据安全服务', '合计': 60},
    {'二级分类': '数据安全', '三级分类': '数据安全小计', '合计': 300},
    {'二级分类': '产教融合', '三级分类': '产教融合产品转售', '合计': 50},
    {'二级分类': '产教融合', '三级分类': '产教融合服务', '合计': 70},
    {'二级分类': '产教融合', '三级分类': '产教融合小计', '合计': 120},
    {'二级分类': '安全集成服务', '三级分类': '安全集成服务', '合计': 200},
    {'二级分类': '服务产品经理业绩总计', '三级分类': '', '合计': 1230},
]

year25_data = (500, 200, 100)  # 综合、数据、产教

# 查找行
zonghe_row = None
shuju_row = None
chanjiao_row = None
total_row = None

print(f'year25_data: 综合={year25_data[0]}, 数据={year25_data[1]}, 产教={year25_data[2]}')
print(f'pdt_rows数量: {len(pdt_rows)}')
for idx, row in enumerate(pdt_rows):
    print(f'行{idx+2}: 二级={row.get("二级分类")}, 三级={row.get("三级分类")}')

for row in pdt_rows:
    if (row.get('三级分类') or '').endswith('小计'):
        if row['二级分类'] == '综合服务':
            zonghe_row = row
        elif row['二级分类'] == '数据安全':
            shuju_row = row
        elif row['二级分类'] == '产教融合':
            chanjiao_row = row
    elif row['二级分类'] == '服务产品经理业绩总计':
        total_row = row

print(f'找到的行: 综合={zonghe_row is not None}, 数据={shuju_row is not None}, 产教={chanjiao_row is not None}, 总计={total_row is not None}')

if zonghe_row and shuju_row and chanjiao_row and total_row:
    zonghe_25, shuju_25, chanjiao_25 = year25_data
    zonghe_row['25年合计'] = zonghe_25
    shuju_row['25年合计'] = shuju_25
    chanjiao_row['25年合计'] = chanjiao_25
    total_row['25年合计'] = zonghe_25 + shuju_25 + chanjiao_25
    
    def calc_yoy(s_val, w_val):
        if w_val and w_val != 0:
            return f'{round((s_val - w_val) / w_val * 100, 2)}%'
        return ''
    
    zonghe_row['25年同比'] = calc_yoy(zonghe_row['合计'], zonghe_25)
    shuju_row['25年同比'] = calc_yoy(shuju_row['合计'], shuju_25)
    chanjiao_row['25年同比'] = calc_yoy(chanjiao_row['合计'], chanjiao_25)
    total_row['25年同比'] = calc_yoy(total_row['合计'], total_row['25年合计'])
    
    print('数据设置完成!')
    print(f'综合: 合计={zonghe_row["合计"]}, 25年={zonghe_row["25年合计"]}, 同比={zonghe_row["25年同比"]}')
    print(f'数据: 合计={shuju_row["合计"]}, 25年={shuju_row["25年合计"]}, 同比={shuju_row["25年同比"]}')
    print(f'产教: 合计={chanjiao_row["合计"]}, 25年={chanjiao_row["25年合计"]}, 同比={chanjiao_row["25年同比"]}')
    print(f'总计: 合计={total_row["合计"]}, 25年={total_row["25年合计"]}, 同比={total_row["25年同比"]}')
else:
    print('未找到所有需要的行!')

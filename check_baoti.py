import openpyxl
from collections import Counter

wb = openpyxl.load_workbook(r'C:\Users\Faye\Desktop\集运业务流量包数据统计_20260511.xlsx', read_only=True)
ws = wb['Sheet1']
rows = list(ws.iter_rows(values_only=True))
headers = rows[0]
print('所有字段：')
for i, h in enumerate(headers):
    print(f'  [{i}] {h}')

# Find 包体 column
baoti_col = None
type_col = None
success_col = None
date_col = None
for i, h in enumerate(headers):
    if h == '包体':
        baoti_col = i
    if h == '类型':
        type_col = i
    if h == '成功订单':
        success_col = i
    if h == '日期':
        date_col = i

print(f'\n包体列索引: {baoti_col}')
print(f'类型列索引: {type_col}')
print(f'成功订单列索引: {success_col}')

# Get unique values in 包体 field (for non-流量包 rows)
baoti_vals = Counter()
for row in rows[1:]:
    row_type = row[type_col]
    if row_type and str(row_type).strip() == '流量包':
        continue
    baoti = row[baoti_col]
    if baoti:
        baoti_vals[baoti] += 1

print(f'\n不含流量包的“包体”字段值分布（共{len(baoti_vals)}种）：')
for k, v in baoti_vals.most_common(20):
    print(f'  {k}: {v}行')

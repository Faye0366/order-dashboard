import openpyxl
wb = openpyxl.load_workbook(r'C:\Users\Faye\Desktop\集运业务流量包数据统计_20260511.xlsx', read_only=True)
ws = wb['Sheet1']
rows = list(ws.iter_rows(values_only=True))
headers = rows[0]
print('字段列表（共' + str(len(headers)) + '个）:')
for i, h in enumerate(headers):
    print('  [' + str(i) + '] ' + str(h))

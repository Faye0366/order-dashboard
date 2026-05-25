import openpyxl
from datetime import datetime, timedelta

wb = openpyxl.load_workbook(r"C:\Users\Faye\Desktop\集运业务流量包数据统计_20260511.xlsx", read_only=True)
ws = wb["Sheet1"]
rows = list(ws.iter_rows(values_only=True))
headers = rows[0]

col_map = {h: i for i, h in enumerate(headers)}
date_col = col_map['日期']
success_col = col_map['成功订单']
type_col = col_map['类型']

def to_date(v):
    if isinstance(v, datetime): return v
    if isinstance(v, str):
        try: return datetime.strptime(v[:10], "%Y-%m-%d")
        except: return None
    return None

may_sum = 0
apr_sum = 0

for row in rows[1:]:
    d = to_date(row[date_col])
    if d is None: continue
    if str(row[type_col]).strip() != '流量包': continue
    s = row[success_col]
    try: s = int(s) if s else 0
    except: s = 0
    if d.year == 2026 and d.month == 5 and d.day <= 10:
        may_sum += s
    elif d.year == 2026 and d.month == 4 and d.day <= 10:
        apr_sum += s

print(f"May 1-10: {may_sum}")
print(f"Apr 1-10: {apr_sum}")
mom = ((may_sum - apr_sum) / apr_sum * 100) if apr_sum else 0
print(f"MoM: {mom:.1f}%")

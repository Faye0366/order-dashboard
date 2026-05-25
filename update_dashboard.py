#!/usr/bin/env python3
"""
一键更新看板数据
使用方式：双击 update_dashboard.bat，或直接运行本脚本
流程：桌面 Excel → 重新计算 JSON → 自动嵌入 HTML
"""

import subprocess
import json
import re
import sys
import os
from datetime import datetime

WORKSPACE   = r"C:\Users\Faye\WorkBuddy\2026-05-11-task-1"
HTML_FILE   = os.path.join(WORKSPACE, "order_dashboard.html")
JSON_FILE   = os.path.join(WORKSPACE, "dashboard_data.json")
PYTHON_EXE  = r"C:\Users\Faye\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
PROCESS_PY  = os.path.join(WORKSPACE, "process_data.py")

# ── 1. 运行 process_data.py 生成最新 JSON ────────────────────────────────────
print("=" * 50)
print("  集运业务订单看板 · 一键更新")
print("=" * 50)
print("\n[1/3] 正在从 Excel 读取最新数据...")

env = os.environ.copy()
env["PYTHONUTF8"] = "1"
result = subprocess.run(
    [PYTHON_EXE, PROCESS_PY],
    capture_output=True, text=True, encoding="utf-8", cwd=WORKSPACE, env=env
)
if result.returncode != 0:
    print("❌ 数据处理失败：")
    print(result.stderr)
    sys.exit(1)
print(result.stdout.strip())

# ── 2. 加载 JSON，计算元数据 ──────────────────────────────────────────────────
print("\n[2/3] 正在计算看板元数据...")

with open(JSON_FILE, "r", encoding="utf-8") as f:
    data = json.load(f)

monthly_keys = sorted(data["monthly"].keys())     # ["2026-01", ..., "2026-05"]
month_labels = [str(int(m[5:])) + "月" for m in monthly_keys]  # ["1月", ..., "5月"]

# 月均用前 N-1 个完整月份
n_full_months = len(monthly_keys) - 1
last_full_month  = monthly_keys[-2] if len(monthly_keys) >= 2 else monthly_keys[0]
prev_full_month  = monthly_keys[-3] if len(monthly_keys) >= 3 else monthly_keys[0]
last_full_cn     = str(int(last_full_month[5:])) + "月"
prev_full_cn     = str(int(prev_full_month[5:])) + "月"

# 数据截止日
all_dates = [
    d
    for month_data in data["daily_by_month"].values()
    for d in month_data.keys()
]
max_date_str = max(all_dates) if all_dates else datetime.now().strftime("%Y-%m-%d")
max_dt        = datetime.strptime(max_date_str, "%Y-%m-%d")
max_month_cn  = str(max_dt.month) + "月"
max_day_int   = max_dt.day
today_str     = datetime.now().strftime("%Y-%m-%d")

meta        = data.get("meta", {})
range_start = meta.get("liuliang_range", "2026-01-01 to ?").split(" to ")[0]

print(f"   数据截止：{max_date_str}  |  覆盖月份：{', '.join(monthly_keys)}")

# ── 3. 读取 HTML，逐项替换 ──────────────────────────────────────────────────────
print("\n[3/3] 正在更新 order_dashboard.html...")

with open(HTML_FILE, "r", encoding="utf-8") as f:
    html = f.read()

# ─ 3-A：替换 DATA 对象 ─────────────────────────────────────────────────────────
data_to_embed = {k: v for k, v in data.items() if k != "meta"}
data_json_str = json.dumps(data_to_embed, ensure_ascii=False, indent=2)

start_marker = "const DATA = {"
end_marker   = "\nconst MONTHS = ["
start_idx    = html.index(start_marker)
end_idx      = html.index(end_marker, start_idx)
html = html[:start_idx] + "const DATA = " + data_json_str + ";" + html[end_idx:]

# ─ 3-B：替换 MONTHS / MONTH_LABELS ───────────────────────────────────────────
months_js      = json.dumps(monthly_keys, ensure_ascii=False)
month_labels_js = json.dumps(month_labels, ensure_ascii=False)
html = re.sub(r"const MONTHS = \[.*?\];",      f"const MONTHS = {months_js};",       html)
html = re.sub(r"const MONTH_LABELS = \[.*?\];", f"const MONTH_LABELS = {month_labels_js};", html)

# ─ 3-C：更新 header 更新日期 ─────────────────────────────────────────────────
html = re.sub(r"更新：\d{4}-\d{2}-\d{2}", f"更新：{today_str}", html)

# ─ 3-D：更新 JS 字符串里的数据范围 ──────────────────────────────────────────
html = re.sub(
    r"(订单|数据范围)：(\d{4}-\d{2}-\d{2}) ~ \d{4}-\d{2}-\d{2}",
    lambda m: m.group(1) + "：" + m.group(2) + " ~ " + max_date_str,
    html
)

# ─ 3-E：更新 footer 文件名 ──────────────────────────────────────────────────
html = re.sub(
    r"数据来源：集运业务流量包数据统计.*?\.xlsx",
    "数据来源：集运业务流量包数据统计.xlsx",
    html
)

# ─ 3-F：更新 HTML 卡片里的静态日期文字 ─────────────────────────────────────
html = re.sub(r"截至\d+月\d+日",     f"截至{max_month_cn}{max_day_int}日",            html)
html = re.sub(r"\d+月1-\d+日累计",   f"{max_month_cn}1-{max_day_int}日累计",          html)
html = re.sub(r"不含短时包·\d+月1-\d+日", f"不含短时包·{max_month_cn}1-{max_day_int}日", html)

# ─ 3-G：更新月均计算用的月份数量（slice 取前 N 个完整月） ───────────────────
html = re.sub(r"(MONTHS\.slice\(0, )\d+(\)\.map)", f"\\g<1>{n_full_months}\\g<2>", html)

# ─ 3-H：更新月均 badge 的对比月份（如"4月较3月"） ───────────────────────────
html = re.sub(
    r'(const lastFull = DATA\.monthly\[)"[^"]*"(\])',
    f'\\g<1>"{last_full_month}"\\g<2>', html
)
html = re.sub(
    r'(const prevFull = DATA\.monthly\[)"[^"]*"(\])',
    f'\\g<1>"{prev_full_month}"\\g<2>', html
)
html = re.sub(
    r'(const lastFullNods = DATA\.monthly_nods\[)"[^"]*"(\])',
    f'\\g<1>"{last_full_month}"\\g<2>', html
)
html = re.sub(
    r'(const prevFullNods = DATA\.monthly_nods\[)"[^"]*"(\])',
    f'\\g<1>"{prev_full_month}"\\g<2>', html
)
html = re.sub(
    r"(avgBadge\.textContent = ')[^']*月较[^']*月 '",
    f"\\g<1>{last_full_cn}较{prev_full_cn} '", html
)
html = re.sub(
    r"(avgNodsBadge\.textContent = ')[^']*月较[^']*月 '",
    f"\\g<1>{last_full_cn}较{prev_full_cn} '", html
)

# ─ 3-I：写入 HTML ────────────────────────────────────────────────────────────
with open(HTML_FILE, "w", encoding="utf-8") as f:
    f.write(html)

print(f"""
{'=' * 50}
✅ 更新完成！
   数据截止：{max_date_str}
   覆盖月份：{', '.join(monthly_keys)}
   输出文件：order_dashboard.html
{'=' * 50}
""")

import json
with open(r'C:\Users\Faye\WorkBuddy\2026-05-11-task-1\dashboard_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)
print('Top-level keys:', list(data.keys()))
if 'annual_kuandai' in data:
    print('annual_kuandai:', data['annual_kuandai'])
if 'monthly_kuandai' in data:
    print('monthly_kuandai:', data['monthly_kuandai'])
else:
    print('monthly_kuandai: NOT FOUND')

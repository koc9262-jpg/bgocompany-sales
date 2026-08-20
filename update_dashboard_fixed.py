#!/usr/bin/env python3
import openpyxl, os, re, json, sys
from datetime import date

script_dir = os.path.dirname(os.path.abspath(__file__))
data_dir = os.path.join(script_dir, 'data')

if not os.path.exists(data_dir):
    print("ERROR: data/ not found", file=sys.stderr); sys.exit(1)

all_xlsx = [f for f in os.listdir(data_dir)
            if re.match(r'\d{4}_b_', f) and f.endswith('.xlsx')]
if not all_xlsx:
    print("INFO: no xlsx in data/ - skip", file=sys.stderr); sys.exit(0)

prefixes = sorted(set(re.match(r'(\d{4})_', f).group(1)
                  for f in all_xlsx if re.match(r'(\d{4})_', f)))
prefix = prefixes[-1]
print(f"Parsing prefix: {prefix}", file=sys.stderr)

files = sorted([f for f in all_xlsx if f.startswith(prefix)])
branch_files = {}
for f in files:
    # 파일명에서 (1), (2) 같은 중복 번호 제거하고 key 생성
    m = re.match(rf'({prefix}_b_\d+_.+?)(?:\s*\(\d+\))?\.xlsx$', f)
    key = m.group(1) if m else re.sub(r'\.xlsx$', '', f)
    branch_files[key] = f

def safe(row, i):
    try: return row[i]
    except (IndexError, TypeError): return None

def pct(v):
    if v is None: return 0.0
    try:
        v = float(v)
        return round(v * 100, 1) if abs(v) < 10 else round(v, 1)
    except: return 0.0

def to_int(v):
    if v is None: return 0
    try: return int(float(v))
    except: return 0

results = []
for key in sorted(branch_files.keys()):
    fname = branch_files[key]
    fpath = os.path.join(data_dir, fname)
    print(f"Reading {fname}", file=sys.stderr)
    try:
        wb = openpyxl.load_workbook(fpath, data_only=True)
    except Exception as e:
        print(f"  ERROR loading {fname}: {e}", file=sys.stderr)
        continue

    # 매출 시트 찾기
    ws = wb.active
    for sname in wb.sheetnames:
        if '매출' in sname or sname.lower() == 'sheet1':
            ws = wb[sname]
            break

    rows = [[cell.value for cell in row] for row in ws.iter_rows()]

    # 지점명: row 0, col 20
    branch_name = safe(rows[0], 20) if rows else None
    if not branch_name:
        # 파일명에서 지점명 추출
        m2 = re.match(rf'{prefix}_b_\d+_(.+?)_총매출', fname)
        branch_name = m2.group(1) if m2 else key

    r2 = rows[2] if len(rows) > 2 else []
    r3 = rows[3] if len(rows) > 3 else []
    r4 = rows[4] if len(rows) > 4 else []
    r7 = rows[7] if len(rows) > 7 else []
    r10 = rows[10] if len(rows) > 10 else []

    # PT 순위 시트
    ranks = []
    for sname in wb.sheetnames:
        if '순위' in sname:
            ws2 = wb[sname]
            for row in list(ws2.iter_rows())[3:]:
                rv = [c.value for c in row]
                nm = safe(rv, 3)
                amt = safe(rv, 4)
                if isinstance(nm, str) and nm.strip() and isinstance(amt, (int, float)) and amt > 0:
                    ranks.append({'name': nm.strip(), 'amount': int(amt)})

    results.append({
        'branch': str(branch_name),
        'target_fc':  to_int(safe(r2, 2)),
        'target_pt':  to_int(safe(r2, 4)),
        'target_tot': to_int(safe(r2, 6)),
        'actual_fc':  to_int(safe(r3, 2)),
        'actual_pt':  to_int(safe(r3, 4)),
        'actual_tot': to_int(safe(r3, 6)),
        'rate_fc':    pct(safe(r4, 2)),
        'rate_pt':    pct(safe(r4, 4)),
        'rate_tot':   pct(safe(r4, 6)),
        'fc_new_cnt': to_int(safe(r7, 10)),
        'fc_new_amt': to_int(safe(r7, 11)),
        'fc_re_cnt':  to_int(safe(r7, 12)),
        'fc_re_amt':  to_int(safe(r7, 13)),
        'pt_new_cnt': to_int(safe(r7, 17)),
        'pt_new_amt': to_int(safe(r7, 18)),
        'pt_re_cnt':  to_int(safe(r7, 19)),
        'pt_re_amt':  to_int(safe(r7, 20)),
        'total_card': to_int(safe(r10, 12)),
        'cash_out':   to_int(safe(r10, 14)),
        'account':    to_int(safe(r10, 16)),
        'deduction':  to_int(safe(r10, 18)),
        'real_cash_out': to_int(safe(r10, 20)),
        'ranks': ranks
    })
    print(f"  OK: {branch_name}", file=sys.stderr)

if not results:
    print("ERROR: no results parsed", file=sys.stderr); sys.exit(1)

today = date.today()
year, month, day = today.year, today.month, today.day
data_json = json.dumps(results, ensure_ascii=False)

html_path = os.path.join(script_dir, 'index.html')
with open(html_path, encoding='utf-8') as f:
    html = f.read()

html = re.sub(r'const D_RAW = \[.*?\];', f'const D_RAW = {data_json};', html, flags=re.DOTALL)
html = re.sub(r'20\d\d년 \d+월 \d+일', f'{year}년 {month:02d}월 {day:02d}일', html)

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'Done: updated {len(results)} branches', file=sys.stderr)

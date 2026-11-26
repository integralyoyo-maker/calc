#!/usr/bin/env python3
"""違規入案代碼表 xlsx → data/violation_codes.json（現行）與 data/violation_codes_old.json（已不適用）
用法：python3 tools/build_codes.py <代碼表.xlsx> [民國日期YYYMMDD，預設取檔名中的西元日期]
需要 openpyxl。工作表：條款代碼表、不得郵繳條款、慢車不得郵繳條款。"""
import sys, re, json, os, datetime
import openpyxl

src = sys.argv[1]
wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data')
os.makedirs(out_dir, exist_ok=True)

TODAY = datetime.date.today()

# ---------- 檔案日期 ----------
m = re.search(r'(20\d{2})(\d{2})(\d{2})', os.path.basename(src))
file_roc = None
if m:
    y, mo, d = map(int, m.groups()); file_roc = f'{y-1911}{mo:02d}{d:02d}'
if len(sys.argv) > 2: file_roc = sys.argv[2]

rows = list(wb['條款代碼表'].iter_rows(values_only=True))
body = rows[1:]

# ---------- 圖例（取自工作表右側，若 Excel 改版則隨之更新）----------
vehicle = {}; handling = {}
for r in rows[:60]:
    if len(r) > 18 and r[17] is not None and r[18] is not None: vehicle[str(r[17])] = str(r[18])
    if len(r) > 22 and r[21] is not None and r[22] is not None and r[20] in (None, '特殊處罰'):
        handling[str(r[21])] = str(r[22])
# 特殊處罰圖例只取到 'h' 為止（其後為吊扣、歸責、禁考圖例）
hk = {}
seen_end = False
for r in rows[:60]:
    if len(r) > 22 and r[21] is not None and r[22] is not None:
        if r[20] in ('吊扣/吊銷', '歸責對象', '禁考註記'): seen_end = True
        if not seen_end: hk[str(r[21])] = str(r[22])
handling = hk
months = {}
for r in wb['吊扣代碼'].iter_rows(min_row=2, values_only=True):
    if r[0] is not None and isinstance(r[2], int): months[str(r[0])] = r[2]

# ---------- 應到案（不得郵繳）規則 ----------
rules = []
def add_rule(disp, note=None):
    s = str(disp).strip()
    if not s or s == 'None': return
    ln = None
    if note and re.match(r'\d碼', str(note)): ln = int(str(note)[0])
    if '到' in s:
        a, b = [x.strip() for x in re.split(r'\s*到\s*', s.replace('\n', ' ').replace(' ', '到', 0)) if x.strip()][:2]
        rules.append(('range', int(a), int(b)))
        return
    s2 = s.replace('\n', '')
    mm = re.match(r"^(\d+)\s*&&\s*rule_no\[(\d)\]=='(\d)'$", s2)
    if mm: rules.append(('prefix_idx', mm.group(1), int(mm.group(2)), mm.group(3))); return
    mm = re.match(r'^(\d+)\(358沒有\)$', s2)
    if mm: rules.append(('prefix_ex', mm.group(1), ['358'], ln)); return
    mm = re.match(r'^(\d+)$', s2)
    if mm: rules.append(('prefix', mm.group(1), ln)); return
    print('!! 無法解析的不得郵繳規則：', repr(s), file=sys.stderr)

for r in wb['不得郵繳條款'].iter_rows(min_row=2, values_only=True):
    if r[2] is None: continue
    s = str(r[2])
    m2 = re.match(r'^(\d+)到(\d+)$', s)
    if m2: rules.append(('range', int(m2.group(1)), int(m2.group(2)))); continue
    add_rule(r[2], r[3])
for r in wb['慢車不得郵繳條款'].iter_rows(min_row=3, values_only=True):
    if r[3] is not None: add_rule(r[3])

def must_appear(code):
    s = str(code)
    for rl in rules:
        k = rl[0]
        if k == 'range':
            if s.isdigit() and rl[1] <= int(s) <= rl[2]: return True
        elif k == 'prefix':
            if s.startswith(rl[1]) and (rl[2] is None or len(s) == rl[2]): return True
        elif k == 'prefix_ex':
            if s.startswith(rl[1]) and not any(s.startswith(e) for e in rl[2]) and (rl[3] is None or len(s) == rl[3]): return True
        elif k == 'prefix_idx':
            if s.startswith(rl[1]) and len(s) > rl[2] and s[rl[2]] == rl[3]: return True
    return False

# ---------- 現行/已不適用判斷 ----------
def roc_to_date(t):
    y, mo, d = map(int, t)
    try: return datetime.date(y + 1911, mo, d)
    except ValueError: return None

def status(note):
    """備註由新到舊記錄；取第一個『新增/不適用/重新適用』決定現況。"""
    if not note: return 'cur', None
    s = str(note)
    for seg in re.split(r'[；;]', s):
        if '重新適用' in seg: return 'cur', seg
        if '不適用' in seg:
            mm = re.search(r'(\d+)/(\d+)/(\d+)', seg)
            if mm:
                d = roc_to_date(mm.groups())
                if d and d > TODAY: return 'cur', s   # 尚未生效的停用
            return 'old', s
        if '新增' in seg:
            mm = re.search(r'(\d+)/(\d+)/(\d+)', seg)
            if mm:
                d = roc_to_date(mm.groups())
                if d and d > TODAY: return 'future', s
            return 'cur', s
    return 'cur', s

def num(v):
    return int(v) if isinstance(v, (int, float)) else 0

cur, old = [], []
max_date = (0, 0, 0)
for r in body:
    if r[1] is None: continue
    code = re.sub(r'^\(刪除\)', '', str(r[1]))
    st, note = status(r[0])
    for mm in re.findall(r'(\d+)/(\d+)/(\d+)', str(r[0] or '')):
        t = tuple(map(int, mm))
        if 100 <= t[0] <= 200 and t > max_date: max_date = t
    if str(r[1]).startswith('(刪除)'): st = 'old'
    h = [x for x in re.split(r'[、,，]', str(r[13])) if x not in ('', '0', 'None')] if r[13] is not None else []
    rec = {
        'c': code, 't': re.sub(r'\s+', ' ', str(r[2] or '')).strip(),
        'f': [num(r[3]), num(r[4]), num(r[5]), num(r[6])],
        'v': str(r[7]) if r[7] is not None else '0',
        'o': str(r[8]) if r[8] is not None else '0',
        'p': num(r[9]),
        's': str(r[10]) if r[10] not in (None, 0) else '',
        'k': str(r[11]) if r[11] not in (None, 0) else '',
        'h': h,
        'a': 1 if must_appear(code) else 0,
    }
    if note: rec['n'] = str(note)
    if st == 'future': rec['n'] = (rec.get('n', '') + '（尚未生效）')
    (old if st == 'old' else cur).append(rec)

note_version = f'{max_date[0]}{max_date[1]:02d}{max_date[2]:02d}' if max_date[0] else ''
version = file_roc or note_version   # 版本＝Excel 檔名上的日期（民國）；無日期時退回備註欄最新日期
meta = {
    'version': version, 'noteVersion': note_version,
    'vehicle': vehicle, 'handling': handling, 'months': months,
    'counts': {'cur': len(cur), 'old': len(old)},
}
def dump(path, obj):
    with open(path, 'w', encoding='utf-8') as f: json.dump(obj, f, ensure_ascii=False, separators=(',', ':'))
    print(path, os.path.getsize(path) // 1024, 'KB')
dump(os.path.join(out_dir, 'violation_codes.json'), {'meta': meta, 'rows': cur})
dump(os.path.join(out_dir, 'violation_codes_old.json'), {'rows': old})
print(meta['version'], meta['noteVersion'], meta['counts'], 'rules:', len(rules), '應到案(現行):', sum(x['a'] for x in cur))

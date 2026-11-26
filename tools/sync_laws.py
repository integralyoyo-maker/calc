#!/usr/bin/env python3
"""從全國法規資料庫開放資料同步五部交通法規到 data/laws/。
來源：https://law.moj.gov.tw/api/Ch/Law/JSON （中央法規 JSON，可能為 zip）。
失敗時不覆蓋既有法規檔，只寫 data/laws/status.json 記錄原因（方便從網站或 repo 查看）。"""
import io, json, os, sys, time, zipfile, urllib.request, datetime

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'laws')
URL = 'https://law.moj.gov.tw/api/Ch/Law/JSON'
URL_ORDER = 'https://law.moj.gov.tw/api/Ch/Order/JSON'
WANT = [  # (短名, 法規全名)
    ('處罰條例', '道路交通管理處罰條例'),
    ('安全規則', '道路交通安全規則'),
    ('裁罰基準細則', '違反道路交通管理事件統一裁罰基準及處理細則'),
    ('設置規則', '道路交通標誌標線號誌設置規則'),
    ('事故處理辦法', '道路交通事故處理辦法'),
]
os.makedirs(OUT, exist_ok=True)
now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime('%Y-%m-%d %H:%M')

def status(ok, msg, extra=None):
    d = {'ok': ok, 'time': now, 'message': msg}
    if extra: d.update(extra)
    with open(os.path.join(OUT, 'status.json'), 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    print(('OK ' if ok else 'FAIL ') + msg)

try:
    def fetch(url):
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 calc-sync'})
        raw = urllib.request.urlopen(req, timeout=180).read()
        if raw[:2] == b'PK':
            z = zipfile.ZipFile(io.BytesIO(raw))
            name = next(n for n in z.namelist() if n.lower().endswith('.json'))
            raw = z.read(name)
        data = json.loads(raw.decode('utf-8-sig'))
        return data['Laws'] if isinstance(data, dict) and 'Laws' in data else (data['Orders'] if isinstance(data, dict) and 'Orders' in data else data)
    laws = fetch(URL) + fetch(URL_ORDER)   # 法律 + 命令（規則、辦法、細則屬於命令）
    by_name = {}
    for l in laws:
        by_name.setdefault(l.get('LawName'), []).append(l)
    index = []
    files = {}
    att_info = {}
    for short, full in WANT:
        # 名稱可能帶「（92.09.24 訂定）」之類後綴；只取未廢止者
        c = [l for n, ls in by_name.items() if n and (n == full or n.startswith(full + '（')) for l in ls if not l.get('LawAbandonNote')]
        if len(c) != 1:
            raise RuntimeError(f'找不到唯一的「{full}」（找到 {len(c)} 筆）：' + '、'.join(f"{l.get('LawName')}[廢止註記:{(l.get('LawAbandonNote') or '無')[:20]}]" for ns in by_name.values() for l in ns if l.get('LawName', '').startswith(full))[:600])
        l = c[0]
        pcode = l['LawURL'].split('pcode=')[-1]
        arts = [{'t': a.get('ArticleType', 'A'), 'no': (a.get('ArticleNo') or '').strip(),
                 'c': (a.get('ArticleContent') or '').replace('\r\n', '\n').replace('\r', '\n').rstrip()}
                for a in l['LawArticles']]
        if not any(a['t'] == 'A' for a in arts):
            raise RuntimeError(f'「{full}」沒有條文')
        att_info[pcode] = {'name': full, 'attachments': l.get('LawAttachements') or []}
        files[pcode] = {'name': full, 'short': short, 'pcode': pcode, 'url': l['LawURL'],
                        'modified': l.get('LawModifiedDate', ''), 'effective': l.get('LawEffectiveDate', ''),
                        'effectiveNote': (l.get('LawEffectiveNote') or '').replace('\r\n', '\n'),
                        'attachments': len(l.get('LawAttachements') or []), 'articles': arts}
        index.append({'short': short, 'name': full, 'pcode': pcode, 'modified': l.get('LawModifiedDate', ''),
                      'articles': sum(1 for a in arts if a['t'] == 'A')})
    with open(os.path.join(OUT, 'attachments_info.json'), 'w', encoding='utf-8') as f:
        json.dump(att_info, f, ensure_ascii=False, indent=1)
    # 與現有檔案比對：內容完全相同就不動（避免每天都產生提交）
    def dumps(d): return json.dumps(d, ensure_ascii=False, separators=(',', ':'))
    changed = False
    for pcode, d in files.items():
        path = os.path.join(OUT, pcode + '.json')
        if not os.path.exists(path) or open(path, encoding='utf-8').read() != dumps(d):
            changed = True
    ipath = os.path.join(OUT, 'index.json')
    prev_ok = True
    spath = os.path.join(OUT, 'status.json')
    if os.path.exists(spath):
        prev_ok = json.load(open(spath, encoding='utf-8')).get('ok', True)
    if changed or not os.path.exists(ipath):
        for pcode, d in files.items():
            with open(os.path.join(OUT, pcode + '.json'), 'w', encoding='utf-8') as f:
                f.write(dumps(d))
        with open(ipath, 'w', encoding='utf-8') as f:
            json.dump({'syncedAt': now, 'source': URL, 'laws': index}, f, ensure_ascii=False, indent=1)
        status(True, '同步成功（有更新）', {'laws': index})
    elif not prev_ok:
        status(True, '同步成功（無變動）', {'laws': index})
    else:
        print('OK 無變動')
except Exception as e:
    status(False, f'{type(e).__name__}: {e}')
    sys.exit(0)   # 不讓排程標紅；失敗原因記錄在 status.json

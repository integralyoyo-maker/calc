#!/usr/bin/env python3
"""從全國法規資料庫開放資料同步五部交通法規到 data/laws/。
來源：https://law.moj.gov.tw/api/Ch/Law/JSON （中央法規 JSON，可能為 zip）。
失敗時不覆蓋既有法規檔，只寫 data/laws/status.json 記錄原因（方便從網站或 repo 查看）。"""
import io, json, os, re, sys, time, zipfile, shutil, subprocess, tempfile, urllib.request, datetime

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

FIG_DIR = os.path.join(OUT, 'figs')
MAX_PAGES = 8
def http_get(url, tries=3):
    last = None
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 calc-sync'})
            return urllib.request.urlopen(req, timeout=120).read()
        except Exception as e:
            last = e; time.sleep(2 * (i + 1))
    raise last

def trim_save(img, path, maxw=900):
    from PIL import Image, ImageChops
    img = img.convert('RGB')
    bg = Image.new('RGB', img.size, (255, 255, 255))
    bbox = ImageChops.difference(img, bg).getbbox()
    if bbox:
        pad = 12
        bbox = (max(bbox[0] - pad, 0), max(bbox[1] - pad, 0), min(bbox[2] + pad, img.width), min(bbox[3] + pad, img.height))
        img = img.crop(bbox)
    if img.width > maxw:
        img = img.resize((maxw, int(img.height * maxw / img.width)))
    img.save(path, 'JPEG', quality=82, optimize=True)

def build_figs(att):
    """設置規則等「第N條完整條文」附件 → 轉成圖片（PDF 逐頁轉檔；僅有 ODT 時取其內嵌圖片）。
    以 FileId 做增量更新：FileId 沒變就不重做。回傳 (figs 對照表, 失敗清單)。"""
    from PIL import Image
    os.makedirs(FIG_DIR, exist_ok=True)
    mpath = os.path.join(FIG_DIR, 'manifest.json')
    old = json.load(open(mpath, encoding='utf-8')) if os.path.exists(mpath) else {}
    cand = {}
    for a in att:
        m = re.match(r'第\s*([\d-]+)\s*條完整條文\.(\w+)$', a.get('FileName', ''), re.I)
        if not m: continue
        key, ext = m.group(1), m.group(2).upper()
        fid = a['FileURL'].split('FileId=')[-1]
        if key not in cand or (ext == 'PDF' and cand[key]['ext'] != 'PDF'):
            cand[key] = {'ext': ext, 'fid': fid, 'url': a['FileURL']}
    manifest, errors = {}, []
    for key, c in sorted(cand.items()):
        prev = old.get(key)
        if prev and prev.get('fid') == c['fid'] and prev.get('ext') == c['ext'] and all(os.path.exists(os.path.join(FIG_DIR, f)) for f in prev.get('files', [])):
            manifest[key] = prev; continue
        try:
            data = http_get(c['url'])
            for f in (old.get(key) or {}).get('files', []):
                try: os.remove(os.path.join(FIG_DIR, f))
                except OSError: pass
            files = []
            with tempfile.TemporaryDirectory() as td:
                if c['ext'] == 'PDF':
                    src = os.path.join(td, 'a.pdf'); open(src, 'wb').write(data)
                    subprocess.run(['pdftoppm', '-r', '110', '-png', '-l', str(MAX_PAGES), src, os.path.join(td, 'p')], check=True, timeout=180)
                    pages = sorted(f for f in os.listdir(td) if f.startswith('p') and f.endswith('.png'))
                    for i, f in enumerate(pages, 1):
                        name = f'{key}-{i}.jpg'; trim_save(Image.open(os.path.join(td, f)), os.path.join(FIG_DIR, name)); files.append(name)
                else:   # ODT：取內嵌圖片
                    z = zipfile.ZipFile(io.BytesIO(data))
                    pics = [n for n in z.namelist() if n.startswith('Pictures/') and n.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))][:MAX_PAGES]
                    for i, n in enumerate(pics, 1):
                        name = f'{key}-{i}.jpg'; trim_save(Image.open(io.BytesIO(z.read(n))), os.path.join(FIG_DIR, name)); files.append(name)
            if not files: raise RuntimeError('沒有產生圖片')
            manifest[key] = {'fid': c['fid'], 'ext': c['ext'], 'files': files, 'url': c['url']}
        except Exception as e:
            errors.append(f'{key}: {type(e).__name__}: {str(e)[:80]}')
            if prev: manifest[key] = prev
        time.sleep(0.3)
    for key in set(old) - set(manifest):   # 已不存在的附件：清掉舊圖
        for f in old[key].get('files', []):
            try: os.remove(os.path.join(FIG_DIR, f))
            except OSError: pass
    for k, v in manifest.items(): v.setdefault('url', cand.get(k, {}).get('url', ''))
    json.dump(manifest, open(mpath, 'w', encoding='utf-8'), ensure_ascii=False, sort_keys=True, separators=(',', ':'))
    return manifest, errors


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
    fig_errors = []
    for pcode, info in att_info.items():
        if any(re.match(r'第\s*[\d-]+\s*條完整條文', a.get('FileName', '')) for a in info['attachments']):
            figs, fig_errors = build_figs(info['attachments'])
            files[pcode]['figs'] = {k: {'n': len(v['files']), 'url': v.get('url', '')} for k, v in figs.items()}
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
        status(True, '同步成功（有更新）', {'laws': index, 'figErrors': fig_errors})
    elif not prev_ok:
        status(True, '同步成功（無變動）', {'laws': index, 'figErrors': fig_errors})
    else:
        print('OK 無變動')
except Exception as e:
    status(False, f'{type(e).__name__}: {e}')
    sys.exit(0)   # 不讓排程標紅；失敗原因記錄在 status.json

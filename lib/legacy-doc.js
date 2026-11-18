/* legacy-doc.js
 * 處理「副檔名是 .doc 但實際上不是 Word 二進位檔」的檔案：HTML、MHT(單一網頁)、RTF、純文字。
 * 全部在瀏覽器本機執行：HTML 在「禁止腳本、禁止對外連線」的隔離 iframe 內渲染，檔案內容不會離開電腦。
 * 對外介面：window.LegacyDoc = { sniff, toHtml, renderPages }
 */
(function () {
  'use strict';

  // ------------------------------------------------------------------ 工具
  const latin1 = (u8) => { let s = ''; for (let i = 0; i < u8.length; i += 8192) s += String.fromCharCode.apply(null, u8.subarray(i, i + 8192)); return s; };
  const escHtml = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escAttr = (s) => escHtml(s).replace(/"/g, '&quot;');
  const td = (label, bytes, fatal) => new TextDecoder(label, { fatal: !!fatal }).decode(bytes);

  // ------------------------------------------------------------------ 格式辨識
  // 回傳 'rtf' | 'mht' | 'html' | 'wordxml' | 'text' | null
  function sniff(u8) {
    const head = latin1(u8.subarray(0, 4096));
    const trimmed = head.replace(/^﻿|^\xEF\xBB\xBF|^[\s\0]+/, '');
    if (/^\{\\rtf/i.test(trimmed)) return 'rtf';
    if (/^MIME-Version:/i.test(trimmed) || (/Content-Type:\s*multipart\/related/i.test(head) && /boundary=/i.test(head))) return 'mht';
    if (/^<\?xml/i.test(trimmed) && /(w:wordDocument|schemas-microsoft-com:office:word|<w:document\b)/i.test(head)) return 'wordxml';
    if (/<(!doctype\s+html|html|head|body|table|div|p|span|meta|font|h[1-6]|br)\b/i.test(head)) return 'html';
    // UTF-16 BOM 的 HTML
    if (u8[0] === 0xFF && u8[1] === 0xFE) return 'html';
    // 純文字：前 4KB 沒有 NUL 與過多控制字元
    let bad = 0; for (let i = 0; i < Math.min(u8.length, 4096); i++) { const c = u8[i]; if (c === 0 || (c < 9) || (c > 13 && c < 32)) bad++; }
    if (u8.length && bad === 0) return 'text';
    return null;
  }

  // ------------------------------------------------------------------ 文字編碼偵測（台灣常見 Big5）
  function decodeText(u8) {
    if (u8[0] === 0xEF && u8[1] === 0xBB && u8[2] === 0xBF) return td('utf-8', u8.subarray(3));
    if (u8[0] === 0xFF && u8[1] === 0xFE) return td('utf-16le', u8.subarray(2));
    if (u8[0] === 0xFE && u8[1] === 0xFF) return td('utf-16be', u8.subarray(2));
    const head = latin1(u8.subarray(0, 4096));
    let m = head.match(/<meta[^>]+charset\s*=\s*["']?\s*([\w\-]+)/i) || head.match(/<\?xml[^>]+encoding\s*=\s*["']([\w\-]+)/i);
    if (m) { try { return td(m[1].toLowerCase(), u8); } catch (e) { /* 未知編碼，往下 */ } }
    try { return td('utf-8', u8, true); } catch (e) { /* 不是 UTF-8 */ }
    try { return td('big5', u8); } catch (e) { return td('windows-1252', u8); }
  }

  // ------------------------------------------------------------------ MHT (MIME HTML)
  function decodeQP(str) {
    const s = str.replace(/=\r?\n/g, '');
    const out = []; for (let i = 0; i < s.length; i++) { if (s[i] === '=' && /^[0-9A-Fa-f]{2}$/.test(s.substr(i + 1, 2))) { out.push(parseInt(s.substr(i + 1, 2), 16)); i += 2; } else out.push(s.charCodeAt(i) & 0xFF); }
    return new Uint8Array(out);
  }
  function parseHeaders(block) {
    const h = {}; const lines = block.replace(/\r?\n[ \t]+/g, ' ').split(/\r?\n/);
    for (const l of lines) { const k = l.indexOf(':'); if (k > 0) h[l.slice(0, k).trim().toLowerCase()] = l.slice(k + 1).trim(); }
    return h;
  }
  function partBytes(headers, body) {
    const cte = (headers['content-transfer-encoding'] || '').toLowerCase();
    if (cte === 'base64') { const b = atob(body.replace(/[^A-Za-z0-9+/=]/g, '')); const u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; }
    if (cte === 'quoted-printable') return decodeQP(body);
    const u = new Uint8Array(body.length); for (let i = 0; i < body.length; i++) u[i] = body.charCodeAt(i) & 0xFF; return u;
  }
  function mhtToHtml(u8) {
    const raw = latin1(u8);
    const sep = raw.search(/\r?\n\r?\n/);
    const top = parseHeaders(raw.slice(0, sep));
    const bm = (top['content-type'] || '').match(/boundary="?([^";\s]+)"?/i);
    if (!bm) throw new Error('MHT 結構不完整（找不到分隔線）');
    const parts = raw.slice(sep).split('--' + bm[1]);
    let htmlBytes = null, htmlHeaders = null; const res = {};
    for (const p of parts) {
      const k = p.search(/\r?\n\r?\n/); if (k < 0) continue;
      const hd = parseHeaders(p.slice(0, k)); const body = p.slice(k).replace(/^\r?\n\r?\n/, '').replace(/\r?\n$/, '');
      const ct = (hd['content-type'] || '').toLowerCase();
      if (ct.startsWith('text/html') && !htmlBytes) { htmlBytes = partBytes(hd, body); htmlHeaders = hd; }
      else if (ct.startsWith('image/')) {
        const bytes = partBytes(hd, body); let b64 = ''; for (let i = 0; i < bytes.length; i += 8192) b64 += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
        const uri = 'data:' + ct.split(';')[0] + ';base64,' + btoa(b64);
        const loc = hd['content-location']; const cid = (hd['content-id'] || '').replace(/[<>]/g, '');
        if (loc) res[loc] = uri; if (cid) res['cid:' + cid] = uri;
      }
    }
    if (!htmlBytes) throw new Error('MHT 內找不到文件本文');
    const cs = ((htmlHeaders['content-type'] || '').match(/charset="?([\w\-]+)/i) || [])[1];
    let html; try { html = cs ? td(cs.toLowerCase(), htmlBytes) : decodeText(htmlBytes); } catch (e) { html = decodeText(htmlBytes); }
    for (const key of Object.keys(res)) html = html.split(key).join(res[key]);
    return html;
  }

  // ------------------------------------------------------------------ RTF → HTML
  const CS_LABEL = { 136: 'big5', 134: 'gbk', 128: 'shift_jis', 129: 'euc-kr', 161: 'windows-1253', 162: 'windows-1254', 177: 'windows-1255', 178: 'windows-1256', 186: 'windows-1257', 204: 'windows-1251', 238: 'windows-1250', 222: 'windows-874' };
  const CP_LABEL = { 950: 'big5', 936: 'gbk', 932: 'shift_jis', 949: 'euc-kr', 874: 'windows-874', 1250: 'windows-1250', 1251: 'windows-1251', 1252: 'windows-1252', 1253: 'windows-1253', 1254: 'windows-1254', 1255: 'windows-1255', 1256: 'windows-1256', 1257: 'windows-1257' };
  const SKIP_DEST = new Set(['stylesheet', 'info', 'header', 'headerl', 'headerr', 'headerf', 'footer', 'footerl', 'footerr', 'footerf', 'footnote', 'listtable', 'listoverridetable', 'rsidtbl', 'generator', 'latentstyles', 'themedata', 'colorschememapping', 'datastore', 'fldinst', 'nonshppict', 'bkmkstart', 'bkmkend', 'xe', 'tc', 'private', 'upr', 'shp', 'shpinst', 'pgptbl', 'revtbl', 'userprops', 'mmathPr', 'wgrffmtfilter', 'xmlnstbl', 'fchars', 'lchars', 'defchp', 'defpap', 'pnseclvl', 'template', 'ud', 'passwordhash', 'protusertbl', 'background', 'blipuid', 'object', 'objdata', 'result']);
  const BRD_STYLE = { brdrs: 'solid', brdrth: 'solid', brdrdb: 'double', brdrdot: 'dotted', brdrdash: 'dashed', brdrdashsm: 'dashed', brdrhair: 'solid', brdrsh: 'solid' };

  function rtfToHtml(u8) {
    const s = latin1(u8), n = s.length;
    const fonts = {}, colors = [];
    let ansicpg = 1252, deff = 0, i = 0;
    const page = { w: 11906, h: 16838, ml: 1800, mr: 1800, mt: 1440, mb: 1440 };
    let fontCur = null, rgb = { r: null, g: null, b: null };

    const stack = [];
    const newState = () => ({ b: 0, i: 0, ul: 0, strike: 0, fs: 24, cf: 0, f: deff, afs: null, af: null, ab: null, ai: null, dest: null, uc: 1, skip: 0, first: false });
    let st = newState();

    // 輸出狀態
    const out = [];
    let spans = [], lastStyle = null, bytes = [];
    let para = { jc: '', li: 0, fi: 0, ri: 0, sb: 0, sa: 0, sl: 0, slmult: 1, intbl: false, pagebb: false };
    let row = null, rowCells = [], cellParas = [], tableRows = [];
    let pict = null;
    const resetPara = () => { para = { jc: '', li: 0, fi: 0, ri: 0, sb: 0, sa: 0, sl: 0, slmult: 1, intbl: false, pagebb: false }; };
    resetPara();

    const curFid = () => (st.dest === 'fonttbl' && fontCur !== null) ? fontCur : st.f;
    const fontLabel = () => {
      const f = fonts[curFid()]; const cs = f && f.cs;
      if (cs && CS_LABEL[cs]) return CS_LABEL[cs];
      return CP_LABEL[ansicpg] || 'windows-1252';
    };
    // ea=true：東亞(中日韓)字元。RTF 以 \afs / \af / \ab / \ai（配合 \dbch）另外指定它們的字級、字型、粗斜體
    const cssStyle = (ea) => {
      const fid = (ea && st.af !== null && fonts[st.af]) ? st.af : st.f;
      const f = fonts[fid]; let s = '';
      if (f && f.name) s += `font-family:'${f.name.replace(/'/g, '')}','新細明體','PMingLiU','微軟正黑體',serif;`;
      s += `font-size:${((ea && st.afs) ? st.afs : st.fs) / 2}pt;`;
      if (ea && st.ab !== null ? st.ab : st.b) s += 'font-weight:bold;';
      if (ea && st.ai !== null ? st.ai : st.i) s += 'font-style:italic;';
      const deco = (st.ul ? 'underline ' : '') + (st.strike ? 'line-through' : '');
      if (deco.trim()) s += `text-decoration:${deco.trim()};`;
      if (st.cf && colors[st.cf]) s += `color:${colors[st.cf]};`;
      return s;
    };
    const isEA = (ch) => { const c = ch.codePointAt(0); return (c >= 0x2E80 && c <= 0x9FFF) || (c >= 0xAC00 && c <= 0xD7AF) || (c >= 0xF900 && c <= 0xFAFF) || (c >= 0xFF00 && c <= 0xFFEF) || (c >= 0x3000 && c <= 0x303F) || c >= 0x20000; };
    const emitSeg = (txt, ea) => {
      const sty = cssStyle(ea);
      const safe = escHtml(txt).replace(/\t/g, '&emsp;');
      if (lastStyle === sty && spans.length) spans[spans.length - 1] += safe;
      else { lastStyle = sty; spans.push(`\u0000${sty}\u0001` + safe); }
    };
    // 依「是否東亞字元」把文字切段，各段套用對應的字級/字型/粗斜體
    const emit = (txt) => {
      if (!txt) return;
      let seg = '', segEA = null;
      for (const ch of txt) {
        const ea = isEA(ch);
        if (segEA !== null && ea !== segEA) { emitSeg(seg, segEA); seg = ''; }
        segEA = ea; seg += ch;
      }
      if (seg) emitSeg(seg, segEA);
    };
    // spans 內以 \0style\1text 暫存，最後轉成 <span>
    const spansHtml = () => spans.map(x => { const k = x.indexOf('\u0001'); return `<span style="${x.slice(1, k)}">${x.slice(k + 1)}</span>`; }).join('');
    const flushBytes = () => {
      if (!bytes.length) return;
      const b = new Uint8Array(bytes); bytes = [];
      let label = fontLabel();
      // 字型 charset 是 0/1（西文）但文件內碼頁是 CJK 時，雙位元組仍依文件內碼頁解碼
      const f = fonts[curFid()]; if ((!f || !f.cs || f.cs === 0 || f.cs === 1) && ['big5', 'gbk', 'shift_jis', 'euc-kr'].includes(CP_LABEL[ansicpg])) label = CP_LABEL[ansicpg];
      let t; try { t = td(label, b); } catch (e) { t = td('windows-1252', b); }
      textOut(t);
    };
    const textOut = (t) => {
      if (st.dest === 'fonttbl') { if (fontCur !== null) fonts[fontCur].name += t.replace(/;/g, ''); return; }
      if (st.dest) return;
      if (para.intbl === false && tableRows.length && !row) closeTable();
      emit(t);
    };

    const paraStyle = () => {
      let s = '';
      if (para.jc) s += `text-align:${para.jc};`;
      if (para.li) s += `margin-left:${para.li / 15}px;`;
      if (para.ri) s += `margin-right:${para.ri / 15}px;`;
      if (para.fi) s += `text-indent:${para.fi / 15}px;`;
      s += `margin-top:${para.sb / 15}px;margin-bottom:${para.sa / 15}px;`;
      if (para.sl) {
        if (para.slmult) s += `line-height:${(para.sl / 240).toFixed(3)};`;
        else s += `line-height:${Math.abs(para.sl) / 15}px;`;
      }
      return s + 'white-space:pre-wrap;word-wrap:break-word;';
    };
    const endPara = () => {
      flushBytes();
      let inner = spansHtml();
      if (!inner) inner = `<span style="font-size:${st.fs / 2}pt;">&nbsp;</span>`;
      const html = `<p style="${paraStyle()}">${inner}</p>`;
      spans = []; lastStyle = null;
      if (para.pagebb) { pushBlock('<div class="pgbrk" style="height:0;page-break-before:always"></div>'); para.pagebb = false; }
      if (para.intbl && row) cellParas.push(html); else pushBlock(html);
    };
    const pushBlock = (html) => { if (tableRows.length && !row) closeTable(); out.push(html); };

    // ---- 表格
    const newCell = () => ({ right: 0, b: { t: null, l: null, b: null, r: null }, shd: null, vmgf: false, vmrg: false, hmgf: false, hmrg: false, va: '' });
    let curCell = null, brdSide = null;
    const brdCss = (b) => b ? `${Math.max(1, Math.round((b.w || 10) / 15))}px ${b.style} ${b.color || '#000'}` : 'none';
    function closeCell() {
      flushBytes();
      if (spans.length || !cellParas.length) { para.intbl = true; endPara(); }
      rowCells.push(cellParas.join('')); cellParas = [];
    }
    function closeRow() {
      if (!row) return;
      tableRows.push({ def: row, contents: rowCells });
      rowCells = []; cellParas = []; row = null;
    }
    function closeTable() {
      if (!tableRows.length) return;
      const rows = tableRows; tableRows = [];
      const bounds = rows.map(r => { const left = r.def.left; const b = [left]; r.def.cells.forEach(c => b.push(c.right)); return b; });
      const all = [].concat.apply([], bounds).sort((a, b) => a - b);
      const union = []; for (const v of all) if (!union.length || v - union[union.length - 1] > 20) union.push(v);
      const idx = (v) => { let best = 0, bd = 1e9; union.forEach((u, k) => { const d = Math.abs(u - v); if (d < bd) { bd = d; best = k; } }); return best; };
      const ncol = Math.max(1, union.length - 1);
      // 先把每列的儲存格轉成 {s,e,cell,html}
      const grid = rows.map((r, ri) => {
        const cells = []; const b = bounds[ri];
        r.def.cells.forEach((c, k) => {
          const html = r.contents[k] !== undefined ? r.contents[k] : '<p>&nbsp;</p>';
          if (c.hmrg && cells.length) { cells[cells.length - 1].e = idx(b[k + 1]); return; }
          cells.push({ s: idx(b[k]), e: idx(b[k + 1]), c, html, skip: false, rowspan: 1 });
        });
        return cells;
      });
      // 直向合併
      grid.forEach((cells, ri) => cells.forEach(cell => {
        if (!cell.c.vmgf) return;
        for (let rj = ri + 1; rj < grid.length; rj++) {
          const nxt = grid[rj].find(x => x.s === cell.s && x.c.vmrg);
          if (!nxt) break; nxt.skip = true; cell.rowspan++;
        }
      }));
      let h = `<table style="border-collapse:collapse;table-layout:fixed;width:${(union[union.length - 1] - union[0]) / 15}px;margin-left:${Math.max(0, union[0]) / 15}px;"><colgroup>`;
      for (let k = 0; k < ncol; k++) h += `<col style="width:${(union[k + 1] - union[k]) / 15}px">`;
      h += '</colgroup>';
      rows.forEach((r, ri) => {
        const gap = Math.max(2, Math.round((r.def.gaph || 108) / 15));
        h += `<tr${r.def.trrh ? ` style="height:${Math.abs(r.def.trrh) / 15}px"` : ''}>`;
        grid[ri].forEach(cell => {
          if (cell.skip) return;
          const c = cell.c; let s = `padding:2px ${gap}px;vertical-align:${c.va || 'top'};`;
          s += `border-top:${brdCss(c.b.t)};border-left:${brdCss(c.b.l)};border-bottom:${brdCss(c.b.b)};border-right:${brdCss(c.b.r)};`;
          if (c.shd) s += `background:${c.shd};`;
          h += `<td colspan="${Math.max(1, cell.e - cell.s)}" rowspan="${cell.rowspan}" style="${s}">${cell.html}</td>`;
        });
        h += '</tr>';
      });
      out.push(h + '</table>');
    }

    // ---- 控制字處理
    function control(word, param, hasParam) {
      if (st.first) {
        st.first = false;
        if (word === 'fonttbl') { st.dest = 'fonttbl'; return; }
        if (word === 'colortbl') { st.dest = 'colortbl'; return; }
        if (word === 'pict') { st.dest = 'pict'; pict = { type: '', hex: '', wg: 0, hg: 0, sx: 100, sy: 100 }; return; }
        if (SKIP_DEST.has(word)) { st.dest = 'skip'; return; }
      }
      if (st.dest === 'skip') return;
      if (st.dest === 'colortbl') {
        if (word === 'red') rgb.r = param; else if (word === 'green') rgb.g = param; else if (word === 'blue') rgb.b = param;
        return;
      }
      if (st.dest === 'fonttbl') {
        if (word === 'f') { fontCur = param; fonts[param] = fonts[param] || { name: '', cs: null }; }
        else if (word === 'fcharset' && fontCur !== null) fonts[fontCur].cs = param;
        return;
      }
      if (st.dest === 'pict') {
        if (word === 'pngblip') pict.type = 'png'; else if (word === 'jpegblip') pict.type = 'jpeg';
        else if (word === 'picwgoal') pict.wg = param; else if (word === 'pichgoal') pict.hg = param;
        else if (word === 'picscalex') pict.sx = param; else if (word === 'picscaley') pict.sy = param;
        return;
      }
      switch (word) {
        // 文件層級
        case 'ansicpg': ansicpg = param; break;
        case 'deff': deff = param; st.f = param; break;
        case 'paperw': page.w = param; break; case 'paperh': page.h = param; break;
        case 'margl': page.ml = param; break; case 'margr': page.mr = param; break;
        case 'margt': page.mt = param; break; case 'margb': page.mb = param; break;
        // 字元
        case 'plain': st.b = 0; st.i = 0; st.ul = 0; st.strike = 0; st.fs = 24; st.cf = 0; st.f = deff; st.afs = null; st.af = null; st.ab = null; st.ai = null; break;
        case 'b': flushBytes(); st.b = (hasParam && param === 0) ? 0 : 1; break;
        case 'i': flushBytes(); st.i = (hasParam && param === 0) ? 0 : 1; break;
        case 'ul': case 'uld': case 'uldb': case 'ulw': flushBytes(); st.ul = (hasParam && param === 0) ? 0 : 1; break;
        case 'ulnone': flushBytes(); st.ul = 0; break;
        case 'strike': flushBytes(); st.strike = (hasParam && param === 0) ? 0 : 1; break;
        case 'fs': flushBytes(); st.fs = param || 24; break;
        case 'cf': flushBytes(); st.cf = param; break;
        case 'f': flushBytes(); st.f = param; break;
        case 'afs': flushBytes(); st.afs = param || null; break;
        case 'af': flushBytes(); st.af = param; break;
        case 'ab': flushBytes(); st.ab = (hasParam && param === 0) ? 0 : 1; break;
        case 'ai': flushBytes(); st.ai = (hasParam && param === 0) ? 0 : 1; break;
        case 'uc': st.uc = param; break;
        case 'u': { flushBytes(); const code = param < 0 ? param + 65536 : param; textOut(String.fromCharCode(code)); st.skip = st.uc; break; }
        // 特殊字元
        case 'tab': textOut('\t'); break;
        case 'line': flushBytes(); if (!st.dest) spans.push(`\u0000${cssStyle()}\u0001<br>`); lastStyle = null; break;
        case 'emdash': textOut('—'); break; case 'endash': textOut('–'); break; case 'bullet': textOut('•'); break;
        case 'lquote': textOut('‘'); break; case 'rquote': textOut('’'); break; case 'ldblquote': textOut('“'); break; case 'rdblquote': textOut('”'); break;
        case 'emspace': case 'enspace': textOut(' '); break;
        case 'page': flushBytes(); if (spans.length) endPara(); pushBlock('<div class="pgbrk" style="height:0;page-break-before:always"></div>'); break;
        // 段落
        case 'pard': flushBytes(); resetPara(); break;
        case 'par': case 'nestcell': endPara(); break;
        case 'qc': para.jc = 'center'; break; case 'qr': para.jc = 'right'; break; case 'qj': para.jc = 'justify'; break; case 'ql': para.jc = 'left'; break;
        case 'li': para.li = param; break; case 'fi': para.fi = param; break; case 'ri': para.ri = param; break;
        case 'sb': para.sb = param; break; case 'sa': para.sa = param; break;
        case 'sl': para.sl = param; break; case 'slmult': para.slmult = param; break;
        case 'pagebb': para.pagebb = true; break;
        case 'intbl': para.intbl = true; break;
        // 表格
        case 'trowd': closeRowIfOpen(); row = { left: 0, cells: [], gaph: 108, trrh: 0 }; curCell = newCell(); break;
        case 'trleft': if (row) row.left = param; break;
        case 'trgaph': if (row) row.gaph = param; break;
        case 'trrh': if (row) row.trrh = param; break;
        case 'clbrdrt': brdSide = 't'; if (curCell) curCell.b.t = curCell.b.t || null; break;
        case 'clbrdrl': brdSide = 'l'; break; case 'clbrdrb': brdSide = 'b'; break; case 'clbrdrr': brdSide = 'r'; break;
        case 'brdrnil': case 'brdrnone': if (curCell && brdSide) curCell.b[brdSide] = null; break;
        case 'brdrw': if (curCell && brdSide && curCell.b[brdSide]) curCell.b[brdSide].w = param; break;
        case 'brdrcf': if (curCell && brdSide && curCell.b[brdSide] && colors[param]) curCell.b[brdSide].color = colors[param]; break;
        case 'clcbpat': if (curCell && colors[param]) curCell.shd = colors[param]; break;
        case 'clvmgf': if (curCell) curCell.vmgf = true; break; case 'clvmrg': if (curCell) curCell.vmrg = true; break;
        case 'clmgf': if (curCell) curCell.hmgf = true; break; case 'clmrg': if (curCell) curCell.hmrg = true; break;
        case 'clvertalc': if (curCell) curCell.va = 'middle'; break; case 'clvertalb': if (curCell) curCell.va = 'bottom'; break;
        case 'cellx': if (row && curCell) { curCell.right = param; row.cells.push(curCell); curCell = newCell(); brdSide = null; } break;
        case 'cell': closeCell(); break;
        case 'row': flushBytes(); closeRow(); break;
        default:
          if (BRD_STYLE[word] && curCell && brdSide) curCell.b[brdSide] = { style: BRD_STYLE[word], w: 10, color: '#000' };
      }
    }
    function closeRowIfOpen() { if (row && rowCells.length) closeRow(); }

    // ---- 主迴圈
    while (i < n) {
      const c = s[i];
      if (c === '{') {
        flushBytes();
        stack.push(st); st = Object.assign({}, st); st.first = true; i++;
        // {\* ... } 為可忽略目的地（\shppict 例外：內含圖片要處理）
        if (s[i] === '\\' && s[i + 1] === '*') {
          const m = /^\\\*\s*\\([a-zA-Z]+)/.exec(s.substr(i, 40));
          if (!m || m[1] !== 'shppict') st.dest = 'skip';
          st.first = false;
        }
      } else if (c === '}') {
        flushBytes();
        if (st.dest === 'fonttbl' && fontCur !== null) { /* 字型名稱已累積 */ }
        if (st.dest === 'pict' && pict) {
          if ((pict.type === 'png' || pict.type === 'jpeg') && pict.hex.length > 20) {
            const hex = pict.hex; let bin = ''; for (let k = 0; k + 1 < hex.length; k += 2) bin += String.fromCharCode(parseInt(hex.substr(k, 2), 16));
            const w = pict.wg ? Math.round(pict.wg * pict.sx / 100 / 15) : 0, h = pict.hg ? Math.round(pict.hg * pict.sy / 100 / 15) : 0;
            const parent = stack[stack.length - 1] || st;
            if (!parent.dest || parent.dest === 'pict') {
              if (para.intbl === false && tableRows.length && !row) closeTable();
              spans.push(`\u0000\u0001<img src="data:image/${pict.type};base64,${btoa(bin)}" style="${w ? `width:${w}px;` : ''}${h ? `height:${h}px;` : ''}vertical-align:bottom">`); lastStyle = null;
            }
          }
          pict = null;
        }
        st = stack.pop() || st; i++;
      } else if (c === '\\') {
        const d = s[i + 1];
        if (d === '\\' || d === '{' || d === '}') { if (st.skip > 0) st.skip--; else { flushBytes(); textOut(d); } i += 2; }
        else if (d === "'") {
          const hx = s.substr(i + 2, 2); i += 4;
          if (st.skip > 0) { st.skip--; continue; }
          if (st.dest === 'skip' || st.dest === 'colortbl' || st.dest === 'pict') continue;
          bytes.push(parseInt(hx, 16));
        }
        else if (d === '~') { flushBytes(); textOut(' '); i += 2; }
        else if (d === '_') { flushBytes(); textOut('-'); i += 2; }
        else if (d === '-' || d === '*' || d === ':' || d === '|') { i += 2; }
        else if (d === '\n' || d === '\r') { endPara(); i += 2; }
        else if (/[a-zA-Z]/.test(d)) {
          let j = i + 1; while (j < n && /[a-zA-Z]/.test(s[j])) j++;
          const word = s.slice(i + 1, j);
          let hasParam = false, param = 0;
          if (s[j] === '-' || /\d/.test(s[j] || '')) { let k = j; if (s[k] === '-') k++; while (k < n && /\d/.test(s[k])) k++; param = parseInt(s.slice(j, k), 10); hasParam = true; j = k; }
          if (s[j] === ' ') j++;
          i = j;
          control(word, param, hasParam);
        }
        else { i += 2; }
      } else if (c === '\r' || c === '\n') { i++; }
      else {
        // 一般字元
        if (st.skip > 0) { st.skip--; i++; continue; }
        if (st.dest === 'colortbl') { if (c === ';') { colors.push(rgb.r === null ? null : `rgb(${rgb.r},${rgb.g},${rgb.b})`); rgb = { r: null, g: null, b: null }; } i++; continue; }
        if (st.dest === 'pict') { if (/[0-9a-fA-F]/.test(c)) pict.hex += c; i++; continue; }
        if (st.dest === 'skip') { i++; continue; }
        if (st.dest === 'fonttbl') { flushBytes(); if (c !== ';' && fontCur !== null) fonts[fontCur].name += c; i++; continue; }
        st.first = false;
        // 連續的高位元組字元（沒有用 \' 轉義的 8-bit 文字）
        if (c.charCodeAt(0) >= 0x80) { bytes.push(c.charCodeAt(0)); i++; continue; }
        flushBytes(); textOut(c); i++;
      }
    }
    flushBytes(); if (spans.length) endPara(); closeRowIfOpen(); closeTable();
    return { html: out.join(''), page };
  }

  // ------------------------------------------------------------------ 轉 HTML（統一入口）
  // 回傳 {html(完整文件字串), page:{w,h,ml,mr,mt,mb}(px)}
  function toHtml(u8, kind) {
    const PX = 1 / 15;
    if (kind === 'rtf') {
      const r = rtfToHtml(u8), p = r.page;
      return { html: `<html><head><meta charset="utf-8"></head><body>${r.html}</body></html>`, page: { w: Math.round(p.w * PX), h: Math.round(p.h * PX), ml: Math.round(p.ml * PX), mr: Math.round(p.mr * PX), mt: Math.round(p.mt * PX), mb: Math.round(p.mb * PX) } };
    }
    if (kind === 'mht') return { html: mhtToHtml(u8), page: {} };
    if (kind === 'text') return { html: `<html><head><meta charset="utf-8"></head><body><pre style="font-family:'新細明體','PMingLiU',monospace;font-size:12pt;white-space:pre-wrap;margin:0">${escHtml(decodeText(u8))}</pre></body></html>`, page: {} };
    return { html: decodeText(u8), page: {} };   // html
  }

  // ------------------------------------------------------------------ HTML → 分頁 JPG
  // @page 設定（Word 另存網頁常見）：size / margin
  function pageFromCss(html) {
    const out = {}; const m = html.match(/@page[^{]*\{([^}]*)\}/i); if (!m) return out;
    const unit = (v, u) => { const x = parseFloat(v); return u === 'cm' ? x * 37.795 : u === 'mm' ? x * 3.7795 : u === 'in' ? x * 96 : u === 'pt' ? x * 96 / 72 : x; };
    const sz = m[1].match(/size:\s*([\d.]+)(cm|mm|in|pt)\s+([\d.]+)(cm|mm|in|pt)/i);
    if (sz) { out.w = Math.round(unit(sz[1], sz[2])); out.h = Math.round(unit(sz[3], sz[4])); }
    const mg = m[1].match(/margin:\s*([\d.]+)(cm|mm|in|pt)(?:\s+([\d.]+)(cm|mm|in|pt))?(?:\s+([\d.]+)(cm|mm|in|pt))?(?:\s+([\d.]+)(cm|mm|in|pt))?/i);
    if (mg) {
      const t = unit(mg[1], mg[2]), r = mg[3] ? unit(mg[3], mg[4]) : t, b = mg[5] ? unit(mg[5], mg[6]) : t, l = mg[7] ? unit(mg[7], mg[8]) : r;
      out.mt = Math.round(t); out.mr = Math.round(r); out.mb = Math.round(b); out.ml = Math.round(l);
    }
    return out;
  }

  async function renderPages(htmlStr, pageOpt, scale, toJpg, onProgress) {
    const P = Object.assign({ w: 794, h: 1123, ml: 57, mr: 57, mt: 57, mb: 57 }, pageFromCss(htmlStr), pageOpt || {});
    const parsed = new DOMParser().parseFromString(htmlStr, 'text/html');
    parsed.querySelectorAll('script,iframe,object,embed,link,meta[http-equiv]').forEach(e => e.remove());
    const meta = parsed.createElement('meta'); meta.setAttribute('http-equiv', 'Content-Security-Policy'); meta.setAttribute('content', "default-src 'none'; img-src data:; style-src 'unsafe-inline'");
    parsed.head.insertBefore(meta, parsed.head.firstChild);
    const base = parsed.createElement('style');
    const bodyW = P.w - P.ml - P.mr;
    base.textContent = `html,body{margin:0;background:#fff;} body{box-sizing:content-box;width:${bodyW}px;padding:0 ${P.mr}px 0 ${P.ml}px;font-family:'新細明體','PMingLiU','Microsoft JhengHei',sans-serif;font-size:12pt;overflow-wrap:break-word;} img{max-width:100%;} table{max-width:none;}`;
    parsed.head.appendChild(base);

    const iframe = document.createElement('iframe');
    iframe.setAttribute('sandbox', 'allow-same-origin');
    iframe.style.cssText = `position:fixed;left:-30000px;top:0;width:${P.w}px;height:${P.h}px;border:0;`;
    document.body.appendChild(iframe);
    try {
      await new Promise((res) => { iframe.onload = res; iframe.srcdoc = '<!doctype html>' + parsed.documentElement.outerHTML; });
      const doc = iframe.contentDocument, body = doc.body;
      if (doc.fonts && doc.fonts.ready) await doc.fonts.ready;
      // 內容比頁面寬（固定寬度表格）：放大頁面寬度以免被截掉
      const sw = Math.max(doc.documentElement.scrollWidth, body.scrollWidth);
      if (sw > P.w + 2) { const ratio = P.h / P.w; P.w = sw; P.h = Math.round(sw * ratio); iframe.style.width = P.w + 'px'; await new Promise(r => setTimeout(r, 50)); }
      const total = Math.ceil(body.scrollHeight);
      const contentH = P.h - P.mt - P.mb;
      if (!(contentH > 100)) throw new Error('頁面尺寸異常');

      // 強制換頁位置 + 可安全切頁的位置（區塊底部，避免切到文字行中間）
      const top0 = body.getBoundingClientRect().top;
      const forced = new Set();
      doc.querySelectorAll('*').forEach(el => {
        const cs = doc.defaultView.getComputedStyle(el);
        const before = cs.breakBefore === 'page' || cs.pageBreakBefore === 'always', after = cs.breakAfter === 'page' || cs.pageBreakAfter === 'always';
        if (before || after) { const r = el.getBoundingClientRect(); if (before) forced.add(Math.round(r.top - top0)); if (after) forced.add(Math.round(r.bottom - top0)); }
      });
      const safe = [];
      doc.querySelectorAll('p,tr,li,h1,h2,h3,h4,h5,h6,div,img,pre,blockquote').forEach(el => {
        if (el.tagName === 'DIV' && el.querySelector('p,tr,li,div,table,h1,h2,h3,h4,h5,h6')) return;
        safe.push(Math.round(el.getBoundingClientRect().bottom - top0));
      });
      safe.sort((a, b) => a - b);
      const forcedArr = [...forced].filter(v => v > 0 && v < total).sort((a, b) => a - b);

      const cuts = []; let y0 = 0;
      while (y0 < total - 2) {
        const target = y0 + contentH;
        const f = forcedArr.find(v => v > y0 + 2 && v <= target);
        let cut;
        if (f !== undefined) cut = f;
        else if (target >= total) cut = total;
        else { let best = null; for (const v of safe) { if (v > y0 + contentH * 0.4 && v <= target) best = v; } cut = best !== null ? best : target; }
        cuts.push([y0, cut]); y0 = cut;
      }
      if (!cuts.length) cuts.push([0, total]);

      const results = [];
      for (let k = 0; k < cuts.length; k++) {
        const [a, b] = cuts[k];
        if (onProgress) onProgress(k + 1, cuts.length);
        const slice = await window.html2canvas(body, { scale, backgroundColor: '#ffffff', logging: false, x: 0, y: a, width: P.w, height: b - a, windowWidth: P.w, windowHeight: total, scrollX: 0, scrollY: 0 });
        const pc = document.createElement('canvas'); pc.width = Math.round(P.w * scale); pc.height = Math.round(P.h * scale);
        const ctx = pc.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, pc.width, pc.height);
        ctx.drawImage(slice, 0, Math.round(P.mt * scale));
        results.push(await toJpg(pc));
      }
      return results;
    } finally { iframe.remove(); }
  }

  window.LegacyDoc = { sniff, toHtml, renderPages, _rtfToHtml: rtfToHtml, _decodeText: decodeText };
})();

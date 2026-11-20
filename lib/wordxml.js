/* Word 2003 XML (WordprocessingML 2003) → .docx（全程在瀏覽器記憶體內，需先載入 JSZip）
 * 支援：段落/字元格式、樣式、表格(合併/框線/底色)、內嵌圖片(w:binData)、項目符號標籤文字、分頁/分節頁面設定。
 * 不支援：頁首頁尾、文字方塊、向量圖形(v:rect 等)。
 */
(function () {
  const NS_W = 'http://schemas.microsoft.com/office/word/2003/wordml';
  const NS_WX = 'http://schemas.microsoft.com/office/word/2003/auxHint';
  const NS_V = 'urn:schemas-microsoft-com:vml';
  const OOXML_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const camel = s => s.replace(/-([a-z])/g, (_, c) => c.toUpperCase());
  const ELEM_FIX = { vmerge: 'vMerge', hmerge: 'hMerge', gridspan: 'gridSpan', 'h-ansi': 'hAnsi' };
  const ATTR_FIX = { fareast: 'eastAsia', 'h-ansi': 'hAnsi', hansi: 'hAnsi' };
  const SKIP = new Set(['rsid', 'rsidR', 'rsidRPr', 'rsidRDefault', 'rsidP', 'rsidTr', 'proofErr', 'permStart', 'permEnd', 'bookmarkStart', 'bookmarkEnd', 'lang', 'noProof', 'latentStyles', 'listPr', 'hdr', 'ftr', 'footnote', 'endnote', 'fldChar']);

  function lenToEmu(v) {
    const m = /^(-?[\d.]+)\s*(pt|in|cm|mm|px)?$/.exec((v || '').trim());
    if (!m) return 0;
    const n = parseFloat(m[1]), u = m[2] || 'px';
    return Math.round(n * ({ pt: 12700, in: 914400, cm: 360000, mm: 36000, px: 9525 }[u]));
  }
  function styleVal(style, key) {
    const m = new RegExp('(?:^|;)\\s*' + key + '\\s*:\\s*([^;]+)').exec(style || '');
    return m ? m[1] : '';
  }

  // 由 base64 讀 PNG / JPEG / GIF / BMP 的像素尺寸
  function pixelSize(b64) {
    if (!b64) return null;
    try {
      const bin = atob(b64.slice(0, 80000)), u = i => bin.charCodeAt(i);
      if (u(0) === 0x89 && bin.slice(1, 4) === 'PNG') return { w: (u(16) << 24 | u(17) << 16 | u(18) << 8 | u(19)), h: (u(20) << 24 | u(21) << 16 | u(22) << 8 | u(23)) };
      if (bin.slice(0, 3) === 'GIF') return { w: u(6) | u(7) << 8, h: u(8) | u(9) << 8 };
      if (bin.slice(0, 2) === 'BM') return { w: u(18) | u(19) << 8, h: Math.abs(u(22) | u(23) << 8) };
      if (u(0) === 0xFF && u(1) === 0xD8) {
        let i = 2;
        while (i < bin.length - 9) {
          if (u(i) !== 0xFF) { i++; continue; }
          const m = u(i + 1);
          if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) return { h: u(i + 5) << 8 | u(i + 6), w: u(i + 7) << 8 | u(i + 8) };
          i += 2 + (u(i + 2) << 8 | u(i + 3));
        }
      }
    } catch (e) { /* ignore */ }
    return null;
  }

  async function toDocx(u8) {
    let text = new TextDecoder('utf-8').decode(u8);
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('XML 格式錯誤，無法解析');
    const root = doc.documentElement;
    if (root.namespaceURI !== NS_W || root.localName !== 'wordDocument') throw new Error('不是 Word 2003 XML 文件');

    // 內嵌圖片 (w:binData)
    const bins = {};
    for (const b of doc.getElementsByTagNameNS(NS_W, 'binData')) {
      const name = b.getAttributeNS(NS_W, 'name') || b.getAttribute('w:name');
      if (name) bins[name] = (b.textContent || '').replace(/\s+/g, '');
    }
    const media = [];     // {name, b64}
    const imgRid = (src) => {
      const b64 = bins[src];
      if (!b64) return null;
      const ext = (/\.(png|jpe?g|gif|bmp)$/i.exec(src) || [, 'png'])[1].toLowerCase().replace('jpg', 'jpeg');
      let idx = media.findIndex(m => m.src === src);
      if (idx < 0) { media.push({ src, name: 'image' + (media.length + 1) + '.' + ext, b64 }); idx = media.length - 1; }
      return { rid: 'rIdImg' + (idx + 1), ext };
    };

    let skippedPics = 0, drawId = 0;

    function attrs(el) {
      let s = '';
      for (const a of el.attributes) {
        if (a.namespaceURI === NS_W) {
          const n = ATTR_FIX[a.localName] || camel(a.localName);
          if (/^rsid/.test(n)) continue;
          s += ` w:${n}="${esc(a.value)}"`;
        }
      }
      return s;
    }

    function pict(el) {
      let out = '';
      for (const sh of el.getElementsByTagNameNS(NS_V, 'shape')) {
        const im = sh.getElementsByTagNameNS(NS_V, 'imagedata')[0];
        if (!im) continue;
        const r = imgRid(im.getAttribute('src') || '');
        if (!r) { skippedPics++; continue; }
        const st = sh.getAttribute('style') || '';
        let w = lenToEmu(styleVal(st, 'width')), h = lenToEmu(styleVal(st, 'height'));
        if (!w || !h) {                                   // 尺寸缺失：以圖片原始像素尺寸（96dpi）補上，並依比例補另一邊
          const d = pixelSize(bins[im.getAttribute('src') || '']);
          if (d) {
            if (w && !h) h = Math.round(w * d.h / d.w); else if (h && !w) w = Math.round(h * d.w / d.h);
            else { w = d.w * 9525; h = d.h * 9525; }
          } else { w = w || 2540000; h = h || 1905000; }
        }
        const id = ++drawId;
        out += `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${w}" cy="${h}"/><wp:docPr id="${id}" name="Picture ${id}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${r.rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
      }
      if (!out) skippedPics++;
      return out;
    }

    function conv(el) {
      let out = '';
      for (const n of el.childNodes) {
        if (n.nodeType === 3) continue;                      // 元素間的空白
        if (n.nodeType !== 1) continue;
        if (n.namespaceURI === NS_WX) {                       // wx:sect / wx:sub-section 等：只展開子元素
          if (n.localName === 'borders') continue;
          out += conv(n); continue;
        }
        if (n.namespaceURI !== NS_W) continue;
        const ln = n.localName;
        if (SKIP.has(ln)) continue;
        if (ln === 'sectPr') { out += sect(n); continue; }
        if (ln === 'sym') {                                   // Wingdings 等符號字元（如 □ 勾選框）
          const code = parseInt((n.getAttributeNS(NS_W, 'char') || '0'), 16) & 0xFF;
          const MAP = { 0x6F: '□', 0xA8: '□', 0x71: '❑', 0x6E: '■', 0xA7: '▪', 0xFE: '☑', 0xFD: '☒', 0x78: '☒', 0xFC: '✓', 0xFB: '✗', 0xA1: '○', 0x6C: '●', 0xD8: '➢', 0x76: '❖', 0x75: '◆', 0xE0: '→' };
          out += `<w:t xml:space="preserve">${MAP[code] || '□'}</w:t>`; continue;
        }
        if (ln === 'pict') { out += pict(n); continue; }
        if (ln === 'cr') { out += '<w:br/>'; continue; }
        if (ln === 'p') { out += paragraph(n); continue; }
        if (ln === 't') { out += `<w:t xml:space="preserve">${esc(n.textContent)}</w:t>`; continue; }
        const name = ELEM_FIX[ln] || camel(ln);
        const inner = conv(n);
        out += inner ? `<w:${name}${attrs(n)}>${inner}</w:${name}>` : `<w:${name}${attrs(n)}/>`;
      }
      return out;
    }

    function paragraph(p) {
      // 項目符號/編號：w:pPr/w:listPr/wx:t 的文字標籤補在段落開頭
      let label = '';
      const pPr = [...p.children].find(c => c.namespaceURI === NS_W && c.localName === 'pPr');
      if (pPr) {
        const lp = [...pPr.children].find(c => c.localName === 'listPr' && c.namespaceURI === NS_W);
        if (lp) {
          const t = lp.getElementsByTagNameNS(NS_WX, 't')[0];
          if (t) label = t.getAttributeNS(NS_WX, 'val') || t.getAttribute('wx:val') || '';
        }
      }
      let body = conv(p);
      body = applyGrid(body);
      if (label) body = body.replace(/^(<w:pPr>[\s\S]*?<\/w:pPr>)?/, (m) => m + `<w:r><w:t xml:space="preserve">${esc(label)}  </w:t></w:r>`);
      return `<w:p>${body}</w:p>`;
    }

    // 文件格線（docGrid type=lines）：Word 會讓每行高度成為格線間距的倍數，空白行與單行文字因此比一般行高
    let gridPitch = 0;
    for (const g of doc.getElementsByTagNameNS(NS_W, 'docGrid')) {
      if (g.getAttributeNS(NS_W, 'type') === 'lines' && +g.getAttributeNS(NS_W, 'line-pitch') > 0) { gridPitch = +g.getAttributeNS(NS_W, 'line-pitch'); break; }
    }
    function applyGrid(body) {
      if (!gridPitch) return body;
      const m = /^<w:pPr>[\s\S]*?<\/w:pPr>/.exec(body);
      const pPr = m ? m[0] : '';
      if (/<w:snapToGrid[^>]*w:val="off"/.test(pPr)) return body;
      const sp = /<w:spacing(?![A-Za-z])[^>]*>/.exec(pPr);
      let np;
      if (sp) {
        if (/ w:line="/.test(sp[0])) return body;                       // 作者已指定行距
        np = pPr.replace(sp[0], sp[0].replace('<w:spacing', `<w:spacing w:line="${gridPitch}" w:lineRule="atLeast"`));
      } else if (pPr) np = pPr.replace('<w:pPr>', `<w:pPr><w:spacing w:line="${gridPitch}" w:lineRule="atLeast"/>`);
      else return `<w:pPr><w:spacing w:line="${gridPitch}" w:lineRule="atLeast"/></w:pPr>` + body;
      return np + body.slice(pPr.length);
    }

    // 分節與頁尾
    const footers = [];          // {rid, xml}
    function sect(sp) {
      const g = (e, k, d) => (e && e.getAttributeNS(NS_W, k)) || d;
      const kids = [...sp.children].filter(c => c.namespaceURI === NS_W);
      const one = (n) => kids.find(c => c.localName === n);
      let x = '';
      const ftr = kids.find(c => c.localName === 'ftr' && (g(c, 'type', 'odd') === 'odd'));
      if (ftr) {
        const rid = 'rIdFtr' + (footers.length + 1);
        footers.push({ rid, xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="${OOXML_W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">${conv(ftr)}</w:ftr>` });
        x += `<w:footerReference w:type="default" r:id="${rid}"/>`;
      }
      const ty = one('type'); if (ty) x += `<w:type w:val="${g(ty, 'val', 'nextPage')}"/>`;
      const pg = one('pgSz'), mar = one('pgMar');
      x += `<w:pgSz w:w="${g(pg, 'w', 11906)}" w:h="${g(pg, 'h', 16838)}"${g(pg, 'orient', '') === 'landscape' ? ' w:orient="landscape"' : ''}/>`;
      x += `<w:pgMar w:top="${g(mar, 'top', 1440)}" w:right="${g(mar, 'right', 1440)}" w:bottom="${g(mar, 'bottom', 1440)}" w:left="${g(mar, 'left', 1440)}" w:header="${g(mar, 'header', 720)}" w:footer="${g(mar, 'footer', 720)}" w:gutter="${g(mar, 'gutter', 0)}"/>`;
      return `<w:sectPr>${x}</w:sectPr>`;
    }

    // 內文
    const bodyEl = doc.getElementsByTagNameNS(NS_W, 'body')[0];
    if (!bodyEl) throw new Error('找不到文件內文');
    let bodyXml = conv(bodyEl);
    // 若最後一節沒有任何內容（緊接在上一個分節符之後），改以上一節的設定作為文件最後一節，避免多出空白頁
    {
      const fin = bodyXml.lastIndexOf('<w:sectPr>');
      if (fin > 0 && bodyXml.endsWith('</w:sectPr>')) {
        const prev = bodyXml.lastIndexOf('<w:sectPr>', fin - 1);
        const prevEnd = prev >= 0 ? bodyXml.indexOf('</w:sectPr>', prev) + '</w:sectPr>'.length : -1;
        if (prev >= 0 && /^(<\/w:pPr>)?<\/w:p>$/.test(bodyXml.slice(prevEnd, fin))) {
          const prevSect = bodyXml.slice(prev, prevEnd);
          bodyXml = bodyXml.slice(0, prev) + bodyXml.slice(prevEnd, fin) + prevSect;
        }
      }
    }
    if (!bodyXml.endsWith('</w:sectPr>')) bodyXml += '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    const sect_ = '';

    // 樣式
    let stylesInner = '';
    const stEl = doc.getElementsByTagNameNS(NS_W, 'styles')[0];
    if (stEl) stylesInner = conv(stEl);
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${OOXML_W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>${stylesInner}</w:styles>`;

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${OOXML_W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${bodyXml}${sect_}</w:body></w:document>`;

    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="bmp" ContentType="image/bmp"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' + footers.map((f, i) => `<Override PartName="/word/footer${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`).join('') + '</Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      media.map((m, i) => `<Relationship Id="rIdImg${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${m.name}"/>`).join('') +
      footers.map((f, i) => `<Relationship Id="${f.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer${i + 1}.xml"/>`).join('') + '</Relationships>');
    const imgRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      media.map((m, i) => `<Relationship Id="rIdImg${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${m.name}"/>`).join('') + '</Relationships>';
    footers.forEach((f, i) => { zip.file(`word/footer${i + 1}.xml`, f.xml); zip.file(`word/_rels/footer${i + 1}.xml.rels`, imgRels); });
    zip.file('word/styles.xml', stylesXml);
    zip.file('word/document.xml', documentXml);
    media.forEach(m => zip.file('word/media/' + m.name, m.b64, { base64: true }));
    const out = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { docx: out, info: { pics: media.length, skippedPics } };
  }

  window.WordXml = { toDocx };
})();

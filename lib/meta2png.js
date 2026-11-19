/* WMF / EMF → PNG（全程在瀏覽器本機）。依賴 lib/wmf.min.js、lib/emf.min.js（rtf.js，MIT）。 */
(function () {
  function toArrayBuffer(u8) { return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength); }
  window.metaToPng = async function (data, ext, wTw, hTw, blip) {
    const Lib = ext === 'emf' ? window.EMFJS : window.WMFJS;
    if (!Lib) return null;
    let wPx = Math.round((wTw || 0) / 15), hPx = Math.round((hTw || 0) / 15);
    if ((!wPx || !hPx) && blip && blip.ptW) { wPx = Math.round(blip.ptW / 9525); hPx = Math.round(blip.ptH / 9525); }
    if (!wPx || !hPx) { wPx = 400; hPx = 300; }
    const ab = toArrayBuffer(data);
    const r = new Lib.Renderer(ab);
    let xe = wPx, ye = hPx;
    if (ext === 'emf' && data.length >= 40) {           // EMF 標頭 rclFrame（0.01 mm）= 邏輯尺寸
      const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
      const fw = dv.getInt32(32, true) - dv.getInt32(24, true), fh = dv.getInt32(36, true) - dv.getInt32(28, true);
      if (fw > 0 && fh > 0) { xe = fw; ye = fh; }
    }
    const svg = r.render({ width: wPx + 'px', height: hPx + 'px', wExt: xe, hExt: ye, xExt: xe, yExt: ye, mapMode: 8 });
    if (!svg.getAttribute('xmlns')) svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    // 以實際內容外框重設 viewBox，避免各家產生的 metafile 座標系不同而裁切／縮小
    try {
      const holder = document.createElement('div');
      holder.style.cssText = 'position:fixed;left:-99999px;top:0;visibility:hidden';
      holder.appendChild(svg); document.body.appendChild(holder);
      const bb = svg.getBBox();
      holder.remove();
      if (bb.width > 0 && bb.height > 0) svg.setAttribute('viewBox', [bb.x, bb.y, bb.width, bb.height].join(' '));
    } catch (e) { /* 保持原 viewBox */ }
    svg.setAttribute('width', wPx); svg.setAttribute('height', hPx);
    svg.setAttribute('preserveAspectRatio', 'none');
    const xml = new XMLSerializer().serializeToString(svg);
    const url = URL.createObjectURL(new Blob([xml], { type: 'image/svg+xml;charset=utf-8' }));
    try {
      const img = new Image();
      await new Promise((res, rej) => { img.onload = res; img.onerror = () => rej(new Error('svg')); img.src = url; });
      const S = 2, c = document.createElement('canvas');
      c.width = wPx * S; c.height = hPx * S;
      const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
      g.drawImage(img, 0, 0, c.width, c.height);
      const blob = await new Promise(res => c.toBlob(res, 'image/png'));
      return new Uint8Array(await blob.arrayBuffer());
    } finally { URL.revokeObjectURL(url); }
  };
})();

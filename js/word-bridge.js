// Word のアドインとして開かれたとき(?addin=dialog)だけ働く。
// Word の作業ウィンドウ(addin/taskpane.js)から文書を受け取り、「Wordに反映」で送り返す。
(function(){
  'use strict';
  if(!/[?&]addin=dialog(&|$)/.test(location.search)) return;
  const editor = window.KouriEditor;
  if(!editor) return;


  // ---------- ページを画像(PNG)にする ----------
  // 字雲フォントで描く漢字は canvas で、それ以外(仮名・句読点・括弧・数字など)は、
  // エディタの書体の指定のまま SVG の縦書きの文字として描いて、重ね合わせる。
  const JIGMO = (cp) => (cp >= 0x2E80 && cp <= 0x2FFF) || (cp >= 0x3005 && cp <= 0x3007) || (cp >= 0x31C0 && cp <= 0x31EF)
    || (cp >= 0x3400 && cp <= 0x4DBF) || (cp >= 0x4E00 && cp <= 0x9FFF) || (cp >= 0xF900 && cp <= 0xFAFF) || (cp >= 0x20000 && cp <= 0x3FFFF);
  const BASE_INK = 'rgb(43, 42, 40)';
  const SEL_COLORS = ['rgb(236, 58, 80)', 'rgb(51, 88, 240)'];  // 選択・キャレット直前の色は、出力には使わない
  const xesc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  async function renderPageToPng(page, scale){
    const pr = page.getBoundingClientRect();
    const ratio = pr.width / page.offsetWidth || 1;   // 表示倍率
    const W = page.offsetWidth, H = page.offsetHeight;
    const rel = (r) => ({ left: (r.left - pr.left) / ratio, top: (r.top - pr.top) / ratio, width: r.width / ratio, height: r.height / ratio });
    const svg = [];
    const kanji = [];
    const els = page.querySelectorAll('*');
    for(const el of els){
      if(el.closest('.gaiji-tint') || (el.closest('svg') && el.tagName.toLowerCase() !== 'svg')) continue;
      const cs = getComputedStyle(el);
      if(cs.display === 'none' || cs.visibility === 'hidden' || Number(cs.opacity) === 0) continue;
      let color = cs.color;
      if(SEL_COLORS.indexOf(color) >= 0) color = BASE_INK;
      const tag = el.tagName.toLowerCase();
      if(tag === 'img'){
        const r = rel(el.getBoundingClientRect());
        if(r.width > 0 && el.src) svg.push('<image href="' + xesc(el.src) + '" x="' + r.left + '" y="' + r.top + '" width="' + r.width + '" height="' + r.height + '"/>');
        continue;
      }
      if(tag === 'svg'){
        const r = rel(el.getBoundingClientRect());
        const c = el.cloneNode(true);
        c.setAttribute('x', r.left); c.setAttribute('y', r.top); c.setAttribute('width', r.width); c.setAttribute('height', r.height);
        c.removeAttribute('style');
        let str = new XMLSerializer().serializeToString(c).replace(/currentColor/g, color);
        svg.push(str);
        continue;
      }
      const vertical = cs.writingMode.indexOf('vertical') === 0;
      const sw = (parseFloat(cs.webkitTextStrokeWidth) || 0) / ratio;  // 字の縁取り(太さの出し方)
      for(const node of el.childNodes){
        if(node.nodeType !== 3 || !node.nodeValue) continue;
        const text = node.nodeValue;
        for(let k = 0; k < text.length; ){
          const cp = text.codePointAt(k), ch = String.fromCodePoint(cp), len = ch.length;
          const rg = document.createRange();
          rg.setStart(node, k); rg.setEnd(node, k + len);
          const rect = rg.getBoundingClientRect();
          k += len;
          if(!rect.width && !rect.height) continue;
          if(/\s/.test(ch)) continue;
          const r = rel(rect);
          const fs = parseFloat(cs.fontSize) / ratio * (ratio); // 画面上の大きさではなく、ページ内の大きさ
          if(JIGMO(cp)){
            kanji.push({ ch, cx: r.left + r.width / 2, cy: r.top + r.height / 2, size: fs, color, weight: cs.fontWeight, family: cs.fontFamily, sw });
          } else if(vertical){
            svg.push('<text x="' + (r.left + r.width / 2) + '" y="' + r.top + '" writing-mode="vertical-rl" font-size="' + fs + '" font-weight="' + cs.fontWeight
              + '" font-family="' + xesc(cs.fontFamily) + '" fill="' + color + '"' + (sw ? ' stroke="' + color + '" stroke-width="' + sw + '"' : '') + '>' + xesc(ch) + '</text>');
          } else {
            svg.push('<text x="' + r.left + '" y="' + (r.top + r.height / 2) + '" dominant-baseline="central" font-size="' + fs + '" font-weight="' + cs.fontWeight
              + '" font-family="' + xesc(cs.fontFamily) + '" fill="' + color + '"' + (sw ? ' stroke="' + color + '" stroke-width="' + sw + '"' : '') + '>' + xesc(ch) + '</text>');
          }
        }
      }
    }
    const svgText = '<svg xmlns="http://www.w3.org/2000/svg" width="' + (W * scale) + '" height="' + (H * scale) + '" viewBox="0 0 ' + W + ' ' + H + '">'
      + '<rect width="' + W + '" height="' + H + '" fill="#fff"/>' + svg.join('') + '</svg>';
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(W * scale); canvas.height = Math.round(H * scale);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
    const img = new Image();
    await new Promise((resolve, reject) => {
      img.onload = resolve; img.onerror = () => reject(new Error('ページの画像を作れませんでした'));
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svgText);
    });
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    // 漢字は、字雲フォントで描く(読み込みを待ってから)
    ctx.setTransform(scale, 0, 0, scale, 0, 0);
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    const loads = new Set();
    kanji.forEach(g => loads.add(g.weight + ' ' + g.size + 'px ' + g.family + '\u0000' + g.ch));
    await Promise.all([...loads].map(k => { const [f, c] = k.split('\u0000'); return document.fonts.load(f, c).catch(() => {}); }));
    kanji.forEach(g => {
      ctx.font = g.weight + ' ' + g.size + 'px ' + g.family;
      ctx.fillStyle = g.color;
      ctx.fillText(g.ch, g.cx, g.cy);
      if(g.sw){ ctx.lineWidth = g.sw; ctx.strokeStyle = g.color; ctx.lineJoin = 'round'; ctx.strokeText(g.ch, g.cx, g.cy); }
    });
    return { dataUrl: canvas.toDataURL('image/png'), width: W, height: H };
  }

  async function renderAllPages(){
    if(document.fonts && document.fonts.ready) await document.fonts.ready;
    const pages = [...document.querySelectorAll('.page')];
    const out = [];
    for(const p of pages) out.push(await renderPageToPng(p, 2));
    return out;
  }

  const bar = document.createElement('div');
  bar.id = 'word-bar';
  bar.setAttribute('role', 'region');
  bar.setAttribute('aria-label', 'Word との受け渡し');
  bar.style.cssText = 'position:fixed;left:50%;bottom:12px;transform:translateX(-50%);z-index:9999;display:flex;gap:8px;align-items:center;'
    + 'padding:8px 12px;background:#fff;border:1px solid #767676;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.25);font:14px/1.4 "Noto Sans JP",sans-serif;color:#1a1a1c;';
  const msg = document.createElement('span');
  msg.textContent = 'Word の本文を読み込んでいます…';
  const mk = (label, primary) => {
    const b = document.createElement('button');
    b.type = 'button'; b.textContent = label; b.disabled = true;
    b.style.cssText = 'padding:6px 14px;border-radius:6px;font:inherit;cursor:pointer;border:1px solid #0017C1;'
      + (primary ? 'background:#0017C1;color:#fff;' : 'background:#fff;color:#0017C1;');
    return b;
  };
  const applyImg = mk('Wordに反映(見た目そのまま・画像)', true);
  const applyTxt = mk('Wordに反映(文字・ルビ)', false);
  const cancel = mk('反映せずに閉じる', false);
  const sizeSel = document.createElement('select');
  sizeSel.title = 'Word に入れる画像の幅';
  sizeSel.style.cssText = 'padding:6px;border-radius:6px;font:inherit;border:1px solid #767676;';
  [['縦向きの用紙(幅 15.5cm)', 440], ['横向きの用紙(幅 24cm)', 690]].forEach(([l, v]) => { const o = document.createElement('option'); o.value = v; o.textContent = l; sizeSel.appendChild(o); });
  bar.append(msg, applyImg, applyTxt, sizeSel, cancel);
  document.body.appendChild(bar);
  const setBusy = (b) => { applyImg.disabled = b; applyTxt.disabled = b; cancel.disabled = b; };

  const post = (obj) => { try{ Office.context.ui.messageParent(JSON.stringify(obj)); } catch(err){ msg.textContent = 'Word と通信できません:' + err.message; } };

  applyTxt.addEventListener('click', () => {
    setBusy(true);
    msg.textContent = 'Word に反映しています…';
    post({ type: 'apply', doc: editor.getDocument() });
  });
  applyImg.addEventListener('click', async () => {
    setBusy(true);
    try{
      msg.textContent = 'ページを画像にしています…';
      const pages = await renderAllPages();
      msg.textContent = 'Word に送っています…';
      post({ type: 'img-start', doc: editor.getDocument(), count: pages.length, width: Number(sizeSel.value), sizes: pages.map(p => [p.width, p.height]) });
      // 画像が大きいので、小分けにして送る
      const CH = 60000;
      for(let i = 0; i < pages.length; i++){
        const b64 = pages[i].dataUrl.replace(/^data:image\/png;base64,/, '');
        const parts = Math.ceil(b64.length / CH);
        for(let j = 0; j < parts; j++) post({ type: 'img-chunk', page: i, part: j, parts, data: b64.slice(j * CH, (j + 1) * CH) });
      }
      post({ type: 'img-end' });
    } catch(err){
      msg.textContent = '画像にできませんでした:' + err.message;
      setBusy(false);
    }
  });
  cancel.addEventListener('click', () => post({ type: 'close' }));

  const s = document.createElement('script');
  s.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
  s.onerror = () => { msg.textContent = 'Office.js を読み込めませんでした。ネットワークを確認してください。'; };
  s.onload = () => {
    Office.onReady(() => {
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (e) => {
        let m; try{ m = JSON.parse(e.message); } catch(err){ return; }
        if(m.type !== 'load') return;
        editor.loadDocument(m.doc);
        editor.markSaved();
        msg.textContent = 'Word の本文を編集しています。';
        setBusy(false);
      });
      post({ type: 'ready' });
    });
  };
  document.head.appendChild(s);
})();

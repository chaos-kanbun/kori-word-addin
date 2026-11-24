// 交利の文書(JSON)を、Word に挿入できる OOXML(段落の並び)にする。
// ブラウザ(作業ウィンドウ)と Node.js(テスト)の両方で使えます。
(function(root){
  'use strict';
  const PUA_MIN = 0xF0000;
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const FONT = '<w:rFonts w:ascii="MS Mincho" w:hAnsi="MS Mincho" w:eastAsia="MS Mincho"/>';

  // 本文の字の並びから、Word の本文に置く文字列を作る(外字は 〓 にする)
  function plainChar(ch){
    const cp = ch.codePointAt(0);
    return cp >= PUA_MIN ? '〓' : ch;
  }

  function run(text, opt){
    opt = opt || {};
    let rpr = FONT;
    if(opt.size) rpr += '<w:sz w:val="' + opt.size + '"/><w:szCs w:val="' + opt.size + '"/>';
    if(opt.u) rpr += '<w:u w:val="' + opt.u + '"/>';
    return '<w:r><w:rPr>' + rpr + '</w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
  }

  // 振り仮名(ルビ)。base は親文字の run 列、rt は振り仮名の文字列
  function ruby(baseRuns, rt){
    return '<w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="10"/><w:hpsRaise w:val="20"/>'
      + '<w:hpsBaseText w:val="21"/><w:lid w:val="ja-JP"/></w:rubyPr>'
      + '<w:rt><w:r><w:rPr>' + FONT + '<w:sz w:val="10"/></w:rPr><w:t>' + esc(rt) + '</w:t></w:r></w:rt>'
      + '<w:rubyBase>' + baseRuns + '</w:rubyBase></w:ruby></w:r>';
  }

  // 文書(JSON)→ 段落の OOXML(<w:p>…</w:p> の並び)
  // 対応:本文、振り仮名・熟字訓(ルビ)、送り仮名・訓点・左の振り仮名/送り仮名(小さい字で字の後ろに続ける)、傍線・波線(下線)
  function buildParagraphsXml(doc){
    const text = typeof doc.text === 'string' ? doc.text : '';
    const annos = new Map((doc.annotations || []).map(([k, v]) => [Number(k), v || {}]));
    const spans = (doc.jukujikunSpans || []).filter(s => s && s.end > s.start);
    const spanAt = new Map(spans.map(s => [s.start, s]));
    const lines = doc.lineSpans || [];
    const underline = (i) => {
      for(const l of lines){ if(i >= l.start && i < l.end) return l.type === 'wavy' ? 'wave' : 'single'; }
      return '';
    };
    const smallAfter = (a) => {
      if(!a) return '';
      let out = '';
      // 送り仮名 → 訓点 → 左の振り仮名・左の送り仮名 の順に、小さい字で続ける
      [a.okuri, a.kunten, a.hFuri, a.hOkuri].forEach(s => { if(s) out += run(s, { size: 14 }); });
      return out;
    };
    const paras = [];
    let cur = '';
    for(let i = 0; i < text.length; ){
      const cp = text.codePointAt(i);
      const ch = String.fromCodePoint(cp);
      if(ch === '\n'){ paras.push(cur); cur = ''; i += ch.length; continue; }
      const sp = spanAt.get(i);
      if(sp){
        let base = '';
        for(let j = sp.start; j < sp.end && j < text.length; ){
          const c = String.fromCodePoint(text.codePointAt(j));
          base += run(plainChar(c), { u: underline(j) });
          j += c.length;
        }
        cur += ruby(base, sp.text || '');
        for(let j = sp.start; j < sp.end && j < text.length; ){
          const c = String.fromCodePoint(text.codePointAt(j));
          cur += smallAfter(annos.get(j));
          j += c.length;
        }
        i = Math.max(sp.end, i + ch.length);
        continue;
      }
      const a = annos.get(i);
      const base = run(plainChar(ch), { u: underline(i) });
      cur += (a && a.furi) ? ruby(base, a.furi) : base;
      cur += smallAfter(a);
      i += ch.length;
    }
    paras.push(cur);
    return paras.map(p => '<w:p><w:pPr><w:rPr>' + FONT + '</w:rPr></w:pPr>' + p + '</w:p>').join('');
  }

  // insertOoxml に渡せる、Flat OPC 形式のパッケージ
  function buildFlatOpc(doc){
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
      + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">'
      + '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      + '</Relationships></pkg:xmlData></pkg:part>'
      + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
      + '<pkg:xmlData><w:document xmlns:w="' + W + '"><w:body>' + buildParagraphsXml(doc) + '</w:body></w:document></pkg:xmlData></pkg:part>'
      + '</pkg:package>';
  }

  const api = { buildParagraphsXml, buildFlatOpc, plainChar };
  if(typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.KouriOoxml = api;
})(typeof window !== 'undefined' ? window : globalThis);

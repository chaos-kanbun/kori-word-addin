// Word の作業ウィンドウ側の処理。
// 流れ:範囲を内容コントロール(タグ kouri:<id>)で包む → ダイアログで交利を開く → 「反映」で OOXML として書き戻す
(function(){
  'use strict';
  const TAG_PREFIX = 'kouri:';
  const $ = (id) => document.getElementById(id);
  const status = (t, err) => { const el = $('status'); el.textContent = t; el.className = err ? 'err' : ''; };
  let dialog = null;

  Office.onReady((info) => {
    if(info.host !== Office.HostType.Word){ status('Word の中で開いてください。', true); return; }
    $('btn-selection').disabled = false; $('btn-body').disabled = false;
    status('文字を選んで「選択範囲を交利で開く」を押してください。');
    $('btn-selection').addEventListener('click', () => openEditor('selection'));
    $('btn-body').addEventListener('click', () => openEditor('body'));
  });

  const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const getRecord = (id) => Office.context.document.settings.get(TAG_PREFIX + id);
  function saveRecord(id, rec){
    return new Promise((resolve) => {
      Office.context.document.settings.set(TAG_PREFIX + id, rec);
      Office.context.document.settings.saveAsync(() => resolve());
    });
  }

  // 範囲を内容コントロールで包み(すでに交利の範囲ならそれを使い)、読み込む文書を返す
  async function prepare(kind){
    return Word.run(async (context) => {
      const range = kind === 'body' ? context.document.body.getRange('Whole') : context.document.getSelection();
      const parent = range.parentContentControlOrNullObject;
      parent.load('isNullObject,tag');
      await context.sync();
      let cc, id;
      if(!parent.isNullObject && parent.tag && parent.tag.indexOf(TAG_PREFIX) === 0){
        cc = parent; id = parent.tag.slice(TAG_PREFIX.length);
      } else {
        id = newId();
        cc = range.insertContentControl();
        cc.tag = TAG_PREFIX + id;
        cc.title = '交利';
        cc.appearance = 'BoundingBox';
      }
      cc.load('text');
      await context.sync();
      const wordText = cc.text;
      const rec = getRecord(id);
      // 前回 Word に反映した後、Word 側で本文が変わっていなければ、注釈ごと元の文書を使う
      const doc = (rec && rec.wordText === wordText && rec.doc)
        ? rec.doc
        : { version: 1, text: wordText.replace(/\r\n?/g, '\n').replace(/\n$/, '') };
      return { id, doc };
    });
  }

  async function applyToWord(id, doc){
    await Word.run(async (context) => {
      const cc = context.document.contentControls.getByTag(TAG_PREFIX + id).getFirst();
      cc.insertOoxml(KouriOoxml.buildFlatOpc(doc), 'Replace');
      await context.sync();
      cc.tag = TAG_PREFIX + id; // 念のため、書き戻したあとも目印を残す
      cc.load('text');
      await context.sync();
      await saveRecord(id, { doc, wordText: cc.text });
    });
  }

  // 交利のページを画像にしたものを、内容コントロールに並べる(見た目は交利のまま)。文字としての本文は、画像の代替テキストに入れる
  async function applyImagesToWord(id, doc, pages, width, sizes){
    await Word.run(async (context) => {
      const cc = context.document.contentControls.getByTag(TAG_PREFIX + id).getFirst();
      cc.clear();
      await context.sync();
      for(let i = 0; i < pages.length; i++){
        const w = width, h = width * sizes[i][1] / sizes[i][0];
        let pic;
        if(i === 0){
          pic = cc.insertInlinePictureFromBase64(pages[i], 'Start');
        } else {
          const para = cc.insertParagraph('', 'End');
          pic = para.insertInlinePictureFromBase64(pages[i], 'Start');
        }
        pic.lockAspectRatio = false;
        pic.width = w; pic.height = h;
        pic.altTextTitle = '交利 ' + (i + 1) + '/' + pages.length + ' ページ';
        pic.altTextDescription = doc.text.slice(0, 1000);
        await context.sync();
      }
      cc.tag = TAG_PREFIX + id;
      cc.load('text');
      await context.sync();
      await saveRecord(id, { doc, wordText: cc.text, image: true });
    });
  }

  async function openEditor(kind){
    if(dialog){ status('交利のウィンドウがすでに開いています。', true); return; }
    let prep;
    try{ prep = await prepare(kind); }
    catch(err){ status('Word の本文を読み込めませんでした:' + err.message, true); return; }
    const url = new URL('../index.html?addin=dialog', location.href).href;
    status('交利を開いています…');
    Office.context.ui.displayDialogAsync(url, { height: 92, width: 92, displayInIframe: false }, (res) => {
      if(res.status !== Office.AsyncResultStatus.Succeeded){
        status('ウィンドウを開けませんでした:' + res.error.message, true); return;
      }
      dialog = res.value;
      let img = null;  // 画像を受け取っている途中のもの
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
        let m; try{ m = JSON.parse(arg.message); } catch(err){ return; }
        if(m.type === 'img-start'){
          img = { doc: m.doc, width: m.width, sizes: m.sizes, parts: Array.from({ length: m.count }, () => []) };
          status('画像を受け取っています…');
          return;
        }
        if(m.type === 'img-chunk' && img){ img.parts[m.page][m.part] = m.data; return; }
        if(m.type === 'img-end' && img){
          try{
            status('Word に反映しています…');
            await applyImagesToWord(prep.id, img.doc, img.parts.map(p => p.join('')), img.width, img.sizes);
            status('Word に反映しました。');
          } catch(err){ status('Word への反映に失敗しました:' + err.message, true); }
          img = null; dialog.close(); dialog = null;
          return;
        }
        if(m.type === 'ready'){
          dialog.messageChild(JSON.stringify({ type: 'load', doc: prep.doc }));
          status('交利で編集中です。');
        } else if(m.type === 'apply'){
          try{ await applyToWord(prep.id, m.doc); status('Word に反映しました。'); }
          catch(err){ status('Word への反映に失敗しました:' + err.message, true); }
          dialog.close(); dialog = null;
        } else if(m.type === 'close'){
          dialog.close(); dialog = null; status('反映せずに閉じました。');
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => { dialog = null; status('交利のウィンドウが閉じられました。'); });
    });
  }
})();

// Word のアドインとして開かれたとき(?addin=dialog)だけ働く。
// Word の作業ウィンドウ(addin/taskpane.js)から文書を受け取り、「Wordに反映」で送り返す。
(function(){
  'use strict';
  if(!/[?&]addin=dialog(&|$)/.test(location.search)) return;
  const editor = window.KouriEditor;
  if(!editor) return;

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
  const apply = mk('Wordに反映して閉じる', true);
  const cancel = mk('反映せずに閉じる', false);
  bar.append(msg, apply, cancel);
  document.body.appendChild(bar);

  const post = (obj) => { try{ Office.context.ui.messageParent(JSON.stringify(obj)); } catch(err){ msg.textContent = 'Word と通信できません:' + err.message; } };

  apply.addEventListener('click', () => {
    apply.disabled = true; cancel.disabled = true;
    msg.textContent = 'Word に反映しています…';
    post({ type: 'apply', doc: editor.getDocument() });
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
        apply.disabled = false; cancel.disabled = false;
      });
      post({ type: 'ready' });
    });
  };
  document.head.appendChild(s);
})();

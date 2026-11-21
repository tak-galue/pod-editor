// 共通ユーティリティ
const U = (() => {
  let seq = 0;
  const uid = () => 'b' + Date.now().toString(36) + (seq++).toString(36) + Math.random().toString(36).slice(2, 5);
  const clone = (o) => JSON.parse(JSON.stringify(o));
  // XMLに入れられない制御文字を除去してエスケープ
  const esc = (s) => String(s ?? '')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const escHtml = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');

  let toastTimer;
  function toast(msg, ms = 2600) {
    const t = document.getElementById('toast');
    t.textContent = msg; t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), ms);
  }

  // ---- IndexedDB（自動保存）----
  const DB = 'pod-editor', STORE = 'kv';
  function db() {
    return new Promise((res, rej) => {
      const r = indexedDB.open(DB, 1);
      r.onupgradeneeded = () => r.result.createObjectStore(STORE);
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    });
  }
  async function idbSet(k, v) {
    try { const d = await db(); await new Promise((res, rej) => { const tx = d.transaction(STORE, 'readwrite'); tx.objectStore(STORE).put(v, k); tx.oncomplete = res; tx.onerror = () => rej(tx.error); }); return true; }
    catch (e) { console.warn('idbSet', e); return false; }
  }
  async function idbGet(k) {
    try { const d = await db(); return await new Promise((res, rej) => { const r = d.transaction(STORE).objectStore(STORE).get(k); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); }
    catch (e) { console.warn('idbGet', e); return undefined; }
  }
  // 複数キーの書込・削除と一覧（'index'）の更新を、1つのトランザクションでまとめて行う。
  // どれか1つでも失敗（容量不足など）すると全体が取り消され、false を返す。
  async function idbCommit({ put = {}, del = [], index = null } = {}) {
    try {
      const d = await db();
      await new Promise((res, rej) => {
        const tx = d.transaction(STORE, 'readwrite'), st = tx.objectStore(STORE);
        for (const [k, v] of Object.entries(put)) st.put(v, k);
        for (const k of del) st.delete(k);
        if (index) { const r = st.get('index'); r.onsuccess = () => st.put(index(Array.isArray(r.result) ? r.result : []), 'index'); }
        tx.oncomplete = res;
        tx.onerror = () => rej(tx.error);
        tx.onabort = () => rej(tx.error || new Error('保存が中断されました'));
      });
      return true;
    } catch (e) { console.warn('idbCommit', e); return false; }
  }
  async function idbDel(k) {
    try { const d = await db(); await new Promise((res, rej) => { const tx = d.transaction(STORE, 'readwrite'); tx.objectStore(STORE).delete(k); tx.oncomplete = res; tx.onerror = () => rej(tx.error); }); return true; }
    catch (e) { console.warn('idbDel', e); return false; }
  }

  // ---- 文字単位の差分（変更一覧用）----
  function diff(a, b) {
    a = a || ''; b = b || '';
    let p = 0; while (p < a.length && p < b.length && a[p] === b[p]) p++;
    let s = 0; while (s < a.length - p && s < b.length - p && a[a.length - 1 - s] === b[b.length - 1 - s]) s++;
    const A = a.slice(p, a.length - s), B = b.slice(p, b.length - s);
    const out = [];
    if (p) out.push(['=', a.slice(0, p)]);
    if (A.length * B.length > 250000) { if (A) out.push(['-', A]); if (B) out.push(['+', B]); }
    else {
      const n = A.length, m = B.length, dp = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
      for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--)
        dp[i][j] = A[i] === B[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
      let i = 0, j = 0;
      const push = (op, ch) => { const l = out[out.length - 1]; if (l && l[0] === op) l[1] += ch; else out.push([op, ch]); };
      while (i < n && j < m) {
        if (A[i] === B[j]) { push('=', A[i]); i++; j++; }
        else if (dp[i + 1][j] >= dp[i][j + 1]) push('-', A[i++]); else push('+', B[j++]);
      }
      while (i < n) push('-', A[i++]); while (j < m) push('+', B[j++]);
    }
    if (s) { const l = out[out.length - 1]; if (l && l[0] === '=') l[1] += a.slice(a.length - s); else out.push(['=', a.slice(a.length - s)]); }
    return out;
  }

  function readFile(file, as = 'arrayBuffer') {
    return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); as === 'text' ? r.readAsText(file) : as === 'dataURL' ? r.readAsDataURL(file) : r.readAsArrayBuffer(file); });
  }
  // claude.ai 上ではダウンロード機能（downloads）経由、ローカルでは通常のダウンロード
  async function download(blob, name) {
    if (window.claude && typeof window.claude.use === 'function') {
      const dl = await window.claude.use('downloads').catch(() => null);
      if (dl) {
        try { await dl.save({ filename: name, data: blob }); return 'saved'; }
        catch (e) {
          if (e && e.code === 'declined') return 'declined';
          if (e && e.code === 'rate_limited') return 'busy';
          throw new Error((e && e.message) || 'ファイルを保存できませんでした');
        }
      }
      return 'unavailable';
    }
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    return 'saved';
  }
  // ページ内の確認ダイアログ（埋め込み環境では confirm() が使えないため）
  function ask(title, msg, ok = '続ける') {
    const d = document.getElementById('dlgConfirm');
    document.getElementById('cfTitle').textContent = title;
    document.getElementById('cfMsg').textContent = msg;
    document.getElementById('cfOk').textContent = ok;
    d.returnValue = '';
    d.showModal();
    return new Promise((res) => d.addEventListener('close', () => res(d.returnValue === 'yes'), { once: true }));
  }
  function imgSize(dataUrl) {
    return new Promise((res) => { const im = new Image(); im.onload = () => res({ w: im.naturalWidth, h: im.naturalHeight }); im.onerror = () => res({ w: 800, h: 600 }); im.src = dataUrl; });
  }
  return { uid, clone, esc, escHtml, toast, idbSet, idbGet, idbDel, idbCommit, diff, readFile, download, ask, imgSize };
})();

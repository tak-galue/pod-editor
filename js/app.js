// POD原稿エディタ 本体
(() => {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));

  // ---------- 状態 ----------
  let S = newState();
  function newState() {
    const now = Date.now();
    return { version: 1, id: U.uid(), createdAt: now, updatedAt: now, fileName: '', blocks: [], images: {}, orig: {}, redline: [], settings: { title: '', author: '', opener: true, header: true, folio: true, size: 'A5' } };
  }
  let cur = null;            // フォーカス中ブロックID
  const sel = new Set();     // 選択ブロックID（ガター操作）
  let anchor = null;
  const hist = { undo: [], redo: [] };
  let lastTypeCp = 0;
  let pdfScale = 1;
  let view = 'upload'; // 'upload'（アップロード画面）| 'edit'

  const byId = (id) => S.blocks.find((b) => b.id === id);
  const idx = (id) => S.blocks.findIndex((b) => b.id === id);
  const elOf = (id) => $(`#paper .blk[data-id="${id}"]`);

  // ---------- 履歴 ----------
  const snap = () => JSON.stringify({ blocks: S.blocks, redline: S.redline });
  function checkpoint() {
    const s = snap();
    if (hist.undo[hist.undo.length - 1] !== s) { hist.undo.push(s); if (hist.undo.length > 80) hist.undo.shift(); }
    hist.redo = [];
  }
  function restore(s) { const o = JSON.parse(s); S.blocks = o.blocks; S.redline = o.redline; renderAll(); scheduleSave(); }
  function undo() { if (!hist.undo.length) return U.toast('これ以上戻せません'); const s = snap(); let p = hist.undo.pop(); if (p === s && hist.undo.length) p = hist.undo.pop(); hist.redo.push(s); restore(p); }
  function redo() { if (!hist.redo.length) return; hist.undo.push(snap()); restore(hist.redo.pop()); }

  // ---------- 自動保存 ----------
  // 保存は saveChain で1本に直列化する。flushSave／作業の切替／削除は、実行中の保存が終わるまで待つ。
  let saveTimer = null;
  let saveChain = Promise.resolve(true);
  const deletedIds = new Set();   // 削除済みの作業。遅れて走った保存で復活させない
  function scheduleSave() {
    $('#stSave').textContent = '未保存の変更…';
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => { saveTimer = null; saveProject().then(showSaved); }, 900);
    updateStatus();
  }
  function showSaved(ok) {
    $('#stSave').textContent = ok ? '自動保存済み ' + new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' }) : '自動保存に失敗（容量不足の可能性。「保存」で .json に書き出してください）';
    if (!ok) U.toast('自動保存に失敗しました。「保存」で .json に書き出してください', 6000);
  }
  // 保留中の自動保存をすぐ実行し、実行中の保存も含めて完了を待つ（画面切替・別の作業を開く前）
  async function flushSave() {
    if (saveTimer) { clearTimeout(saveTimer); saveTimer = null; saveProject().then(showSaved); }
    return saveChain;
  }
  const isDirty = () => !!saveTimer;

  // ---------- マイページ（このブラウザに保存した作業の一覧）----------
  // IndexedDB のキー：'index'＝一覧（メタ情報のみ）／'proj:<id>'＝作業データ／'red:<id>'＝赤入れ
  // 作業データと一覧は同じトランザクションで更新し、片方だけ保存された状態を作らない。
  // 旧版の単一保存（'project'／'red'）は初回起動時に一覧へ移行する（旧キーは消さずに残す）
  function metaOf(s) {
    return {
      id: s.id, title: s.settings.title || '', author: s.settings.author || '', fileName: s.fileName || '',
      createdAt: s.createdAt, updatedAt: s.updatedAt,
      paras: s.blocks.filter((b) => b.type !== 'pagebreak').length,
      checks: s.redline.length, done: s.redline.filter((r) => r.done).length,
    };
  }
  async function listProjects() { const l = await U.idbGet('index'); return Array.isArray(l) ? l : []; }
  // 保存はその時点の作業 s を対象にする（呼び出し後に作業が切り替わっても混ざらない）
  function saveProject(s = S) {
    const run = saveChain.then(async () => {
      if (!s.blocks.length || deletedIds.has(s.id)) return true;
      s.updatedAt = Date.now();
      const meta = metaOf(s);
      return U.idbCommit({ put: { ['proj:' + s.id]: s }, index: (l) => [meta, ...l.filter((m) => m.id !== s.id)] });
    });
    saveChain = run.catch(() => false);
    return run;
  }
  async function migrateLegacy() {
    if (Array.isArray(await U.idbGet('index'))) return;
    const old = await U.idbGet('project');
    const put = {}, list = [];
    if (old && Array.isArray(old.blocks) && old.blocks.length) {
      const s = Object.assign(newState(), old, { id: U.uid() });
      put['proj:' + s.id] = s; list.push(metaOf(s));
      const red = await U.idbGet('red'); if (Array.isArray(red) && red.length) put['red:' + s.id] = red;
    }
    put.index = list;
    // 失敗した場合は一覧を作らないので、次回の起動時にもう一度移行を試みる
    if (!(await U.idbCommit({ put }))) U.toast('以前の作業の移行に失敗しました（容量不足の可能性）。再読み込みで再試行します', 8000);
  }

  // ---------- 読み込み・切替の排他 ----------
  // 原稿や赤入れの読み込み中に作業を切り替えると、結果が別の作業に混ざるため、同時に1つだけ実行する
  let busy = false;
  async function exclusive(fn) {
    if (busy) { U.toast('読み込み中です。終わってから操作してください'); return false; }
    busy = true; document.body.classList.add('busy');
    try { return await fn(); }
    finally { busy = false; document.body.classList.remove('busy'); }
  }

  async function openProject(id) {
    await flushSave();
    const s = await U.idbGet('proj:' + id);
    if (!s || !Array.isArray(s.blocks)) {
      U.toast('作業データが見つかりませんでした');
      await U.idbCommit({ index: (l) => l.filter((m) => m.id !== id) });
      return renderMyPage();
    }
    const red = await U.idbGet('red:' + id);
    S = Object.assign(newState(), s); hist.undo = []; hist.redo = []; cur = null; sel.clear();
    redDocs = Array.isArray(red) ? red : []; pdfCache.clear();
    view = 'edit'; renderAll(); drawRed().catch(() => {});
    $('#stSave').textContent = '保存済み';
  }
  async function deleteProject(id) {
    const m = (await listProjects()).find((x) => x.id === id);
    const name = m ? (m.title || m.fileName || '無題') : '無題';
    if (!(await U.ask('この作業を削除しますか？', `「${name}」の作業データと赤入れをこのブラウザから削除します。元に戻せないので、残したい場合は先に開いて「保存」で .json に書き出してください。`, '削除する'))) return;
    // 実行中・保留中の保存を止めてから消す（削除後に保存が完了して復活するのを防ぐ）
    deletedIds.add(id);
    if (S.id === id) { clearTimeout(saveTimer); saveTimer = null; }
    await saveChain;
    const ok = await U.idbCommit({ del: ['proj:' + id, 'red:' + id], index: (l) => l.filter((x) => x.id !== id) });
    if (!ok) { deletedIds.delete(id); U.toast('削除できませんでした。もう一度お試しください', 6000); return renderMyPage(); }
    if (S.id === id) { S = newState(); hist.undo = []; hist.redo = []; redDocs = []; pdfCache.clear(); drawRed().catch(() => {}); }
    U.toast('削除しました'); renderMyPage();
  }
  const fmtDate = (t) => (t ? new Date(t).toLocaleString('ja-JP', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }) : '—');
  function el(tag, cls, text) { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
  let myPageSeq = 0;
  async function renderMyPage() {
    const box = $('#myProjects'); if (!box) return;
    const seq = ++myPageSeq;
    const list = (await listProjects()).slice().sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
    if (seq !== myPageSeq) return; // 新しい描画が始まっていれば古い結果は捨てる
    $('#myCount').textContent = list.length ? `${list.length}件` : '';
    if (!list.length) { box.replaceChildren(el('div', 'my-empty', 'まだ作業はありません。下から原稿をアップロードすると、ここに一覧で表示されます。')); return; }
    // 保存データ由来の文字列は textContent／dataset でのみ扱う（HTMLとして解釈させない）
    box.replaceChildren(...list.map((m) => {
      const item = el('div', 'my-item' + (m.id === S.id && S.blocks.length ? ' current' : ''));
      item.dataset.pid = String(m.id);
      const open = el('button', 'my-open'); open.dataset.open = ''; open.title = 'この作業を開いて編集する';
      const title = m.title || m.fileName || '無題';
      open.setAttribute('aria-label', `「${title}」を開く`);
      open.append(
        el('span', 'my-title', title),
        el('span', 'my-sub', [m.author, m.fileName].filter(Boolean).join('　／　')),
        el('span', 'my-meta', `更新 ${fmtDate(m.updatedAt)}　・　${m.paras || 0}段落${m.checks ? `　・　赤入れ ${m.done || 0}/${m.checks}件対応` : ''}`),
      );
      const del = el('button', 'my-del', '削除'); del.dataset.del = ''; del.title = 'この作業を削除'; del.setAttribute('aria-label', `「${title}」を削除`);
      item.append(open, del);
      return item;
    }));
  }

  // ---------- 描画 ----------
  function tagLabel(b) { return (TYPES[b.type] || {}).label || b.type; }
  function ed(field, text, ph = '', cls = '') {
    return `<div class="ed ${cls}" contenteditable="plaintext-only" spellcheck="false" data-field="${field}" data-ph="${U.escHtml(ph)}">${U.escHtml(text || '')}</div>`;
  }
  function changed(b) {
    if (!(b.id in S.orig)) return true;
    return S.orig[b.id] !== blockText(b);
  }
  function blockText(b) {
    if (b.type === 'chapter') return (b.num ? b.num + '　' : '') + (b.text || '');
    if (b.type === 'figure') return b.caption || '';
    if (b.type === 'table') return b.rows.map((r) => r.join('　')).join('\n');
    return b.text || '';
  }
  function blockHtml(b) {
    const redN = S.redline.filter((r) => r.blockId === b.id && !r.done).length;
    let inner = '';
    switch (b.type) {
      case 'chapter':
        inner = `${ed('num', b.num, '', 'chlabel')}${ed('text', b.text, '章タイトル', 'title')}<div class="rh">柱：${ed('runningHead', b.runningHead || '', (b.num ? b.num + '　' : '') + b.text)}</div>`; break;
      case 'figure': {
        const im = S.images[b.img];
        inner = `<figure>${ed('caption', b.caption, '図1-1　キャプション', 'cap')}${im ? `<img src="${im.data}" style="width:${b.width || 100}%" alt="">` : '<div class="muted">（画像なし）</div>'}${ed('note', b.note, '図の説明（任意）', 'note')}<div class="figtools">幅 <input type="range" min="20" max="100" step="5" value="${b.width || 100}" data-act="width"> <span>${b.width || 100}%</span><button data-act="replaceImg">画像差替</button></div></figure>`; break;
      }
      case 'table':
        inner = `<table>${b.rows.map((r, ri) => `<tr>${r.map((c, ci) => `<td>${ed(`cell:${ri}:${ci}`, c)}</td>`).join('')}</tr>`).join('')}</table>`; break;
      case 'toc':
        inner = `<div class="toc"><h3>目次</h3>${tocHtml()}<div class="muted small">※ページ番号はWordで開いた際に自動更新されます</div></div>`; break;
      case 'pagebreak':
        inner = '<div class="pb"><span>改ページ</span></div>'; break;
      default:
        inner = ed('text', b.text, '');
    }
    const cls = ['blk', 't-' + b.type, sel.has(b.id) ? 'sel' : '', cur === b.id ? 'cur' : '', changed(b) ? 'changed' : ''].join(' ');
    return `<div class="${cls}" data-id="${b.id}"${b.align ? ` data-align="${b.align}"` : ''}><div class="gut"><button class="tag" data-act="menu" tabindex="-1">${tagLabel(b)}</button>${redN ? `<span class="flags">赤入れ ${redN}</span>` : ''}</div>${inner}</div>`;
  }
  function tocHtml() {
    return S.blocks.filter((b) => b.type === 'chapter' || b.type === 'h2')
      .map((b) => `<div class="${b.type === 'chapter' ? 'l1' : 'l2'}">${U.escHtml(b.type === 'chapter' ? (b.num ? b.num + '　' : '') + b.text : b.text)}</div>`).join('');
  }
  function renderAll() {
    const up = view === 'upload' || !S.blocks.length;
    $('#empty').hidden = !up; $('#paper').hidden = up;
    document.body.classList.toggle('uploading', up);
    if (up) renderMyPage();
    $('#paper').innerHTML = S.blocks.map(blockHtml).join('');
    renderOutline(); renderChecks(); updateStatus(); syncTypeSelect();
  }
  function rerender(id) {
    const b = byId(id), el = elOf(id); if (!el) return;
    if (!b) { el.remove(); return; }
    const tmp = document.createElement('div'); tmp.innerHTML = blockHtml(b);
    el.replaceWith(tmp.firstElementChild);
  }
  function insertEl(afterId, b) {
    const tmp = document.createElement('div'); tmp.innerHTML = blockHtml(b);
    const ref = afterId ? elOf(afterId) : null;
    if (ref) ref.after(tmp.firstElementChild); else $('#paper').prepend(tmp.firstElementChild);
  }
  let outlineTimer;
  function renderOutlineSoon() { clearTimeout(outlineTimer); outlineTimer = setTimeout(() => { renderOutline(); $$('#paper .t-toc').forEach((e) => rerender(e.dataset.id)); }, 400); }
  function renderOutline() {
    $('#outlineList').innerHTML = S.blocks.filter((b) => b.type === 'chapter' || b.type === 'h2').map((b) =>
      `<a href="#" data-go="${b.id}" class="${b.type === 'chapter' ? 'ch' : 'h2'}">${U.escHtml(b.type === 'chapter' ? (b.num ? b.num + ' ' : '') + b.text : b.text)}</a>`).join('') || '<div class="muted small" style="padding:12px">見出しがありません</div>';
  }
  function updateStatus() {
    const chars = S.blocks.reduce((n, b) => n + blockText(b).length, 0);
    const ch = S.blocks.filter((b) => b.type === 'chapter').length;
    const fig = S.blocks.filter((b) => b.type === 'figure').length;
    const red = S.redline.filter((r) => !r.done).length;
    $('#stLeft').textContent = S.blocks.length ? `${S.fileName || '無題'}　｜　段落 ${S.blocks.length}　章 ${ch}　図版 ${fig}　約${chars.toLocaleString()}字　｜　赤入れ未対応 ${red}` : '';
  }
  function goTo(id, focus = true) {
    const el = elOf(id); if (!el) return;
    el.scrollIntoView({ block: 'center', behavior: 'smooth' });
    if (focus) { const e = el.querySelector('.ed'); if (e) { e.focus({ preventScroll: true }); setCaret(e, 0); } }
    el.animate([{ background: '#ffe9a8' }, { background: 'transparent' }], { duration: 1200 });
  }

  // ---------- キャレット ----------
  function caretOffset(el) {
    const s = getSelection(); if (!s.rangeCount || !el.contains(s.focusNode)) return -1;
    const r = document.createRange(); r.selectNodeContents(el); r.setEnd(s.focusNode, s.focusOffset);
    return r.toString().length;
  }
  function setCaret(el, off) {
    const s = getSelection(), r = document.createRange();
    let n = 0, done = false;
    const walk = (node) => {
      if (done) return;
      if (node.nodeType === 3) { if (n + node.length >= off) { r.setStart(node, off - n); done = true; } else n += node.length; }
      else for (const c of node.childNodes) walk(c);
    };
    walk(el);
    if (!done) { r.selectNodeContents(el); r.collapse(false); } else r.collapse(true);
    s.removeAllRanges(); s.addRange(r);
  }
  const edText = (el) => el.innerText.replace(/\n$/, '');

  // ---------- 編集イベント ----------
  const paper = $('#paper');
  // 入力の区切り（1.5秒以上の間隔）ごとに、モデル更新前の状態を履歴に積む
  const typingCheckpoint = () => { if (Date.now() - lastTypeCp > 1500) checkpoint(); lastTypeCp = Date.now(); };
  paper.addEventListener('input', (e) => {
    typingCheckpoint();
    const t = e.target;
    if (t.matches('input[data-act=width]')) {
      const b = byId(t.closest('.blk').dataset.id); b.width = +t.value;
      t.nextElementSibling.textContent = t.value + '%'; t.closest('figure').querySelector('img').style.width = t.value + '%';
      scheduleSave(); return;
    }
    const edEl = t.closest('.ed'); if (!edEl) return;
    const blk = edEl.closest('.blk'); const b = byId(blk.dataset.id); if (!b) return;
    const f = edEl.dataset.field, v = edText(edEl);
    if (f.startsWith('cell:')) { const [, r, c] = f.split(':'); b.rows[+r][+c] = v; }
    else b[f] = v;
    blk.classList.toggle('changed', changed(b));
    if (b.type === 'chapter' || b.type === 'h2') renderOutlineSoon();
    scheduleSave();
  });
  paper.addEventListener('change', (e) => { if (e.target.matches('input[data-act=width]')) checkpoint(); });
  paper.addEventListener('focusin', (e) => {
    const blk = e.target.closest('.blk'); if (!blk) return;
    if (cur && cur !== blk.dataset.id) elOf(cur)?.classList.remove('cur');
    cur = blk.dataset.id; blk.classList.add('cur'); syncTypeSelect(); markOutline();
  });
  paper.addEventListener('paste', (e) => {
    const edEl = e.target.closest('.ed'); if (!edEl) return;
    e.preventDefault();
    const text = (e.clipboardData.getData('text/plain') || '').replace(/\r\n?/g, '\n');
    document.execCommand('insertText', false, text);
  });
  paper.addEventListener('click', (e) => {
    const act = e.target.closest('[data-act]');
    const blk = e.target.closest('.blk'); if (!blk) return;
    const id = blk.dataset.id;
    if (act && act.dataset.act === 'menu') {
      e.stopPropagation();
      if (e.shiftKey && anchor) { selectRange(anchor, id); return; }
      if (e.metaKey || e.ctrlKey) { sel.has(id) ? sel.delete(id) : sel.add(id); anchor = id; blk.classList.toggle('sel'); return; }
      openMenu(id, act);
      return;
    }
    if (act && act.dataset.act === 'replaceImg') { pickImage((imgId) => { checkpoint(); byId(id).img = imgId; rerender(id); scheduleSave(); }); return; }
    if (!e.shiftKey && sel.size && !e.target.closest('.gut')) clearSel();
    if (blk.classList.contains('t-pagebreak') || blk.classList.contains('t-toc')) { cur = id; $$('#paper .blk.cur').forEach((x) => x.classList.remove('cur')); blk.classList.add('cur'); }
  });
  function clearSel() { sel.clear(); $$('#paper .blk.sel').forEach((x) => x.classList.remove('sel')); }
  function selectRange(a, b) {
    const i = idx(a), j = idx(b); clearSel();
    for (let k = Math.min(i, j); k <= Math.max(i, j); k++) { sel.add(S.blocks[k].id); elOf(S.blocks[k].id)?.classList.add('sel'); }
  }

  paper.addEventListener('keydown', (e) => {
    const edEl = e.target.closest('.ed'); if (!edEl) return;
    const blk = edEl.closest('.blk'); const id = blk.dataset.id; const b = byId(id);
    const f = edEl.dataset.field;
    const off = caretOffset(edEl), len = edText(edEl).length;
    const collapsed = getSelection().isCollapsed;
    if (e.isComposing || e.keyCode === 229) return;
    if (e.key === 'Enter' && !e.shiftKey && f === 'text' && isText(b)) {
      e.preventDefault(); splitAt(id, off); return;
    }
    if (e.key === 'Enter' && !e.shiftKey && f !== 'text' && !f.startsWith('cell')) { e.preventDefault(); focusNeighbor(id, 1, 0); return; }
    if (e.key === 'Backspace' && collapsed && off === 0 && f === 'text' && isText(b)) {
      const i = idx(id), prev = S.blocks[i - 1];
      if (prev && isText(prev)) { e.preventDefault(); mergeInto(prev.id, id); }
      else if (prev && prev.type === 'pagebreak') { e.preventDefault(); checkpoint(); S.blocks.splice(i - 1, 1); elOf(prev.id).remove(); scheduleSave(); }
      return;
    }
    if (e.key === 'Delete' && collapsed && off === len && f === 'text' && isText(b)) {
      const next = S.blocks[idx(id) + 1];
      if (next && isText(next)) { e.preventDefault(); mergeInto(id, next.id); }
      return;
    }
    if ((e.key === 'ArrowUp' || e.key === 'ArrowLeft') && collapsed && off === 0) { if (focusNeighbor(id, -1, 'end')) e.preventDefault(); return; }
    if ((e.key === 'ArrowDown' || e.key === 'ArrowRight') && collapsed && off === len) { if (focusNeighbor(id, 1, 0)) e.preventDefault(); return; }
  });
  function focusNeighbor(id, dir, pos) {
    let i = idx(id) + dir;
    while (i >= 0 && i < S.blocks.length) {
      const el = elOf(S.blocks[i].id)?.querySelector(dir > 0 ? '.ed' : '.ed:last-of-type') || null;
      const eds = elOf(S.blocks[i].id) ? $$('.ed', elOf(S.blocks[i].id)) : [];
      const target = dir > 0 ? eds[0] : eds[eds.length - 1];
      if (target || el) { const t = target || el; t.focus(); setCaret(t, pos === 'end' ? edText(t).length : 0); return true; }
      i += dir;
    }
    return false;
  }
  function splitAt(id, off) {
    checkpoint();
    const b = byId(id), t = b.text || '';
    const before = t.slice(0, off), after = t.slice(off);
    b.text = before;
    const nb = { id: U.uid(), type: ['h2', 'h3', 'lead', 'chapterSub'].includes(b.type) && after === '' ? 'body' : b.type, text: after };
    if (b.align) nb.align = b.align;
    S.blocks.splice(idx(id) + 1, 0, nb);
    rerender(id); insertEl(id, nb);
    const e = elOf(nb.id).querySelector('.ed'); e.focus(); setCaret(e, 0);
    scheduleSave(); renderOutlineSoon();
  }
  function mergeInto(prevId, id) {
    checkpoint();
    const p = byId(prevId), b = byId(id); const pos = (p.text || '').length;
    p.text = (p.text || '') + (b.text || '');
    S.redline.forEach((r) => { if (r.blockId === id) r.blockId = prevId; });
    S.blocks.splice(idx(id), 1);
    elOf(id).remove(); rerender(prevId);
    const e = elOf(prevId).querySelector('.ed[data-field=text]'); e.focus(); setCaret(e, pos);
    scheduleSave(); renderOutlineSoon();
  }

  // ---------- 種類変更・操作 ----------
  function targets() { return sel.size ? S.blocks.filter((b) => sel.has(b.id)).map((b) => b.id) : cur ? [cur] : []; }
  function convert(b, type) {
    if (b.type === type) return b;
    if (!(isText(b) || b.type === 'chapter')) return b;
    const text = b.type === 'chapter' ? (b.num ? b.num + '　' : '') + b.text : b.text || '';
    const base = { id: b.id, type };
    if (b.align && type === 'body') base.align = b.align;
    if (type === 'chapter') {
      const m = Rules.RE_CH.exec(text.trim());
      return m ? { ...base, num: `第${m[1]}章`, text: m[2].trim() } : { ...base, num: '', text: text.trim() };
    }
    if (type === 'h3' && !/^[◇◆]/.test(text)) return { ...base, text: (text.startsWith('「') ? '◇ ' : '◇　') + text };
    if (type === 'lead' && !/^─/.test(text)) return { ...base, text: `─ ${text.replace(/^[─―—]+\s*/, '')} ─` };
    if (b.type === 'h3' && type !== 'h3') return { ...base, text: text.replace(/^[◇◆][\s　]*/, '') };
    if (b.type === 'lead' && type !== 'lead') return { ...base, text: text.replace(/^─\s*/, '').replace(/\s*─$/, '') };
    return { ...base, text };
  }
  function setType(type) {
    const ids = targets(); if (!ids.length) return U.toast('段落を選択してください');
    checkpoint();
    ids.forEach((id) => { const i = idx(id); S.blocks[i] = convert(S.blocks[i], type); rerender(id); });
    renderOutline(); scheduleSave(); syncTypeSelect();
    const e = cur && elOf(cur)?.querySelector('.ed'); if (e && ids.length === 1) e.focus();
  }
  function setAlign(a) {
    const ids = targets(); if (!ids.length) return;
    checkpoint();
    ids.forEach((id) => { const b = byId(id); if (b.type !== 'body') return; if (a) b.align = a; else delete b.align; rerender(id); });
    scheduleSave();
  }
  function insertAfter(id, nb) {
    checkpoint();
    const i = id ? idx(id) : S.blocks.length - 1;
    S.blocks.splice(i + 1, 0, nb);
    if (S.blocks.length === 1) renderAll(); else insertEl(S.blocks[i]?.id, nb);
    scheduleSave(); renderOutlineSoon();
    return nb;
  }
  function deleteBlocks(ids) {
    if (!ids.length) return;
    checkpoint();
    ids.forEach((id) => { S.blocks.splice(idx(id), 1); elOf(id)?.remove(); sel.delete(id); });
    scheduleSave(); renderOutline();
  }
  function toTable() {
    const ids = targets().filter((id) => isText(byId(id)));
    if (ids.length < 1) return U.toast('表にしたい段落をガターのラベルで選択してください（Shift+クリックで範囲選択）');
    checkpoint();
    const rows = [];
    ids.forEach((id) => {
      const t = byId(id).text.trim().replace(/\n+/g, ''); // 手動改行は表のセル内で自動折返しに任せる
      const m = /^(\S+?)(?:[ \t ]{2,}|\t|　+| (?=\S))([\s\S]*)$/.exec(t);
      if (m && rows.length && !m[2]) rows.push([m[1], '']);
      else if (m) rows.push([m[1], m[2].replace(/^[\s　]+/, '')]);
      else if (rows.length) rows[rows.length - 1][1] += '　' + t; // 折り返し行は直前の行に連結
      else rows.push([t, '']);
    });
    const first = idx(ids[0]);
    const tb = { id: U.uid(), type: 'table', rows };
    S.blocks = S.blocks.filter((b) => !ids.includes(b.id));
    S.blocks.splice(first, 0, tb);
    clearSel(); renderAll(); scheduleSave(); goTo(tb.id, false);
    U.toast(`${ids.length}段落を${rows.length}行の表に変換しました`);
  }
  function syncTypeSelect() {
    const b = cur && byId(cur);
    $('#selType').value = b && TYPES[b.type] ? b.type : '';
  }
  function markOutline() {
    let last = null;
    for (const b of S.blocks) { if (b.type === 'chapter' || b.type === 'h2') last = b.id; if (b.id === cur) break; }
    $$('#outlineList a').forEach((a) => a.classList.toggle('on', a.dataset.go === last));
  }

  // ---------- ブロックメニュー ----------
  function openMenu(id, anchorEl) {
    cur = id; const b = byId(id);
    if (!sel.has(id)) { clearSel(); sel.add(id); elOf(id).classList.add('sel'); }
    anchor = id;
    const m = $('#menu');
    const typeBtns = isText(b) || b.type === 'chapter'
      ? '<div class="sec">種類</div>' + Object.entries(TYPES).filter(([k]) => isText({ type: k }) || k === 'chapter')
        .map(([k, v]) => `<button data-type="${k}" class="${b.type === k ? 'cur' : ''}">${v.label}<kbd>Alt+${v.key}</kbd></button>`).join('') : '';
    m.innerHTML = typeBtns + `<div class="sec">挿入（下に）</div>
      <button data-do="addBody">本文段落</button><button data-do="addBreak">改ページ</button><button data-do="addFig">図版（画像）</button><button data-do="addToc">目次</button>
      <div class="sec">操作</div>
      ${sel.size > 1 || isText(b) ? '<button data-do="toTable">選択を表に変換</button>' : ''}
      <button data-do="redline">赤入れ項目に紐付け…</button>
      <button data-do="del" style="color:#c62828">削除${sel.size > 1 ? `（${sel.size}件）` : ''}</button>
      <div class="sec">Shift+クリックで範囲選択 ／ ⌘+クリックで追加選択</div>`;
    const r = anchorEl.getBoundingClientRect();
    m.hidden = false;
    m.style.left = Math.min(r.right + 6, innerWidth - 220) + 'px';
    m.style.top = Math.min(r.top, innerHeight - m.offsetHeight - 10) + 'px';
  }
  document.addEventListener('click', (e) => { if (!e.target.closest('#menu')) $('#menu').hidden = true; });
  $('#menu').addEventListener('click', (e) => {
    const t = e.target.closest('button'); if (!t) return;
    $('#menu').hidden = true;
    if (t.dataset.type) return setType(t.dataset.type);
    const id = cur;
    switch (t.dataset.do) {
      case 'addBody': { const nb = insertAfter(id, { id: U.uid(), type: 'body', text: '' }); const e2 = elOf(nb.id).querySelector('.ed'); e2.focus(); break; }
      case 'addBreak': insertAfter(id, { id: U.uid(), type: 'pagebreak' }); break;
      case 'addToc': insertAfter(id, { id: U.uid(), type: 'toc' }); break;
      case 'addFig': pickImage((imgId) => insertAfter(id, { id: U.uid(), type: 'figure', img: imgId, caption: '', note: '', width: 100 })); break;
      case 'toTable': toTable(); break;
      case 'del': deleteBlocks(targets()); break;
      case 'redline': switchTab('check'); $('#chkNote').focus(); break;
    }
  });
  function pickImage(cb) {
    const inp = $('#fileImg');
    inp.onchange = async () => {
      const f = inp.files[0]; inp.value = ''; if (!f) return;
      const data = await U.readFile(f, 'dataURL'); const sz = await U.imgSize(data);
      const imgId = 'img_' + U.uid();
      S.images[imgId] = { name: f.name, mime: f.type || 'image/png', data, w: sz.w, h: sz.h };
      cb(imgId); scheduleSave();
    };
    inp.click();
  }

  // ---------- 読み込み ----------
  async function openManuscript(file) {
    if (S.blocks.length && view === 'edit' && !(await U.ask('原稿を読み込み直しますか？', '今の編集内容は破棄されます。残したい場合は先に「保存」で作業データを書き出してください。', '読み込む'))) return false;
    const isText = /\.(txt|text)$/i.test(file.name) || file.type === 'text/plain';
    if (!isText && !/\.docx$/i.test(file.name)) { U.toast('原稿は Word（.docx）またはテキスト（.txt）を選んでください'); return false; }
    // 編集中の読み直しは同じ作業を置き換え、アップロード画面からは新しい作業として追加する。
    // 対象は読み込み開始時点で確定させる（解析中に作業が変わっても別の作業を上書きしない）
    const prev = S;
    const reuse = view === 'edit' && prev.blocks.length > 0;
    U.toast('読み込み中…', 8000);
    try {
      const buf = await U.readFile(file);
      const { blocks, images } = isText ? DocxImport.parseText(buf) : await DocxImport.parse(buf);
      if (S !== prev) { U.toast('読み込み中に作業が切り替わったため、読み込みを中止しました', 6000); return false; }
      await flushSave();
      if (S !== prev) { U.toast('読み込み中に作業が切り替わったため、読み込みを中止しました', 6000); return false; }
      S = newState(); S.fileName = file.name; S.blocks = blocks; S.images = images;
      if (reuse) { S.id = prev.id; S.createdAt = prev.createdAt; S.settings = prev.settings; }
      else { redDocs = []; pdfCache.clear(); drawRed().catch(() => {}); }
      const firstText = blocks.find((b) => b.type === 'body' && b.text.trim());
      S.settings.title = firstText ? firstText.text.trim().slice(0, 60) : '';
      snapshotOrig(); hist.undo = []; hist.redo = []; cur = null; sel.clear();
      view = 'edit'; renderAll(); scheduleSave();
      U.toast(`${blocks.filter((b) => b.type !== 'pagebreak').length}段落${Object.keys(images).length ? '・画像' + Object.keys(images).length + '点' : ''}を読み込みました`);
      openRules();
      return true;
    } catch (err) { console.error(err); U.toast('読み込めませんでした：' + err.message, 6000); return false; }
  }
  function snapshotOrig() { S.orig = {}; S.blocks.forEach((b) => (S.orig[b.id] = blockText(b))); }

  // ---------- 自動整形 ----------
  const ruleOn = {};
  Rules.list.forEach((r) => (ruleOn[r.id] = r.on));
  function openRules() {
    if (!S.blocks.length) return U.toast('先に原稿を読み込んでください');
    const draw = () => {
      const { counts } = Rules.preview(U.clone(S.blocks), ruleOn);
      $('#rulesList').innerHTML = Rules.list.map((r) => `<label class="rule"><input type="checkbox" data-rule="${r.id}" ${ruleOn[r.id] ? 'checked' : ''}><div><div>${r.label}</div><div class="rd">${r.desc}</div></div><span class="cnt">${ruleOn[r.id] ? counts[r.id] : '–'}</span></label>`).join('');
    };
    draw();
    $('#rulesList').onchange = (e) => { const c = e.target.closest('[data-rule]'); if (c) { ruleOn[c.dataset.rule] = c.checked; draw(); } };
    $('#dlgRules').showModal();
  }
  $('#rulesApply').addEventListener('click', () => {
    checkpoint();
    const { blocks } = Rules.preview(U.clone(S.blocks), ruleOn);
    S.blocks = blocks;
    // 削除された段落に紐付いた赤入れは近い段落へ
    const ids = new Set(blocks.map((b) => b.id));
    S.redline.forEach((r) => { if (r.blockId && !ids.has(r.blockId)) r.blockId = null; });
    renderAll(); scheduleSave();
    U.toast('自動整形を適用しました（元に戻す：⌘Z）');
  });

  // ---------- 赤入れ（PDF／スキャン画像、複数可）----------
  // redDocs: [{name, kind:'pdf'|'img', bytes?, data?}]　頁番号は全ファイル通し
  let redDocs = [];
  const pdfCache = new Map();
  async function pdfOf(doc) {
    if (!pdfCache.has(doc)) {
      // 別オリジンのWorkerは直接起動できないため、blob: のWorkerからCDNのworkerを読み込む
      if (!pdfjsLib.GlobalWorkerOptions.workerSrc) {
        const src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(new Blob([`importScripts(${JSON.stringify(src)});`], { type: 'text/javascript' }));
      }
      pdfCache.set(doc, pdfjsLib.getDocument({ data: doc.bytes.slice(0) }).promise);
    }
    return pdfCache.get(doc);
  }
  async function addRedFiles(files, { replace = false } = {}) {
    const target = S; // 読み込み開始時の作業。読み込み中に切り替わったら追加しない
    const add = [];
    for (const f of files) {
      if (/\.pdf$/i.test(f.name) || f.type === 'application/pdf') {
        if (!window.pdfjsLib) { U.toast('PDFビューアを読み込めませんでした。ネット接続を確認してください'); continue; }
        add.push({ name: f.name, kind: 'pdf', bytes: new Uint8Array(await U.readFile(f)) });
      } else if (/^image\//.test(f.type) || /\.(png|jpe?g)$/i.test(f.name)) {
        add.push({ name: f.name, kind: 'img', data: await U.readFile(f, 'dataURL') });
      } else U.toast(`${f.name} は赤入れとして読み込めません（PDF／JPG／PNG）`);
    }
    if (!add.length) return;
    if (S !== target) { U.toast('読み込み中に作業が切り替わったため、赤入れを追加しませんでした', 6000); return; }
    redDocs = replace ? add : redDocs.concat(add);
    const ok = await U.idbCommit({ put: { ['red:' + target.id]: redDocs } });
    if (!ok) U.toast('赤入れを保存できませんでした（容量不足の可能性）。表示はされていますが、再読み込みすると消えます', 8000);
    if (S !== target) return;
    await drawRed(); switchTab('pdf');
  }
  async function drawRed() {
    const box = $('#pdfPages'); box.innerHTML = '';
    if (!redDocs.length) { $('#pdfName').textContent = '赤入れ未読込'; return; }
    const width = box.clientWidth || 340;
    let pageNo = 0;
    for (const doc of redDocs) {
      if (doc.kind === 'img') {
        pageNo++;
        const lab = document.createElement('div'); lab.className = 'pg'; lab.textContent = `${pageNo}　${doc.name}`;
        const im = document.createElement('img'); im.src = doc.data; im.dataset.page = pageNo; im.style.width = width * pdfScale + 'px';
        box.append(lab, im); continue;
      }
      try {
        const pdf = await pdfOf(doc);
        for (let p = 1; p <= pdf.numPages; p++) {
          pageNo++;
          const page = await pdf.getPage(p);
          const vp0 = page.getViewport({ scale: 1 });
          const vp = page.getViewport({ scale: (width * pdfScale * devicePixelRatio) / vp0.width });
          const c = document.createElement('canvas'); c.width = vp.width; c.height = vp.height; c.style.width = width * pdfScale + 'px';
          c.dataset.page = pageNo; c.dataset.doc = redDocs.indexOf(doc); c.dataset.p = p;
          const lab = document.createElement('div'); lab.className = 'pg'; lab.textContent = `${pageNo}　${doc.name}（${p}/${pdf.numPages}）`;
          box.append(lab, c);
          await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
        }
      } catch (err) { console.error(err); const e = document.createElement('div'); e.className = 'pg'; e.textContent = `${doc.name} を表示できませんでした`; box.append(e); }
    }
    $('#pdfName').textContent = `${redDocs.length}ファイル・${pageNo}頁`;
  }
  $('#pdfPages').addEventListener('click', async (e) => {
    const t = e.target.closest('canvas,img'); if (!t) return;
    $('#chkPage').value = t.dataset.page;
    const lc = $('#lbCanvas'), ctx = lc.getContext('2d');
    if (t.tagName === 'IMG') {
      const k = Math.min(1, (innerWidth - 60) / t.naturalWidth);
      lc.width = t.naturalWidth; lc.height = t.naturalHeight; lc.style.width = t.naturalWidth * k + 'px';
      ctx.drawImage(t, 0, 0);
    } else {
      const page = await (await pdfOf(redDocs[+t.dataset.doc])).getPage(+t.dataset.p);
      const vp0 = page.getViewport({ scale: 1 });
      const vp = page.getViewport({ scale: Math.min(2.2, (innerWidth - 60) / vp0.width) * devicePixelRatio });
      lc.width = vp.width; lc.height = vp.height; lc.style.width = vp.width / devicePixelRatio + 'px';
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
    }
    $('#lightbox').hidden = false;
  });
  $('#lightbox').addEventListener('click', () => ($('#lightbox').hidden = true));
  $('#pdfZoomIn').onclick = () => { pdfScale = Math.min(3, pdfScale + 0.25); drawRed(); };
  $('#pdfZoomOut').onclick = () => { pdfScale = Math.max(0.5, pdfScale - 0.25); drawRed(); };

  // ---------- 赤入れチェックリスト ----------
  function renderChecks() {
    const list = $('#chkList');
    if (!S.redline.length) { list.innerHTML = '<div class="muted small" style="padding:12px 0">赤入れ項目はまだありません。PDFの頁をクリックすると頁番号が入ります。</div>'; return; }
    list.innerHTML = S.redline.map((r) => {
      const b = r.blockId && byId(r.blockId);
      return `<div class="chk ${r.done ? 'done' : ''}" data-rid="${r.id}"><input type="checkbox" ${r.done ? 'checked' : ''} data-done><div><span class="pg">p.${r.page || '?'}</span> <span class="txt">${U.escHtml(r.note)}</span><span class="ctx">${b ? '→ ' + U.escHtml(blockText(b).slice(0, 34)) : '（段落未紐付け）'}</span></div><button class="del" data-del title="削除">×</button></div>`;
    }).join('');
    updateStatus();
  }
  $('#chkAdd').onclick = () => {
    const note = $('#chkNote').value.trim(); if (!note) return $('#chkNote').focus();
    checkpoint();
    S.redline.push({ id: U.uid(), page: +$('#chkPage').value || null, note, blockId: cur, done: false });
    $('#chkNote').value = ''; renderChecks(); if (cur) rerender(cur); scheduleSave();
  };
  $('#chkNote').addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.isComposing) $('#chkAdd').click(); });
  $('#chkList').addEventListener('click', (e) => {
    const row = e.target.closest('.chk'); if (!row) return;
    const r = S.redline.find((x) => x.id === row.dataset.rid);
    if (e.target.matches('[data-done]')) { checkpoint(); r.done = e.target.checked; renderChecks(); if (r.blockId) rerender(r.blockId); scheduleSave(); return; }
    if (e.target.matches('[data-del]')) { checkpoint(); S.redline = S.redline.filter((x) => x !== r); renderChecks(); if (r.blockId) rerender(r.blockId); scheduleSave(); return; }
    if (r.blockId) goTo(r.blockId);
  });

  // ---------- 変更一覧 ----------
  function renderChanges() {
    const out = [];
    const ids = new Set(S.blocks.map((b) => b.id));
    for (const b of S.blocks) {
      if (!changed(b)) continue;
      const before = S.orig[b.id];
      const after = blockText(b);
      const html = before === undefined ? `<ins>${U.escHtml(after.slice(0, 200))}</ins>` :
        U.diff(before, after).map(([op, s]) => {
          const t = U.escHtml(op === '=' && s.length > 40 ? s.slice(0, 18) + '…' + s.slice(-18) : s);
          return op === '=' ? t : op === '-' ? `<del>${t}</del>` : `<ins>${t}</ins>`;
        }).join('');
      out.push(`<div class="chg" data-go="${b.id}"><div class="k">${tagLabel(b)}${before === undefined ? '・追加' : ''}</div>${html || '<span class="muted">（書式のみ変更）</span>'}</div>`);
    }
    const removed = Object.keys(S.orig).filter((id) => !ids.has(id) && S.orig[id].trim()).length;
    $('#chgList').innerHTML = `<div class="muted small" style="margin-bottom:6px">変更 ${out.length} 段落／削除 ${removed} 段落（空段落を除く）</div>` + out.slice(0, 600).join('');
  }
  $('#chgRefresh').onclick = renderChanges;
  $('#chgList').addEventListener('click', (e) => { const c = e.target.closest('[data-go]'); if (c) goTo(c.dataset.go); });

  // ---------- 検索・置換 ----------
  let findPos = -1;
  function matcher() {
    const f = $('#fFind').value; if (!f) return null;
    try { return $('#fRegex').checked ? new RegExp(f, 'g') : new RegExp(f.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'); }
    catch { $('#fResult').textContent = '正規表現が正しくありません'; return null; }
  }
  const fields = (b) => b.type === 'chapter' ? ['num', 'text', 'runningHead'] : b.type === 'figure' ? ['caption', 'note'] : b.type === 'table' ? [] : ['text'];
  function findNext() {
    const re = matcher(); if (!re) return;
    const n = S.blocks.length;
    for (let k = 1; k <= n; k++) {
      const i = (findPos + k) % n, b = S.blocks[i];
      if ([...fields(b).map((f) => b[f] || ''), ...(b.rows || []).flat()].some((t) => { re.lastIndex = 0; return re.test(t); })) {
        findPos = i; $$('#paper .hit').forEach((x) => x.classList.remove('hit'));
        const el = elOf(b.id); el.classList.add('hit'); el.scrollIntoView({ block: 'center' });
        cur = b.id;
        $('#fResult').textContent = `${i + 1}段落目で見つかりました`; return;
      }
    }
    $('#fResult').textContent = '見つかりません';
  }
  function replaceAll() {
    const re = matcher(); if (!re) return;
    const rep = $('#fRepl').value;
    checkpoint();
    let n = 0;
    for (const b of S.blocks) {
      let hit = false;
      for (const f of fields(b)) { const v = b[f]; if (typeof v !== 'string') continue; const nv = v.replace(re, () => { n++; return '\u0000'; }); if (nv !== v) { b[f] = v.replace(re, rep); hit = true; } }
      if (b.rows) b.rows = b.rows.map((r) => r.map((c) => { const nv = c.replace(re, rep); if (nv !== c) { hit = true; n += (c.match(re) || []).length; } return nv; }));
      if (hit) rerender(b.id);
    }
    $('#fResult').textContent = `${n}件置換しました`; renderOutline(); scheduleSave();
  }
  $('#fNext').onclick = findNext;
  $('#fAll').onclick = replaceAll;
  $('#fFind').addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.isComposing) findNext(); });
  $$('.presets button').forEach((b) => (b.onclick = () => { $('#fFind').value = b.dataset.f; $('#fRepl').value = b.dataset.r; $('#fRegex').checked = !!b.dataset.re; findPos = -1; findNext(); }));

  // ---------- タブ ----------
  function switchTab(name) {
    $$('.tabs button').forEach((b) => b.classList.toggle('on', b.dataset.tab === name));
    $$('.tabpane').forEach((p) => p.classList.toggle('on', p.id === 'tab-' + name));
    if (name === 'changes') renderChanges();
  }
  $$('.tabs button').forEach((b) => (b.onclick = () => switchTab(b.dataset.tab)));
  $('#outlineList').addEventListener('click', (e) => { const a = e.target.closest('[data-go]'); if (a) { e.preventDefault(); goTo(a.dataset.go); } });

  // ---------- ツールバー ----------
  const st = $('#selType');
  st.innerHTML = '<option value="">種類…</option>' + Object.entries(TYPES).filter(([k]) => isText({ type: k }) || k === 'chapter').map(([k, v]) => `<option value="${k}">${v.label}</option>`).join('');
  st.onchange = () => { if (st.value) setType(st.value); };
  // ---------- アップロード画面 ----------
  let pendManu = null, pendRed = [];
  async function showUploader() { await flushSave(); view = 'upload'; pendManu = null; pendRed = []; paintPicked(); renderAll(); }
  function paintPicked() {
    $('#pickedManu').textContent = pendManu ? pendManu.name : '';
    $('#pickedRed').textContent = pendRed.map((f) => f.name).join('、');
    $('#dropManu').classList.toggle('done', !!pendManu); $('#dropRed').classList.toggle('done', pendRed.length > 0);
    $('#btnStart').disabled = !pendManu;
  }
  function takeFiles(files) {
    for (const f of files) {
      if (/\.(docx|txt|text)$/i.test(f.name) || f.type === 'text/plain') pendManu = f;
      else if (/\.(pdf|png|jpe?g)$/i.test(f.name)) pendRed.push(f);
      else if (/\.json$/i.test(f.name)) exclusive(() => loadProject(f));
      else U.toast(`${f.name} は読み込めない形式です`);
    }
    paintPicked();
  }
  $('#btnOpenDocx').onclick = () => exclusive(showUploader);
  $('#fileDocx').onchange = (e) => { const f = e.target.files[0]; e.target.value = ''; if (!f) return; if (view === 'upload') { pendManu = f; paintPicked(); } else exclusive(() => openManuscript(f)); };
  $('#btnOpenPdf').onclick = () => $('#filePdf').click();
  $('#filePdf').onchange = (e) => { const fs = Array.from(e.target.files); e.target.value = ''; if (!fs.length) return; if (view === 'upload') { pendRed.push(...fs); paintPicked(); } else exclusive(() => addRedFiles(fs)); };
  $('#btnStart').onclick = () => exclusive(async () => {
    if (!pendManu) return;
    // 新しい作業としてマイページに追加する（これまでの作業はそのまま一覧に残る）
    const red = pendRed.slice();
    view = 'upload-run';
    if (await openManuscript(pendManu)) { pendManu = null; pendRed = []; paintPicked(); if (red.length) await addRedFiles(red, { replace: true }); }
    else view = 'upload';
  });
  $('#myProjects').addEventListener('click', (e) => {
    const item = e.target.closest('[data-pid]'); if (!item) return;
    if (e.target.closest('[data-del]')) exclusive(() => deleteProject(item.dataset.pid));
    else if (e.target.closest('[data-open]')) exclusive(() => openProject(item.dataset.pid));
  });
  $('#brandHome').onclick = () => exclusive(showUploader);
  for (const id of ['dropManu', 'dropRed']) {
    const z = $('#' + id);
    z.addEventListener('dragover', (e) => { e.preventDefault(); e.stopPropagation(); z.classList.add('drag'); });
    z.addEventListener('dragleave', () => z.classList.remove('drag'));
    z.addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); z.classList.remove('drag'); takeFiles(e.dataTransfer.files); });
  }
  $('#btnRules').onclick = openRules;
  $('#btnUndo').onclick = undo; $('#btnRedo').onclick = redo;
  $('#btnAlignL').onclick = () => setAlign(''); $('#btnAlignC').onclick = () => setAlign('center'); $('#btnAlignR').onclick = () => setAlign('right');
  $('#btnToTable').onclick = toTable;
  // claude.ai の埋め込み表示では印刷ダイアログが使えないので隠す
  if (window.claude) $('#btnPrint').hidden = true;
  $('#btnPrint').onclick = () => window.print();
  $('#btnSaveProj').onclick = () => {
    const name = (S.settings.title || S.fileName || 'project').replace(/[\\/:*?"<>|]/g, '_').slice(0, 40);
    saveFile(new Blob([JSON.stringify(S)], { type: 'application/json' }), `${name}_編集データ.json`, '作業データを保存しました');
  };
  async function saveFile(blob, name, okMsg) {
    try {
      const r = await U.download(blob, name);
      if (r === 'saved') U.toast(okMsg, 5000);
      else if (r === 'declined') U.toast('保存をキャンセルしました');
      else if (r === 'busy') U.toast('別の保存確認が開いています。閉じてからもう一度押してください');
      else U.toast('この表示ではファイルを保存できません');
    } catch (err) { console.error(err); U.toast('保存できませんでした：' + err.message, 6000); }
  }
  async function loadProject(f) {
    try {
      const o = JSON.parse(await U.readFile(f, 'text')); if (!Array.isArray(o.blocks)) throw new Error('作業データの形式ではありません');
      // 既存の作業を上書きしないよう、新しい作業としてマイページに追加する
      await flushSave();
      const now = Date.now();
      S = Object.assign(newState(), o, { id: U.uid(), createdAt: o.createdAt || now, updatedAt: now });
      hist.undo = []; hist.redo = []; cur = null; sel.clear();
      redDocs = []; pdfCache.clear(); drawRed().catch(() => {});
      view = 'edit'; renderAll(); scheduleSave(); U.toast('作業データを開きました（マイページに追加）');
    } catch (err) { U.toast('開けませんでした：' + err.message, 6000); }
  }
  $('#btnLoadProj').onclick = () => $('#fileProj').click();
  $('#fileProj').onchange = async (e) => {
    const f = e.target.files[0]; e.target.value = ''; if (f) exclusive(() => loadProject(f));
  };
  $('#btnExport').onclick = async () => {
    if (!S.blocks.length) return U.toast('原稿がありません');
    U.toast('DOCXを生成中…', 8000);
    try {
      const blob = await DocxExport.build(S);
      const base = (S.fileName || 'manuscript').replace(/\.(docx|txt|text)$/i, '');
      const d = new Date(); const ymd = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
      await saveFile(blob, `${base}_POD用_${ymd}.docx`, '書き出しました。Wordで開いて「フィールドを更新」すると目次のページ番号が入ります');
    } catch (err) { console.error(err); U.toast('書き出しに失敗しました：' + err.message, 6000); }
  };
  $('#btnExportTxt').onclick = async () => {
    if (!S.blocks.length) return U.toast('原稿がありません');
    const base = (S.fileName || 'manuscript').replace(/\.(docx|txt|text)$/i, '');
    await saveFile(new Blob([DocxExport.toText(S)], { type: 'text/plain' }), `${base}_整形済み.txt`, 'テキストを書き出しました');
  };
  $('#btnSettings').onclick = () => {
    const s = S.settings;
    $('#setTitle').value = s.title || ''; $('#setAuthor').value = s.author || '';
    $('#setOpener').checked = s.opener !== false; $('#setHeader').checked = s.header !== false; $('#setFolio').checked = s.folio !== false; $('#setSize').value = s.size || 'A5';
    $('#dlgSettings').showModal();
  };
  $('#setSave').onclick = () => {
    Object.assign(S.settings, { title: $('#setTitle').value, author: $('#setAuthor').value, opener: $('#setOpener').checked, header: $('#setHeader').checked, folio: $('#setFolio').checked, size: $('#setSize').value });
    scheduleSave();
  };

  // ---------- キーボードショートカット ----------
  document.addEventListener('keydown', (e) => {
    const mod = e.metaKey || e.ctrlKey;
    if (mod && e.key.toLowerCase() === 'z' && !e.target.closest('dialog,input')) { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
    if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
    if (mod && e.key.toLowerCase() === 's') { e.preventDefault(); $('#btnSaveProj').click(); return; }
    if (mod && e.key.toLowerCase() === 'f') { e.preventDefault(); switchTab('find'); $('#fFind').focus(); $('#fFind').select(); return; }
    if (e.altKey && /^Digit[0-6]$/.test(e.code)) {
      const k = e.code.slice(5); const t = Object.entries(TYPES).find(([, v]) => v.key === k);
      if (t) { e.preventDefault(); setType(t[0]); }
      return;
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && sel.size && !e.target.closest('.ed,input,textarea,dialog')) { e.preventDefault(); deleteBlocks(targets()); }
    if (e.key === 'Escape') { clearSel(); $('#menu').hidden = true; $('#lightbox').hidden = true; }
  });

  // ---------- ドラッグ＆ドロップ（編集中：原稿は読み直し、PDF・画像は赤入れに追加）----------
  const wrap = $('#paper-wrap');
  wrap.addEventListener('dragover', (e) => e.preventDefault());
  wrap.addEventListener('drop', async (e) => {
    e.preventDefault();
    const fs = Array.from(e.dataTransfer.files);
    if (view !== 'edit') return takeFiles(fs);
    const red = fs.filter((f) => /\.(pdf|png|jpe?g)$/i.test(f.name));
    exclusive(async () => {
      if (red.length) await addRedFiles(red);
      const m = fs.find((f) => /\.(docx|txt|text)$/i.test(f.name)); if (m) await openManuscript(m);
      const j = fs.find((f) => /\.json$/i.test(f.name)); if (j) await loadProject(j);
    });
  });
  // 未保存の変更があるときは保存を試みつつ、離脱の確認を出す（非同期の保存は完了が保証されないため）
  window.addEventListener('beforeunload', (e) => {
    if (!isDirty() && !busy) return;
    flushSave();
    e.preventDefault(); e.returnValue = '';
  });

  // ---------- 起動：マイページ（過去の作業一覧＋新規アップロード）----------
  (async () => {
    await migrateLegacy().catch((e) => console.warn('migrate', e));
    renderAll();
  })();

  // デバッグ・自動テスト用
  window.PODEditor = { get state() { return S; }, openManuscript, openProject, listProjects, showMyPage: showUploader, addRedFiles, edit: () => { view = 'edit'; renderAll(); }, renderAll, setType, toTable, applyRules: () => $('#rulesApply').click(), select: (ids) => { clearSel(); ids.forEach((id) => { sel.add(id); elOf(id)?.classList.add('sel'); }); } };
})();

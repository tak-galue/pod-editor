// .docx → ブロック配列
// 変更履歴は「すべて承諾」した状態で取り込む（w:ins は採用、w:del / w:moveFrom は破棄）
const DocxImport = (() => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' };

  const kids = (el) => Array.from(el.children || []);
  const child = (el, name) => kids(el).find((c) => c.localName === name);
  const attr = (el, name) => el && (el.getAttributeNS(W, name) ?? el.getAttribute('w:' + name));

  async function parse(buf) {
    const zip = await JSZip.loadAsync(buf);
    const docFile = zip.file('word/document.xml');
    if (!docFile) throw new Error('Wordファイル（.docx）ではないようです');
    const xml = new DOMParser().parseFromString(await docFile.async('string'), 'application/xml');

    // リレーション（画像参照）
    const rels = {};
    const relFile = zip.file('word/_rels/document.xml.rels');
    if (relFile) {
      const rx = new DOMParser().parseFromString(await relFile.async('string'), 'application/xml');
      for (const r of rx.getElementsByTagName('Relationship')) rels[r.getAttribute('Id')] = r.getAttribute('Target');
    }
    // スタイルID → スタイル名
    const styleName = {};
    const stFile = zip.file('word/styles.xml');
    if (stFile) {
      const sx = new DOMParser().parseFromString(await stFile.async('string'), 'application/xml');
      for (const s of sx.getElementsByTagNameNS(W, 'style')) {
        const n = s.getElementsByTagNameNS(W, 'name')[0];
        styleName[attr(s, 'styleId')] = (n ? attr(n, 'val') : attr(s, 'styleId') || '').toLowerCase();
      }
    }

    const blocks = [], images = {}, pending = [];
    const imgByRid = {};
    function imageFor(rid) {
      if (!rid || !rels[rid]) return null;
      if (imgByRid[rid]) return imgByRid[rid];
      let target = rels[rid].replace(/^\//, '');
      const path = target.startsWith('word/') ? target : 'word/' + target;
      const f = zip.file(path);
      if (!f) return null;
      const ext = path.split('.').pop().toLowerCase();
      if (!MIME[ext]) return null; // emf/wmf などは非対応
      const id = 'img_' + Object.keys(images).length;
      images[id] = { name: path.split('/').pop(), mime: MIME[ext], data: null };
      pending.push(f.async('base64').then((b64) => (images[id].data = `data:${MIME[ext]};base64,${b64}`)));
      imgByRid[rid] = id;
      return id;
    }

    // 段落内の文字列と画像を順番に収集
    function collect(node, out) {
      for (const c of kids(node)) {
        const ln = c.localName;
        if (ln === 'del' || ln === 'moveFrom' || ln === 'delText' || ln === 'instrText' || ln === 'rt' || ln === 'pPr' || ln === 'rPr' || ln === 'Fallback') continue;
        if (ln === 't') out.push({ t: c.textContent });
        else if (ln === 'tab') out.push({ t: '\t' });
        else if (ln === 'br' || ln === 'cr') {
          if (attr(c, 'type') === 'page') out.push({ pb: true });
          else out.push({ t: '\n' });
        } else if (ln === 'noBreakHyphen') out.push({ t: '‑' });
        else if (ln === 'blip') {
          const id = imageFor(c.getAttributeNS(R, 'embed') || c.getAttribute('r:embed'));
          if (id) out.push({ img: id });
        } else if (ln === 'imagedata') {
          const id = imageFor(c.getAttributeNS(R, 'id') || c.getAttribute('r:id'));
          if (id) out.push({ img: id });
        } else if (ln === 'txbxContent') {
          // テキストボックス内の段落は改行区切りで取り込む
          for (const p of c.getElementsByTagNameNS(W, 'p')) { const o = []; collect(p, o); out.push(...o, { t: '\n' }); }
        } else if (c.children && c.children.length) collect(c, out);
      }
    }

    function pushText(text, extra) {
      blocks.push({ id: U.uid(), type: 'body', text, ...extra });
    }

    function handleP(p) {
      const pPr = child(p, 'pPr');
      let style = '', align = '';
      if (pPr) {
        const ps = child(pPr, 'pStyle'); if (ps) style = styleName[attr(ps, 'val')] || attr(ps, 'val');
        const jc = child(pPr, 'jc'); if (jc) { const v = attr(jc, 'val'); if (v === 'center') align = 'center'; else if (v === 'right' || v === 'end') align = 'right'; }
      }
      const items = []; collect(p, items);
      let buf = '';
      const flush = (force) => {
        const text = buf.replace(/\t/g, '　').replace(/\n+$/, '');
        if (text.length || force) pushText(text, { srcStyle: style, ...(align ? { align } : {}) });
        buf = '';
      };
      let emitted = false;
      for (const it of items) {
        if (it.t != null) buf += it.t;
        else if (it.img) { if (buf.trim()) { flush(); emitted = true; } else buf = ''; blocks.push({ id: U.uid(), type: 'figure', img: it.img, caption: '', note: '', width: 100 }); emitted = true; }
        else if (it.pb) { if (buf.trim()) { flush(); emitted = true; } else buf = ''; blocks.push({ id: U.uid(), type: 'pagebreak' }); emitted = true; }
      }
      if (buf.trim() || !emitted) flush(!emitted);
    }

    function handleTbl(tbl) {
      const rows = [];
      for (const tr of kids(tbl).filter((e) => e.localName === 'tr')) {
        const row = [];
        for (const tc of kids(tr).filter((e) => e.localName === 'tc')) {
          const parts = [];
          for (const p of kids(tc).filter((e) => e.localName === 'p')) { const o = []; collect(p, o); parts.push(o.map((x) => x.t || '').join('')); }
          row.push(parts.join('\n').trim());
        }
        rows.push(row);
      }
      if (rows.length) blocks.push({ id: U.uid(), type: 'table', rows });
    }

    function isTocSdt(sdt) {
      const pr = child(sdt, 'sdtPr');
      if (pr) for (const g of pr.getElementsByTagNameNS(W, 'docPartGallery')) if (/table of contents/i.test(attr(g, 'val') || '')) return true;
      const c = child(sdt, 'sdtContent');
      return !!c && /^\s*目次/.test(c.textContent || '');
    }

    function walk(node) {
      for (const el of kids(node)) {
        const ln = el.localName;
        if (ln === 'p') handleP(el);
        else if (ln === 'tbl') handleTbl(el);
        else if (ln === 'sdt') { if (isTocSdt(el)) blocks.push({ id: U.uid(), type: 'toc' }); else { const c = child(el, 'sdtContent'); if (c) walk(c); } }
        else if (ln === 'del' || ln === 'moveFrom' || ln === 'sectPr') continue;
        else if (el.children && el.children.length) walk(el); // customXml / ins / moveTo など
      }
    }
    const body = xml.getElementsByTagNameNS(W, 'body')[0];
    if (!body) throw new Error('本文が見つかりません');
    walk(body);
    await Promise.all(pending);
    for (const k of Object.keys(images)) { const s = await U.imgSize(images[k].data); images[k].w = s.w; images[k].h = s.h; }
    return { blocks, images };
  }
  // テキスト原稿：1行＝1段落。文字コードは UTF-8 → Shift_JIS の順で判定
  function parseText(buf) {
    const bytes = new Uint8Array(buf);
    let text = new TextDecoder('utf-8').decode(bytes);
    if (/\uFFFD/.test(text)) {
      try { const sj = new TextDecoder('shift_jis').decode(bytes); if ((sj.match(/\uFFFD/g) || []).length < (text.match(/\uFFFD/g) || []).length) text = sj; } catch { /* 非対応ブラウザはUTF-8のまま */ }
    }
    text = text.replace(/^﻿/, '').replace(/\r\n?/g, '\n');
    const blocks = text.split('\n').map((line) => {
      if (/^\s*(\f|＜改ページ＞|<改ページ>|\[改ページ\])\s*$/.test(line)) return { id: U.uid(), type: 'pagebreak' };
      return { id: U.uid(), type: 'body', text: line.replace(/\t/g, '　').replace(/\s+$/, ''), srcStyle: 'plain' };
    });
    return { blocks, images: {} };
  }
  return { parse, parseText };
})();

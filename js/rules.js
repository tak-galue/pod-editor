// POD組版の自動整形ルール
// 電子用原稿（R9）→ POD用最終版で実際に行われた変換を規則化したもの
const TYPES = {
  body: { label: '本文', key: '0' },
  chapter: { label: '章タイトル', key: '1' },
  chapterSub: { label: '章副題', key: '2' },
  h2: { label: '節見出し', key: '3' },
  lead: { label: 'リード', key: '4' },
  h3: { label: '小見出し◇', key: '5' },
  callout: { label: '強調ボックス', key: '6' },
  figure: { label: '図版' },
  table: { label: '表' },
  toc: { label: '目次' },
  pagebreak: { label: '改ページ' },
};
const TEXT_TYPES = ['body', 'chapterSub', 'h2', 'lead', 'h3', 'callout'];
const isText = (b) => TEXT_TYPES.includes(b.type);

const Rules = (() => {
  const D = '0-9０-９';
  const KN = '一二三四五六七八九十';
  const DASH = '－\\-−‐–—ー';
  const RE_CH = new RegExp(`^第\\s*([${D}${KN}]+)\\s*章[\\s　]*(.*)$`);
  const RE_CH_SP = new RegExp(`^第\\s*[${D}${KN}]+\\s*章([\\s　]|$)`); // 「第7章に関する…」は除外
  const RE_MARK =new RegExp(`^[【\\[]第[${D}${KN}]+章[】\\]]$`);
  const RE_SEC = new RegExp(`^[${D}]+\\s*[${DASH}]\\s*[${D}]+[\\s　]+\\S`);
  const RE_FIG = new RegExp(`^図\\s*([${D}]+)\\s*[${DASH}]\\s*([${D}]+)[\\s　]*(.*)$`);
  const RE_LEADDASH = /^[─━―—–-]{1,3}\s*/;
  const UNNUM = /^(はじめに|まえがき|序章|序文|最後に|あとがき|おわりに|著者プロフィール|著者紹介|参考文献|謝辞|付録)$/;
  const half = (s) => s.replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));

  // 章タイトル化：見出し1スタイル、または直前に【第N章】マーカーがある段落
  function toChapter(b) {
    const t = b.text.trim();
    const m = RE_CH.exec(t);
    if (m) return { id: b.id, type: 'chapter', num: `第${m[1]}章`, text: m[2].trim() };
    return { id: b.id, type: 'chapter', num: '', text: t };
  }

  const list = [
    {
      id: 'empty', label: '空段落を削除', desc: '行送り用の空段落を除去（段落間は組版スタイルで空ける）', on: true,
      run: (bs) => bs.filter((b) => !(isText(b) && !(b.text || '').trim())),
    },
    {
      id: 'tocText', label: '目次（テキスト原稿）', desc: '「目次」行と続くページ番号付きの行を自動目次に置き換え（ページ番号は書き出し時に再計算）', on: true,
      run: (bs) => {
        const out = [];
        for (let i = 0; i < bs.length; i++) {
          const b = bs[i];
          if (b.srcStyle === 'plain' && b.type === 'body' && /^目\s*次$/.test((b.text || '').trim())) {
            out.push({ id: b.id, type: 'toc' });
            let j = i + 1;
            const isTocLine = (x) => x && x.type === 'body' && x.srcStyle === 'plain' && (x.text || '').trim().length < 90 && /[0-9０-９]$/.test((x.text || '').trim()) && !/。/.test(x.text);
            while (j < bs.length && (isTocLine(bs[j]) || (bs[j].type === 'body' && !(bs[j].text || '').trim() && isTocLine(bs[j + 1])))) j++;
            i = j - 1; continue;
          }
          out.push(b);
        }
        return out;
      },
    },
    {
      id: 'chapter', label: '章タイトル', desc: '「第N章　…」「はじめに」「あとがき」等を章扉に。柱（ヘッダー）にも使用', on: true,
      run: (bs) => bs.map((b, i) => {
        if (b.type !== 'body') return b;
        const t = (b.text || '').trim();
        let prev = null; for (let k = i - 1; k >= 0; k--) { if (bs[k].type !== 'pagebreak' && (bs[k].text || '').trim()) { prev = bs[k]; break; } }
        const h1 = b.srcStyle === 'heading 1';
        // テキスト原稿はスタイルがないので「第N章　…」の短い行・「はじめに」等の単独行を章とみなす
        const plain = b.srcStyle === 'plain' && t.length < 60 && !/[。」]$/.test(t) && (RE_CH_SP.test(t) || UNNUM.test(t));
        if ((h1 && t.length < 60) || plain || (prev && RE_MARK.test((prev.text || '').trim()) && RE_CH.test(t)) || (UNNUM.test(t) && b.srcStyle && /heading/.test(b.srcStyle)))
          return toChapter(b);
        return b;
      }),
    },
    {
      id: 'marker', label: '【第N章】マーカーを削除', desc: '原稿上の仮マーカー行を除去', on: true,
      run: (bs) => bs.filter((b) => !(isText(b) && RE_MARK.test((b.text || '').trim()))),
    },
    {
      id: 'pagebreak', label: '手動改ページを削除', desc: '章ごとに自動で改ページされるため、原稿内の改ページを除去', on: true,
      run: (bs) => bs.filter((b) => b.type !== 'pagebreak'),
    },
    {
      id: 'chapterSub', label: '章副題', desc: '章タイトル直後の「～…」「──…」を章副題に（記号は除去）', on: true,
      run: (bs) => bs.map((b, i) => {
        if (b.type !== 'body' || i === 0) return b;
        const p = bs[i - 1];
        if (p.type !== 'chapter' && p.type !== 'chapterSub') return b;
        const t = (b.text || '').trim();
        if (p.type === 'chapter' && /^[～〜~─―—]+/.test(t)) return { ...b, type: 'chapterSub', text: t.replace(/^[～〜~─―—\s]+/, '') };
        if (p.type === 'chapterSub' && t.length <= 40 && !/[。」]$/.test(t)) return { ...b, type: 'chapterSub', text: t };
        return b;
      }),
    },
    {
      id: 'h2', label: '節見出し', desc: '見出し2スタイル、または「1－1　…」で始まる短い行', on: true,
      run: (bs) => bs.map((b) => {
        if (b.type !== 'body') return b;
        const t = (b.text || '').trim();
        if (!t || t.length > 70 || /。$/.test(t)) return b;
        if (b.srcStyle === 'heading 2' || RE_SEC.test(t)) return { ...b, type: 'h2', text: t.replace(/\s+/g, (s) => (s.includes('\n') ? '' : s)) };
        return b;
      }),
    },
    {
      id: 'lead', label: 'リード（サブタイトル）', desc: '節見出し直後の「──…」を「─ … ─」形式のリードに', on: true,
      run: (bs) => bs.map((b, i) => {
        if (b.type !== 'body' || i === 0 || bs[i - 1].type !== 'h2') return b;
        const t = (b.text || '').trim();
        if (!RE_LEADDASH.test(t) || t.length > 80) return b;
        const core = t.replace(RE_LEADDASH, '').replace(/[\s　]*[─━―—–-]{1,3}$/, '');
        return { ...b, type: 'lead', text: `─ ${core} ─` };
      }),
    },
    {
      id: 'h3', label: '小見出し（◇）', desc: '「◇」で始まる行・箇条書きスタイルの短い行を小見出しに。「◇　」に統一', on: true,
      run: (bs) => bs.map((b) => {
        if (b.type !== 'body') return b;
        const t = (b.text || '').trim();
        const marked = /^[◇◆]/.test(t);
        if (!marked && !(b.srcStyle === 'list paragraph' && t.length <= 50 && !/。$/.test(t))) return b;
        const core = t.replace(/^[◇◆][\s　]*/, '');
        return { ...b, type: 'h3', text: (core.startsWith('「') ? '◇ ' : '◇　') + core };
      }),
    },
    {
      id: 'figure', label: '図版キャプション', desc: '画像の前後にある「図1－1　…」を図版のキャプションとして統合（「図1-1」表記に）', on: true,
      run: (bs) => {
        const out = bs.slice(); const drop = new Set();
        out.forEach((b, i) => {
          if (b.type !== 'figure' || b.caption) return;
          const cand = [i + 1, i + 2, i + 3, i - 1, i - 2].filter((k) => k >= 0 && k < out.length);
          for (const k of cand) {
            const c = out[k];
            if (drop.has(c.id) || c.type !== 'body') continue;
            const m = RE_FIG.exec((c.text || '').trim());
            if (!m) continue;
            out[i] = { ...b, caption: `図${half(m[1])}-${half(m[2])}　${m[3].trim()}` };
            drop.add(c.id); break;
          }
        });
        return out.filter((b) => !drop.has(b.id));
      },
    },
    {
      id: 'space', label: '空白の正規化', desc: '全角空白の重複を1つに／和文間の半角空白を全角に（例：西田 徹 → 西田　徹）／行末空白を除去', on: true,
      run: (bs) => bs.map((b) => {
        if (!isText(b) && b.type !== 'chapter') return b;
        const t = (b.text || '')
          .replace(/[  ]*　[  ]*/g, '　').replace(/　{2,}/g, '　')
          .replace(/([぀-ヿ一-鿿])[  ]+([぀-ヿ一-鿿])/g, '$1　$2')
          .replace(/[ \t　]+$/gm, '').replace(/^[ \t]+/gm, '');
        return t === b.text ? b : { ...b, text: t };
      }),
    },
    {
      id: 'callout', label: '強調ボックス候補', desc: '「…」だけの短い独立段落を強調ボックスに（作品により要確認）', on: false,
      run: (bs) => bs.map((b) => {
        if (b.type !== 'body') return b;
        const t = (b.text || '').trim();
        if (/^「[^「」]{4,40}」$/.test(t)) return { ...b, type: 'callout', text: t };
        return b;
      }),
    },
    {
      id: 'dash', label: '「――」を「……」に', desc: '会話・独白中のダッシュを三点リーダーへ（表記統一する場合）', on: false,
      run: (bs) => bs.map((b) => (isText(b) && /――/.test(b.text) ? { ...b, text: b.text.replace(/――/g, '……') } : b)),
    },
  ];

  // 変更段落数を数える（ID単位での比較）
  function countChanges(before, after) {
    const m = new Map(after.map((b) => [b.id, JSON.stringify(b)]));
    let n = 0;
    for (const b of before) { const a = m.get(b.id); if (a === undefined || a !== JSON.stringify(b)) n++; }
    return n;
  }
  function preview(blocks, enabled) {
    let cur = blocks; const counts = {};
    for (const r of list) {
      if (!enabled[r.id]) { counts[r.id] = 0; continue; }
      const next = r.run(cur); counts[r.id] = countChanges(cur, next); cur = next;
    }
    return { blocks: cur, counts };
  }
  return { list, preview, RE_CH };
})();

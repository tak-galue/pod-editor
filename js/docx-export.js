// ブロック配列 → POD用 .docx
// 仕様は「POD用最終版」を踏襲：A5／游明朝10.5pt／見出し游ゴシック／ワインレッド(#7A1F2E)
// 章ごとにセクションを分け、柱（ヘッダー）に章タイトル、フッターにノンブル、目次はTOCフィールド
const DocxExport = (() => {
  const NS = 'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  const BRAND = '7A1F2E';
  const SIZES = {
    A5: { w: 8390, h: 11905, top: 1559, right: 850, bottom: 1417, left: 1133 },
    B6: { w: 7257, h: 10319, top: 1300, right: 750, bottom: 1200, left: 1000 },
    SHIROKU: { w: 7201, h: 10659, top: 1300, right: 750, bottom: 1200, left: 1000 },
  };
  const GOTHIC = '<w:rFonts w:ascii="Yu Gothic" w:eastAsia="Yu Gothic" w:hAnsi="Yu Gothic" w:cs="Yu Gothic"/>';
  const esc = U.esc;

  function runs(text, rPr = '') {
    const parts = String(text ?? '').split('\n');
    return parts.map((p, i) => (i ? `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:br/></w:r>` : '') + (p ? `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(p)}</w:t></w:r>` : '')).join('');
  }
  function para(text, { style, jc, extraPPr = '', rPr = '', pre = '', post = '' } = {}) {
    const ppr = (style ? `<w:pStyle w:val="${style}"/>` : '') + extraPPr + (jc ? `<w:ind w:firstLine="0"/><w:jc w:val="${jc}"/>` : '');
    return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${pre}${runs(text, rPr)}${post}</w:p>`;
  }
  const field = (instr) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;

  function stylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="游明朝" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="160" w:line="300" w:lineRule="auto"/><w:ind w:firstLineChars="100" w:firstLine="210"/></w:pPr><w:rPr><w:rFonts w:ascii="游明朝" w:eastAsia="游明朝" w:hAnsi="游明朝" w:cs="游明朝"/><w:color w:val="1A1A1A"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:after="300" w:line="240" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr>${GOTHIC}<w:b/><w:bCs/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:shd w:val="clear" w:color="auto" w:fill="${BRAND}"/><w:spacing w:before="600" w:after="360" w:line="440" w:lineRule="exact"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr>${GOTHIC}<w:b/><w:bCs/><w:color w:val="FFFFFF"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="320" w:after="20" w:line="240" w:lineRule="auto"/><w:ind w:left="480" w:hanging="480" w:firstLineChars="0"/><w:jc w:val="left"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr>${GOTHIC}<w:b/><w:bCs/><w:color w:val="${BRAND}"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="ChapterLabel"><w:name w:val="ChapterLabel"/><w:basedOn w:val="Normal"/><w:next w:val="Heading1"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="1800" w:after="240" w:line="240" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr><w:rPr>${GOTHIC}<w:b/><w:bCs/><w:color w:val="FFFFFF"/><w:sz w:val="26"/><w:szCs w:val="26"/><w:shd w:val="clear" w:color="auto" w:fill="${BRAND}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="ChapterSubtitle"><w:name w:val="ChapterSubtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:after="0"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr><w:rPr>${GOTHIC}<w:color w:val="4D4D4D"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="ChapterRule"><w:name w:val="ChapterRule"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="36" w:space="1" w:color="${BRAND}"/></w:pBdr><w:spacing w:after="400" w:line="200" w:lineRule="exact"/><w:ind w:firstLineChars="0" w:firstLine="0"/></w:pPr><w:rPr><w:sz w:val="12"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="LeadSubtitle"><w:name w:val="LeadSubtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:after="200"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr><w:rPr>${GOTHIC}<w:color w:val="${BRAND}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Callout"><w:name w:val="Callout"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="24" w:space="4" w:color="${BRAND}"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="F4E3E5"/><w:spacing w:before="160" w:line="320" w:lineRule="atLeast"/><w:ind w:left="120" w:right="120" w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr><w:rPr>${GOTHIC}<w:b/><w:bCs/><w:color w:val="${BRAND}"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FigCaption"><w:name w:val="FigCaption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr><w:rPr>${GOTHIC}<w:b/><w:bCs/><w:color w:val="${BRAND}"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FigImage"><w:name w:val="FigImage"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FigNote"><w:name w:val="FigNote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="240"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr><w:rPr>${GOTHIC}<w:color w:val="555555"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="TocTitle"><w:name w:val="TocTitle"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="400" w:after="400"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr><w:rPr>${GOTHIC}<w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{CW}"/></w:tabs><w:spacing w:before="120" w:after="40"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr><w:rPr>${GOTHIC}<w:b/><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{CW}"/></w:tabs><w:spacing w:after="20" w:line="240" w:lineRule="auto"/><w:ind w:left="210" w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="17"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="2" w:color="999999"/></w:pBdr><w:tabs><w:tab w:val="center" w:pos="4252"/><w:tab w:val="right" w:pos="8504"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/></w:pPr><w:rPr>${GOTHIC}<w:color w:val="666666"/><w:sz w:val="15"/><w:szCs w:val="15"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/></w:pPr><w:rPr>${GOTHIC}<w:color w:val="444444"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr></w:style>
<w:style w:type="table" w:customStyle="1" w:styleId="InfoTable"><w:name w:val="InfoTable"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="CFCAC3"/><w:left w:val="single" w:sz="4" w:space="0" w:color="CFCAC3"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="CFCAC3"/><w:right w:val="single" w:sz="4" w:space="0" w:color="CFCAC3"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="CFCAC3"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="CFCAC3"/></w:tblBorders><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>`;
  }

  function hdrXml(inner) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr ${NS}>${inner}</w:hdr>`;
  }
  function ftrXml(inner) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr ${NS}>${inner}</w:ftr>`;
  }

  async function build(state) {
    const s = state.settings || {};
    const pg = SIZES[s.size] || SIZES.A5;
    const CW = pg.w - pg.left - pg.right; // 版面幅
    const zip = new JSZip();
    const rels = []; // {id,type,target}
    const ctypes = [];
    let rid = 10, docPrId = 1;
    const addRel = (type, target) => { const id = 'rId' + rid++; rels.push({ id, type, target }); return id; };
    const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

    // ---- 画像 ----
    const imgRel = {};
    let imgN = 1;
    function imageRel(imgId) {
      if (imgRel[imgId]) return imgRel[imgId];
      const im = state.images[imgId]; if (!im || !im.data) return null;
      const ext = ({ 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'png' })[im.mime] || 'png';
      const name = `image${imgN++}.${ext}`;
      zip.file('word/media/' + name, im.data.split(',')[1], { base64: true });
      imgRel[imgId] = addRel(RT + 'image', 'media/' + name);
      return imgRel[imgId];
    }
    function drawing(imgId, pct) {
      const r = imageRel(imgId); if (!r) return '';
      const im = state.images[imgId];
      const maxW = (CW - 200) * 635; // EMU
      const maxH = (pg.h - pg.top - pg.bottom - 1800) * 635;
      let cx = Math.round((im.w || 800) * 9525), cy = Math.round((im.h || 600) * 9525);
      const target = maxW * Math.min(100, Math.max(10, pct || 100)) / 100;
      const k = Math.min(target / cx, maxH / cy, 1e9);
      cx = Math.round(cx * k); cy = Math.round(cy * k);
      const id = docPrId++;
      return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="図 ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${esc(im.name || 'image')}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${r}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    }

    // ---- 目次エントリ ----
    const blocks = state.blocks;
    const tocEntries = [];
    for (const b of blocks) {
      if (b.type === 'chapter') tocEntries.push({ l: 1, t: (b.num ? b.num + '　' : '') + b.text });
      else if (b.type === 'h2') tocEntries.push({ l: 2, t: b.text });
    }

    // ---- ヘッダー／フッター ----
    let hfN = 1;
    function addPart(kind, xml) {
      const name = `${kind}${hfN++}.xml`;
      zip.file('word/' + name, xml);
      ctypes.push(`<Override PartName="/word/${name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind === 'header' ? 'header' : 'footer'}+xml"/>`);
      return addRel(RT + kind, name);
    }
    const emptyHdr = () => addPart('header', hdrXml('<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>'));
    const emptyFtr = () => addPart('footer', ftrXml('<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>'));
    const pageFtr = (jc) => addPart('footer', ftrXml(`<w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="${jc}"/></w:pPr>${field('PAGE')}</w:p>`));

    function sectPr(sec, isFirst, isLast) {
      let refs = '';
      if (sec.kind === 'cover') {
        const h = emptyHdr(), f = emptyFtr();
        refs = ['default', 'even', 'first'].map((t) => `<w:headerReference w:type="${t}" r:id="${h}"/>`).join('') + ['default', 'even', 'first'].map((t) => `<w:footerReference w:type="${t}" r:id="${f}"/>`).join('');
      } else {
        const head = s.header !== false && sec.running ? sec.running : '';
        const hOdd = head ? addPart('header', hdrXml(`<w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="right"/></w:pPr>${runs(head)}</w:p>`)) : emptyHdr();
        const hEven = head ? addPart('header', hdrXml(`<w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="left"/></w:pPr>${runs(head)}</w:p>`)) : hOdd;
        const hFirst = emptyHdr();
        const fOdd = s.folio !== false ? pageFtr('right') : emptyFtr();
        const fEven = s.folio !== false ? pageFtr('left') : fOdd;
        const fFirst = sec.opener ? emptyFtr() : fOdd;
        refs = `<w:headerReference w:type="default" r:id="${hOdd}"/><w:headerReference w:type="even" r:id="${hEven}"/><w:headerReference w:type="first" r:id="${hFirst}"/><w:footerReference w:type="default" r:id="${fOdd}"/><w:footerReference w:type="even" r:id="${fEven}"/><w:footerReference w:type="first" r:id="${fFirst}"/>`;
      }
      return `<w:sectPr>${refs}<w:type w:val="nextPage"/><w:pgSz w:w="${pg.w}" w:h="${pg.h}"/><w:pgMar w:top="${pg.top}" w:right="${pg.right}" w:bottom="${pg.bottom}" w:left="${pg.left}" w:header="680" w:footer="680" w:gutter="0"/>${isFirst ? '<w:pgNumType w:start="1"/>' : ''}<w:cols w:space="720"/>${sec.kind === 'cover' ? '' : '<w:titlePg/>'}<w:docGrid w:linePitch="360"/></w:sectPr>`;
    }

    // ---- 本文をセクション単位で組み立て ----
    const sections = [];
    let cur = { kind: 'cover', running: '', opener: false, xml: [] };
    let pendingBreak = false;
    const pbPPr = () => (pendingBreak ? ((pendingBreak = false), '<w:pageBreakBefore/>') : '');

    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.type === 'chapter') {
        if (cur.xml.length || cur.kind !== 'cover') sections.push(cur);
        const title = (b.num ? b.num + '　' : '') + b.text;
        cur = { kind: 'chapter', running: b.runningHead || title, opener: !!b.num && s.opener !== false, xml: [] };
        pendingBreak = false;
        const tc = field(`TC "${title.replace(/"/g, '”')}" \\l 1`);
        if (b.num) {
          cur.xml.push(para(`　${b.num}　`, { style: 'ChapterLabel' }));
          cur.xml.push(para(b.text, { style: 'Heading1', post: tc }));
        } else {
          cur.xml.push(para(b.text, { style: 'Heading1', extraPPr: '<w:spacing w:before="600"/>', post: tc }));
        }
        // 章副題を続けて取り込み
        let j = i + 1; const subs = [];
        while (j < blocks.length && blocks[j].type === 'chapterSub') subs.push(blocks[j++]);
        subs.forEach((sb) => cur.xml.push(para(sb.text, { style: 'ChapterSubtitle' })));
        if (b.num) cur.xml.push('<w:p><w:pPr><w:pStyle w:val="ChapterRule"/></w:pPr></w:p>');
        i = j - 1;
        if (cur.opener) pendingBreak = true;
        continue;
      }
      switch (b.type) {
        case 'body': case 'chapterSub':
          cur.xml.push(para(b.text, { jc: b.align === 'center' ? 'center' : b.align === 'right' ? 'right' : '', extraPPr: pbPPr() })); break;
        case 'h2': cur.xml.push(para(b.text, { style: 'Heading2', extraPPr: pbPPr() })); break;
        case 'h3': cur.xml.push(para(b.text, { style: 'Heading3', extraPPr: pbPPr() })); break;
        case 'lead': cur.xml.push(para(b.text, { style: 'LeadSubtitle', extraPPr: pbPPr() })); break;
        case 'callout': cur.xml.push(para(b.text, { style: 'Callout', extraPPr: pbPPr() })); break;
        case 'pagebreak': pendingBreak = true; break;
        case 'figure': {
          if (b.caption) cur.xml.push(para(b.caption, { style: 'FigCaption', extraPPr: pbPPr() }));
          cur.xml.push(`<w:p><w:pPr><w:pStyle w:val="FigImage"/>${pbPPr()}</w:pPr>${drawing(b.img, b.width)}</w:p>`);
          if (b.note) cur.xml.push(para(b.note, { style: 'FigNote' }));
          break;
        }
        case 'table': {
          if (pendingBreak) cur.xml.push(`<w:p><w:pPr>${pbPPr()}<w:spacing w:after="0"/></w:pPr></w:p>`);
          const cols = Math.max(1, ...b.rows.map((r) => r.length));
          const first = cols === 2 ? Math.round(CW * 0.26) : Math.floor(CW / cols);
          const widths = cols === 2 ? [first, CW - first] : Array(cols).fill(first);
          let t = `<w:tbl><w:tblPr><w:tblStyle w:val="InfoTable"/><w:tblW w:w="${CW}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
          for (const row of b.rows) {
            t += '<w:tr>';
            for (let c = 0; c < cols; c++) {
              const head = c === 0 && cols === 2;
              t += `<w:tc><w:tcPr><w:tcW w:w="${widths[c]}" w:type="dxa"/>${head ? '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' : ''}</w:tcPr>${para(row[c] || '', { extraPPr: '<w:spacing w:after="0"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/>', rPr: (head ? GOTHIC + '<w:b/><w:color w:val="4D4D4D"/>' : '') + '<w:sz w:val="19"/><w:szCs w:val="19"/>' })}</w:tc>`;
            }
            t += '</w:tr>';
          }
          cur.xml.push(t + '</w:tbl>');
          cur.xml.push('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>');
          break;
        }
        case 'toc': {
          // 目次は柱なしの独立セクションにする（直前の章の柱が目次頁に出ないように）
          if (cur.xml.length) sections.push(cur);
          cur = { kind: 'front', running: '', opener: false, xml: [] };
          pendingBreak = false;
          cur.xml.push(para('目次', { style: 'TocTitle' }));
          const begin = `<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "2-2" \\f \\h \\z </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
          const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
          if (!tocEntries.length) { cur.xml.push(`<w:p>${begin}${end}</w:p>`); break; }
          tocEntries.forEach((e, k) => {
            cur.xml.push(`<w:p><w:pPr><w:pStyle w:val="TOC${e.l}"/></w:pPr>${k === 0 ? begin : ''}${runs(e.t)}<w:r><w:tab/></w:r>${k === tocEntries.length - 1 ? end : ''}</w:p>`);
          });
          break;
        }
      }
    }
    sections.push(cur);

    // セクション区切りは各セクション末尾の段落の pPr に sectPr を置く
    const bodyParts = [];
    sections.forEach((sec, k) => {
      const isLast = k === sections.length - 1;
      const sp = sectPr(sec, k === 0, isLast);
      bodyParts.push(...sec.xml);
      if (isLast) bodyParts.push(sp);
      else bodyParts.push(`<w:p><w:pPr><w:spacing w:after="0" w:line="20" w:lineRule="exact"/>${sp}</w:pPr></w:p>`);
    });
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${NS}><w:body>${bodyParts.join('')}</w:body></w:document>`;

    // ---- パッケージ ----
    const stylesRel = addRel(RT + 'styles', 'styles.xml');
    const settingsRel = addRel(RT + 'settings', 'settings.xml');
    void stylesRel; void settingsRel;
    zip.file('word/document.xml', docXml);
    zip.file('word/styles.xml', stylesXml().replace(/\{CW\}/g, String(CW)));
    zip.file('word/settings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:updateFields w:val="true"/><w:evenAndOddHeaders/><w:defaultTabStop w:val="840"/><w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
    zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"/>`).join('')}</Relationships>`);
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(s.title || '')}</dc:title><dc:creator>${esc(s.author || '')}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>POD原稿エディタ</Application></Properties>`);
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="bmp" ContentType="image/bmp"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>${ctypes.join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE' });
  }
  // 整形後の本文をプレーンテキストに（テキストで再読込すると同じ見出し判定がかかる形式）
  function toText(state) {
    const out = [];
    for (const b of state.blocks) {
      switch (b.type) {
        case 'chapter': out.push('', (b.num ? b.num + '　' : '') + b.text); break;
        case 'figure': out.push(`［${b.caption || '図'}］` + (b.note ? '　' + b.note : '')); break;
        case 'table': b.rows.forEach((r) => out.push(r.join('　'))); break;
        case 'toc': out.push('目次'); break;
        case 'pagebreak': out.push('＜改ページ＞'); break;
        case 'h2': out.push('', b.text); break;
        default: out.push(b.text || '');
      }
    }
    return out.join('\n').replace(/^\n+/, '') + '\n';
  }
  return { build, toText };
})();

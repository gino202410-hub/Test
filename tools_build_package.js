const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, LevelFormat, BorderStyle,
  Table, TableRow, TableCell, WidthType, ShadingType, VerticalAlign, ImageRun, PageOrientation,
  Footer, PageNumber, PageBreak,
} = require('docx');

const S = '/tmp/claude-0/-home-user-Test/38d207b7-bc95-5028-af5d-189dfd6f6d65/scratchpad/';
const W = '/home/user/Test/';
const OUT = W + 'THE_NEW_MISSION_수정방향제안_1006.docx';

const FONT = { ascii: 'Malgun Gothic', eastAsia: 'Malgun Gothic', hAnsi: 'Malgun Gothic', cs: 'Malgun Gothic' };
const BODY = 21;
const PORTRAIT_TEXT_WIDTH = 8840; // DXA; A4 (11906) minus 1533*2 margins

function runs(text, base = {}) {
  const out = [];
  text.split('**').forEach((p, i) => {
    if (!p) return;
    out.push(new TextRun({ text: p, bold: i % 2 === 1 ? true : base.bold, italics: base.italics, font: FONT, size: base.size || BODY, color: base.color }));
  });
  return out;
}
const splitRow = l => l.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(s => s.trim());

function makeTable(rows) {
  const ncol = rows[0].length;
  const widths = [];
  if (ncol >= 3) {
    const first = Math.round(PORTRAIT_TEXT_WIDTH * 0.18);
    const rest = Math.round((PORTRAIT_TEXT_WIDTH - first) / (ncol - 1));
    widths.push(first); for (let i = 1; i < ncol; i++) widths.push(rest);
  } else for (let i = 0; i < ncol; i++) widths.push(Math.round(PORTRAIT_TEXT_WIDTH / ncol));
  const total = widths.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: 'BBBBBB' };
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
    rows: rows.map((cells, ri) => new TableRow({
      tableHeader: ri === 0,
      children: cells.map((c, ci) => new TableCell({
        width: { size: widths[ci], type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP,
        shading: ri === 0 ? { type: ShadingType.CLEAR, fill: 'EEEEEE', color: 'auto' } : undefined,
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: [new Paragraph({ spacing: { after: 0, line: 270 }, children: runs(c, { bold: ri === 0 ? true : undefined, size: 18 }) })],
      })),
    })),
  });
}

// markdown → blocks. opts.h1: 'break' (page break before H1) | 'plain'
function md(text, opts = {}) {
  const out = []; let tbl = null;
  const flush = () => { if (tbl && tbl.length) { out.push(makeTable(tbl)); out.push(new Paragraph({ spacing: { after: 120 }, children: [] })); } tbl = null; };
  for (const raw of text.split('\n')) {
    const line = raw.replace(/\s+$/, '');
    if (/^\s*\|.*\|\s*$/.test(line)) {
      const cells = splitRow(line);
      if (cells.every(c => /^:?-{2,}:?$/.test(c))) continue;
      (tbl = tbl || []).push(cells); continue;
    } else if (tbl) flush();
    if (!line.trim()) continue;
    if (line === '---') { out.push(new Paragraph({ spacing: { before: 200, after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '999999', space: 1 } }, children: [] })); continue; }
    if (line.startsWith('# ')) { out.push(new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: opts.h1 !== 'plain', spacing: { before: 0, after: 320 }, children: [new TextRun({ text: line.slice(2), bold: true, font: FONT, size: 32 })] })); continue; }
    if (line.startsWith('## ')) { out.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 360, after: 160 }, keepNext: true, children: [new TextRun({ text: line.slice(3), bold: true, font: FONT, size: 26 })] })); continue; }
    if (line.startsWith('### ')) { out.push(new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 360, after: 180 }, keepNext: true, children: [new TextRun({ text: line.slice(4), bold: true, font: FONT, size: 23 })] })); continue; }
    if (line.startsWith('- ')) { out.push(new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80 }, children: runs(line.slice(2)) })); continue; }
    const m = line.match(/^\*\*([^*]+)\*\*\s{2,}(.*)$/);
    if (m) { out.push(new Paragraph({ indent: { left: 1200 }, spacing: { before: 60, after: 160 }, children: [new TextRun({ text: m[1], bold: true, font: FONT, size: BODY }), new TextRun({ text: '    ' + m[2], font: FONT, size: BODY })] })); continue; }
    if (line.startsWith('자막:')) { out.push(new Paragraph({ spacing: { after: 160 }, children: runs(line, { italics: true, color: '444444' }) })); continue; }
    if (line === '끝.') { out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 480 }, children: [new TextRun({ text: '끝.', bold: true, font: FONT, size: BODY })] })); continue; }
    if (line.startsWith('출처:')) { out.push(new Paragraph({ spacing: { after: 160, line: 280 }, children: runs(line, { size: 17, color: '666666' }) })); continue; }
    out.push(new Paragraph({ spacing: { after: 160, line: 320 }, children: runs(line) }));
  }
  flush();
  return out;
}

const read = f => fs.readFileSync(f, 'utf8');
const between = (t, a, b) => { const i = t.indexOf(a); if (i < 0) throw new Error('missing ' + a); const j = b ? t.indexOf(b, i) : -1; return t.slice(i, j < 0 ? undefined : j); };

// ---------- sources ----------
const briefing = read(W + 'THE_NEW_MISSION_브리핑.md');
const synA = read(W + 'THE_NEW_MISSION_시놉시스_v5.md');
const synB = read(W + 'THE_NEW_MISSION_벌집_시놉시스_v2.md');
const sci = read(W + 'THE_NEW_MISSION_과학·비교_간략판.md');
const scen = read('/root/.claude/uploads/38d207b7-bc95-5028-af5d-189dfd6f6d65/cc4d48af-THE_NEW_MISSION________________v8_9_________.md');
const meta = JSON.parse(read(S + 'design_meta.json'));
// section intros from THE_NEW_MISSION_장머리말.md ("## N장 머리말" blocks)
const introsSrc = read(W + 'THE_NEW_MISSION_장머리말.md');
const intros = {};
introsSrc.split(/^## (\d)장 머리말\s*$/m).slice(1).forEach((v, i, arr) => { if (i % 2 === 0) intros[v] = arr[i + 1].trim(); });
['2', '3', '5'].forEach(k => { if (!intros[k]) throw new Error('intro missing: ' + k); });

const synABody = between(synA, '## 로그라인', '## 30초 버전');
const synBBody = between(synB, '## 로그라인', '## 30초 버전');
const sciBody = between(sci, '## 과학적 고려', null);
let scenBody = between(scen, '## 표기', '## 검증·미정 메모');
scenBody = scenBody.replace(/^# /gm, '## ').replace(/^## 표기/m, '### 표기').replace(/^## 나오는 이/m, '### 나오는 이');

// ---------- cover ----------
const cover = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 3600, after: 200 }, children: [new TextRun({ text: 'THE NEW MISSION', bold: true, font: FONT, size: 56 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: '수정 방향 제안서', font: FONT, size: 30 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 1200 }, children: [new TextRun({ text: '2026. 10. 06', font: FONT, size: BODY, color: '555555' })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: '2026 시네마바이브랜드 : HANWHA', font: FONT, size: BODY, color: '555555' })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 1600 }, children: [new TextRun({ text: '기획 시네마직립보행 · 감독 이진호 · 프로듀서 은종훈 · 원안 구정회', font: FONT, size: BODY, color: '555555' })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [new TextRun({ text: '차례', bold: true, font: FONT, size: 24 })] }),
  ...['1  브리핑. 피칭 판에서 달라진 것', '2  시놉시스. 주안 「교신」 · 병렬안 「벌집」', '3  디자인. 꿀벌1호 변형 시안 여덟', '4  과학적 고려와 다른 작품과의 거리', '5  부록. 시나리오 「교신」 v8.9']
    .map(t => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: t, font: FONT, size: BODY })] })),
];

// ---------- portrait section 1: briefing + synopses ----------
const part1 = [
  ...cover,
  ...md(briefing),
  new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, spacing: { after: 320 }, children: [new TextRun({ text: '2. 시놉시스', bold: true, font: FONT, size: 32 })] }),
  ...md(intros['2'], { h1: 'plain' }),
  new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 160 }, children: [new TextRun({ text: '2-1  주안 「교신」 · SF 휴먼 드라마', bold: true, font: FONT, size: 28 })] }),
  ...md(synABody),
  new Paragraph({ heading: HeadingLevel.HEADING_2, pageBreakBefore: true, spacing: { before: 0, after: 160 }, children: [new TextRun({ text: '2-2  병렬안 「벌집」 · SF 어드벤처', bold: true, font: FONT, size: 28 })] }),
  ...md(synBBody),
  // design intro (portrait)
  new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, spacing: { after: 320 }, children: [new TextRun({ text: '3. 디자인. 꿀벌1호 변형 시안 여덟', bold: true, font: FONT, size: 32 })] }),
  ...md(intros['3'], { h1: 'plain' }),
  ...meta.map((m, i) => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80 }, children: runs(`**${m.n}.** ${m.title}`) })),
];

// ---------- landscape section: 8 design pages ----------
const landscape = [];
// usable area: A4 landscape 16838 x 11906 DXA minus margins 600 each side → 15638 x 10706 DXA → inches: 10.86 x 7.43
const maxW = 10.8 * 96, maxH = 6.6 * 96; // px at 96dpi; leaves room for the caption line and footer
meta.forEach((m, i) => {
  const img = fs.readFileSync(S + `design_hi_p${i + 1}.jpg`);
  const ratio = m.w / m.h;
  let w = maxW, h = w / ratio;
  if (h > maxH) { h = maxH; w = h * ratio; }
  landscape.push(new Paragraph({
    pageBreakBefore: i > 0, spacing: { after: 60 },
    children: [new TextRun({ text: `3-${i + 1}  ${m.n}`, bold: true, font: FONT, size: 22 }), new TextRun({ text: `   ${m.title}`, font: FONT, size: 20, color: '555555' })],
  }));
  landscape.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new ImageRun({ type: 'jpg', data: img, transformation: { width: Math.round(w), height: Math.round(h) } })] }));
});

// ---------- portrait section 2: science + appendix ----------
const part3 = [
  new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { after: 320 }, children: [new TextRun({ text: '4. 과학적 고려와 다른 작품과의 거리', bold: true, font: FONT, size: 32 })] }),
  ...md(sciBody),
  new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, spacing: { after: 160 }, children: [new TextRun({ text: '5. 부록. 시나리오 「교신」 v8.9', bold: true, font: FONT, size: 32 })] }),
  ...md(intros['5'], { h1: 'plain' }),
  ...md(scenBody, { h1: 'plain' }),
];

const footer = (label) => new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: `THE NEW MISSION 수정 방향 제안 · 2026.10.06    `, font: FONT, size: 16, color: '888888' }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '888888' })],
  })],
});

const portraitProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1300, left: 1533, right: 1533 } } };
const landscapeProps = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 500, bottom: 700, left: 600, right: 600, footer: 300 } } };

const doc = new Document({
  creator: '시네마직립보행', title: 'THE NEW MISSION 수정 방향 제안',
  styles: {
    default: { document: { run: { font: FONT, size: BODY } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 32, bold: true }, paragraph: { outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 26, bold: true }, paragraph: { outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 23, bold: true }, paragraph: { outlineLevel: 2 } },
    ],
  },
  numbering: { config: [{ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [
    { properties: portraitProps, footers: { default: footer() }, children: part1 },
    { properties: landscapeProps, footers: { default: footer() }, children: landscape },
    { properties: portraitProps, footers: { default: footer() }, children: part3 },
  ],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log('wrote', OUT, (buf.length / 1048576).toFixed(2), 'MB; blocks', part1.length, landscape.length, part3.length);
});

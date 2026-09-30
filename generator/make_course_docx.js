// 統合版 学生配布用 問題（Word）
const fs = require("fs");
const path = require("path");
const sharp = require("sharp");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, SectionType, Header, Footer,
  PageNumber, TabStopType,
} = require("docx");
const { COURSE, expand } = require("./course_seq");
const FIG = require("./figures");

const OUT = process.argv[2] || ".";
const IMG_DIR = path.join(__dirname, "images");
const FONT = { ascii: "Yu Gothic", eastAsia: "Yu Gothic", hAnsi: "Yu Gothic" };
const NAVY = "1F4E79", TEAL = "2A7F9E", LIGHT = "DEEAF6";
const PAGE = { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1134, right: 1134, header: 500, footer: 500 } };
const TEXT_W = 11906 - 1134 * 2; // 本文幅（DXA）

const run = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 20, bold: o.bold, color: o.color });
const border = { style: BorderStyle.SINGLE, size: 4, color: "7F7F7F" };
const borders = { top: border, bottom: border, left: border, right: border };

function chapterHeading(i, ch) {
  return [
    new Paragraph({ spacing: { after: 40 }, shading: { type: ShadingType.CLEAR, fill: ch.kind === "画像" ? TEAL : NAVY, color: "auto" },
      children: [run(` 第${i + 1}章　${ch.kind}`, { bold: true, color: "FFFFFF", size: 20 })] }),
    new Paragraph({ keepNext: true, spacing: { after: 60 }, children: [run(ch.name, { size: 32, bold: true, color: NAVY })] }),
  ];
}

function blockHeading(b) {
  return new Paragraph({
    keepNext: true, spacing: { before: 160, after: 100 },
    border: { top: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 4 }, bottom: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 4 } },
    children: [run(`演習${b.no}　`, { bold: true, color: TEAL, size: 24 }), run(`問${b.first}〜問${b.last}`, { bold: true, size: 24 }),
      run(`（${b.count}問・目安${b.solve}分）`, { size: 18, color: "595959" })],
  });
}

function subHeading(label) {
  return new Paragraph({ keepNext: true, spacing: { before: 160, after: 60 }, shading: { type: ShadingType.CLEAR, fill: LIGHT, color: "auto" },
    children: [run(`■ ${label}`, { bold: true, size: 21, color: NAVY })] });
}

const shortLabel = (b) => (b.subs.length ? b.subs.join("・") : b.chapterName + (b.part ? `（${b.part}）` : ""));

function lectureBanner(l, pageBreakBefore) {
  return new Paragraph({ pageBreakBefore, keepNext: true, spacing: { after: 160 }, shading: { type: ShadingType.CLEAR, fill: NAVY, color: "auto" },
    children: [run(` 第${l.no}回`, { bold: true, color: "FFFFFF", size: 30 }), run(`　問${l.first}〜問${l.last}（${l.last - l.first + 1}問）`, { color: "FFFFFF", size: 22 })] });
}

async function imagePara(src, maxW, maxH) {
  const m = await sharp(src).metadata();
  const buf = await sharp(src).flatten({ background: "#ffffff" }).jpeg({ quality: 85 }).toBuffer();
  const r = Math.min(maxW / m.width, maxH / m.height);
  return new Paragraph({ keepNext: true, alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 },
    children: [new ImageRun({ type: "jpg", data: buf, transformation: { width: Math.round(m.width * r), height: Math.round(m.height * r) } })] });
}

async function questionParas(e) {
  const q = e.q, ps = [];
  ps.push(new Paragraph({
    keepNext: true, keepLines: true, spacing: { before: 180, after: 40 }, indent: { left: 460, hanging: 460 },
    children: [run(`問${e.n} `, { bold: true, color: NAVY }), run(q.q), run(`（${q.src}）`, { size: 16, color: "666666" })],
  }));
  if (e.type === "img") ps.push(await imagePara(fs.readFileSync(path.join(IMG_DIR, q.img)), 360, 250));
  else if (q.fig) ps.push(await imagePara(await FIG[q.fig](false), 190, 170));
  q.choices.forEach((c, i) => ps.push(new Paragraph({
    keepNext: i < q.choices.length - 1, keepLines: true, spacing: { after: 0 }, indent: { left: 760, hanging: 300 },
    children: [run(`${i + 1}．${c}`)],
  })));
  return ps;
}

// 解答欄：演習ブロックごとに 1行7問
function answerSheet(lectures) {
  const cols = 7, cw = Math.floor(TEXT_W / (cols * 2));
  const out = [];
  for (const b of lectures.flatMap((l) => l.blocks)) {
    const l = lectures.find((x) => x.blocks[0] === b);
    if (l) out.push(new Paragraph({ keepNext: true, spacing: { before: 240, after: 40 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 2 } },
      children: [run(`第${l.no}回`, { bold: true, size: 24, color: NAVY })] }));
    out.push(new Paragraph({ keepNext: true, spacing: { before: 140, after: 60 },
      children: [run(`演習${b.no}　問${b.first}〜問${b.last}`, { bold: true, color: TEAL, size: 20 }), run("　　正解数：　　／" + (b.last - b.first + 1), { size: 18 })] }));
    const nums = []; for (let k = b.first; k <= b.last; k++) nums.push(k);
    const rows = [];
    for (let r = 0; r < Math.ceil(nums.length / cols); r++) {
      const cells = [];
      for (let c = 0; c < cols; c++) {
        const k = nums[r * cols + c];
        cells.push(new TableCell({ width: { size: cw, type: WidthType.DXA }, borders, shading: { type: ShadingType.CLEAR, fill: LIGHT, color: "auto" },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run(k ? `問${k}` : "", { bold: true, size: 18 })] })] }));
        cells.push(new TableCell({ width: { size: cw, type: WidthType.DXA }, borders, children: [new Paragraph({ children: [run("")] })] }));
      }
      rows.push(new TableRow({ height: { value: 400, rule: "atLeast" }, cantSplit: true, children: cells }));
    }
    out.push(new Table({ width: { size: cw * cols * 2, type: WidthType.DXA }, columnWidths: Array(cols * 2).fill(cw), rows }));
  }
  return out;
}

function planTable(blocks) {
  const W = [800, 900, 4100, 1800, 1000, 1000];
  const cell = (t, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
    children: [new Paragraph({ alignment: o.center ? AlignmentType.CENTER : AlignmentType.LEFT, children: [run(t, { size: 18, bold: o.bold, color: o.color })] })] });
  const head = new TableRow({ tableHeader: true, children: ["回", "演習", "内容", "問題", "目安", "チェック"].map((t, i) => cell(t, W[i], { fill: NAVY, bold: true, color: "FFFFFF", center: true })) });
  const rows = blocks.map((b, i) => {
    const firstOfLecture = i === 0 || blocks[i - 1].lecture !== b.lecture;
    return new TableRow({ children: [cell(firstOfLecture ? `第${b.lecture}回` : "", W[0], { center: true, bold: true, color: NAVY }),
      cell(`演習${b.no}`, W[1], { center: true, bold: true }), cell(`${b.kind}：${shortLabel(b)}`, W[2]),
      cell(`問${b.first}〜問${b.last}`, W[3], { center: true }), cell(`${b.solve}分`, W[4], { center: true }), cell("□", W[5], { center: true })] });
  });
  return new Table({ width: { size: W.reduce((a, b) => a + b), type: WidthType.DXA }, columnWidths: W, rows: [head, ...rows] });
}

async function build() {
  const { chapters: chs, blocks, lectures, total } = expand();
  const header = new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run(`${COURSE.course}　${COURSE.title}`, { size: 16, color: "808080" })] })] });
  const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "808080" })] })] });
  const hf = { headers: { default: header }, footers: { default: footer } };
  const sections = [];

  // 表紙
  sections.push({ properties: { page: PAGE }, ...hf, children: [
    new Paragraph({ spacing: { before: 600, after: 80 }, children: [run(COURSE.course, { size: 24, color: "595959" })] }),
    new Paragraph({ spacing: { after: 80 }, children: [run(COURSE.title, { size: 52, bold: true, color: NAVY })] }),
    new Paragraph({ spacing: { after: 300 }, children: [run(`${COURSE.subtitle}　全${total}問`, { size: 24 })] }),
    new Paragraph({ spacing: { after: 300 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 4 } },
      tabStops: [{ type: TabStopType.LEFT, position: 4800 }],
      children: [run("学籍番号＿＿＿＿＿＿＿＿", { size: 22 }), run("\t氏名＿＿＿＿＿＿＿＿＿＿＿＿＿＿", { size: 22 })] }),
    new Paragraph({ spacing: { after: 120 }, children: [run("演習の進め方", { size: 24, bold: true, color: NAVY })] }),
    ...["指示された演習の問題を、目安の時間で解く。", "解答は巻末の解答欄に記入する（特に指示のない問題は1つ選ぶ）。", "解説を聞きながら丸つけをし、間違えた問題はメモ欄に理由を書く。"]
      .map((t, i) => new Paragraph({ spacing: { after: 60 }, indent: { left: 400, hanging: 300 }, children: [run(`${i + 1}．${t}`)] })),
    new Paragraph({ spacing: { before: 240, after: 120 }, children: [run("演習一覧", { size: 24, bold: true, color: NAVY })] }),
    planTable(blocks),
  ] });

  // 各章：基礎は2段組、画像は1段組
  for (let i = 0; i < chs.length; i++) {
    const ch = chs[i];
    const children = [...(ch.lectureBefore ? [lectureBanner(ch.lectureBefore, false)] : []), ...chapterHeading(i, ch)];
    for (const e of ch.seq) {
      if (e.type === "lecture") children.push(lectureBanner(e.lecture, true));
      else if (e.type === "solve") children.push(blockHeading(e.block));
      else if (e.type === "sub") children.push(subHeading(e.label));
      else if (e.type === "base" || e.type === "img") children.push(...(await questionParas(e)));
    }
    const cols = ch.kind === "基礎" ? { column: { count: 2, space: 500 } } : {};
    sections.push({ properties: { page: PAGE, type: SectionType.NEXT_PAGE, ...cols }, ...hf, children });
  }

  // 解答欄・メモ
  sections.push({ properties: { page: PAGE, type: SectionType.NEXT_PAGE }, ...hf, children: [
    new Paragraph({ spacing: { after: 60 }, children: [run("解答欄", { size: 32, bold: true, color: NAVY })] }),
    ...answerSheet(lectures),
    new Paragraph({ spacing: { before: 300 }, children: [run(`合計正解数：　　　／ ${total}問`, { size: 24, bold: true })] }),
  ] });
  sections.push({ properties: { page: PAGE, type: SectionType.NEXT_PAGE }, ...hf, children: [
    new Paragraph({ spacing: { after: 60 }, children: [run("間違えた問題・覚えておくことメモ", { size: 28, bold: true, color: NAVY })] }),
    ...Array.from({ length: 26 }, () => new Paragraph({ spacing: { before: 300 }, border: { bottom: { style: BorderStyle.DOTTED, size: 4, color: "A6A6A6", space: 1 } }, children: [run("")] })),
  ] });

  const doc = new Document({ styles: { default: { document: { run: { font: FONT, size: 20 } } } }, sections });
  const file = path.join(OUT, `${COURSE.title}_問題（統合版）.docx`);
  fs.writeFileSync(file, await Packer.toBuffer(doc));
  console.log("wrote", file);
}

build();

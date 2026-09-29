// 学生配布用 問題（Word）
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, SectionType, Header, Footer,
  PageNumber, TabStopType,
} = require("docx");
const DATA = require("./data");
const FIG = require("./figures");

const OUT = process.argv[2] || ".";
const FONT = { ascii: "Yu Gothic", eastAsia: "Yu Gothic", hAnsi: "Yu Gothic" };
const C1 = "①②③④⑤";

function run(text, o = {}) {
  return new TextRun({ text, font: FONT, size: o.size || 20, bold: o.bold, color: o.color });
}

async function questionParas(q, n) {
  const ps = [];
  ps.push(new Paragraph({
    keepNext: true, keepLines: true, spacing: { before: 160, after: 40 },
    indent: { left: 400, hanging: 400 },
    children: [run(`[${n}] `, { bold: true, color: "1F4E79" }), run(q.q), run(`（${q.src}）`, { size: 16, color: "666666" })],
  }));
  if (q.fig) {
    const [w, h] = FIG.sizes[q.fig];
    const scale = 200 / w;
    ps.push(new Paragraph({
      keepNext: true, alignment: AlignmentType.CENTER,
      children: [new ImageRun({ type: "png", data: await FIG[q.fig](false), transformation: { width: w * scale, height: h * scale } })],
    }));
  }
  q.choices.forEach((c, i) => {
    ps.push(new Paragraph({
      keepNext: i < q.choices.length - 1, keepLines: true, spacing: { after: 0 },
      indent: { left: 700, hanging: 300 },
      children: [run(`${i + 1}．${c}`)],
    }));
  });
  return ps;
}

function sectionHeading(text) {
  return new Paragraph({
    keepNext: true, spacing: { before: 240, after: 80 },
    shading: { type: ShadingType.CLEAR, fill: "DEEAF6", color: "auto" },
    children: [run(`■ ${text}`, { bold: true, size: 22, color: "1F4E79" })],
  });
}

function answerTable(total) {
  const cols = 5, rows = Math.ceil(total / cols);
  const W = 9638, cw = Math.floor(W / (cols * 2));
  const border = { style: BorderStyle.SINGLE, size: 4, color: "7F7F7F" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const trs = [];
  for (let r = 0; r < rows; r++) {
    const cells = [];
    for (let c = 0; c < cols; c++) {
      const n = c * rows + r + 1;
      cells.push(new TableCell({
        width: { size: cw, type: WidthType.DXA }, borders,
        shading: { type: ShadingType.CLEAR, fill: "DEEAF6", color: "auto" },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run(n <= total ? `[${n}]` : "", { bold: true })] })],
      }));
      cells.push(new TableCell({ width: { size: cw, type: WidthType.DXA }, borders, children: [new Paragraph({ children: [run("")] })] }));
    }
    trs.push(new TableRow({ height: { value: 420, rule: "atLeast" }, children: cells }));
  }
  return new Table({ width: { size: cw * cols * 2, type: WidthType.DXA }, columnWidths: Array(cols * 2).fill(cw), rows: trs });
}

async function build(theme) {
  const total = theme.sections.reduce((a, s) => a + s.qs.length, 0);
  const header = new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run("作業療法士国家試験対策　基礎医学（神経系）", { size: 16, color: "808080" })] })] });
  const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "808080" })] })] });
  const page = { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1134, right: 1134, header: 500, footer: 500 } };

  // 表紙部（1段組）
  const top = [
    new Paragraph({ spacing: { after: 60 }, children: [run(`テーマ ${theme.id}`, { size: 20, bold: true, color: "FFFFFF" })], shading: { type: ShadingType.CLEAR, fill: "1F4E79", color: "auto" } }),
    new Paragraph({ spacing: { after: 120 }, children: [run(theme.title, { size: 36, bold: true, color: "1F4E79" })] }),
    new Paragraph({
      spacing: { after: 120 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 4 } },
      tabStops: [{ type: TabStopType.LEFT, position: 4200 }],
      children: [run(`全${total}問`, { size: 20 }), run("\t学籍番号＿＿＿＿＿＿　氏名＿＿＿＿＿＿＿＿＿＿", { size: 20 })],
    }),
    new Paragraph({ spacing: { after: 80 }, children: [run("※ 特に指示のない問題は、正しいもの（または問われているもの）を１つ選ぶ。解答は最終ページの解答欄に記入すること。", { size: 17, color: "404040" })] }),
  ];

  const body = [];
  let n = 0;
  for (const s of theme.sections) {
    if (s.name) body.push(sectionHeading(s.name));
    for (const q of s.qs) body.push(...(await questionParas(q, ++n)));
  }

  const ans = [
    new Paragraph({ spacing: { after: 120 }, children: [run("解答欄", { size: 28, bold: true, color: "1F4E79" })] }),
    answerTable(total),
    new Paragraph({ spacing: { before: 240 }, children: [run("正解数：　　　／ " + total + " 問", { size: 22, bold: true })] }),
    new Paragraph({ spacing: { before: 200 }, children: [run("■ 間違えた問題・覚えておくことメモ", { size: 20, bold: true, color: "1F4E79" })] }),
    ...Array.from({ length: 10 }, () => new Paragraph({ spacing: { before: 280 }, border: { bottom: { style: BorderStyle.DOTTED, size: 4, color: "A6A6A6", space: 1 } }, children: [run("")] })),
  ];

  const doc = new Document({
    styles: { default: { document: { run: { font: FONT, size: 20 } } } },
    sections: [
      { properties: { page }, headers: { default: header }, footers: { default: footer }, children: top },
      { properties: { page, type: SectionType.CONTINUOUS, column: { count: 2, space: 500 } }, headers: { default: header }, footers: { default: footer }, children: body },
      { properties: { page, type: SectionType.NEXT_PAGE }, headers: { default: header }, footers: { default: footer }, children: ans },
    ],
  });
  const file = path.join(OUT, `${theme.id}_${theme.title.replace(/・/g, "・")}_問題.docx`);
  fs.writeFileSync(file, await Packer.toBuffer(doc));
  console.log("wrote", file);
}

(async () => { for (const t of DATA) await build(t); })();

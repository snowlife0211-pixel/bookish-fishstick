// 解説用 PPT
const path = require("path");
const pptxgen = require("pptxgenjs");
const DATA = require("./data");
const FIG = require("./figures");
const SUMMARY = require("./summary");

const OUT = process.argv[2] || ".";
const F = "Meiryo";
const NAVY = "17324D", TEAL = "2A7F9E", ICE = "EAF2F8", ACC = "D9534F", ACC_BG = "FBE3E2", INK = "222222", MUTED = "5F6B76";


// 文字数から収まるフォントサイズを計算（全角=1em、半角=0.55em、行高1.25）
function fitSize(text, w, h, max, min = 9) {
  const str = Array.isArray(text) ? text.map((t) => t.text).join("") : text;
  const em = [...str].reduce((a, ch) => a + (/[\x20-\x7E]/.test(ch) ? 0.55 : 1), 0);
  for (let pt = max; pt > min; pt -= 0.5) {
    const perLine = Math.floor((w - 0.05) / (pt / 72));
    const lines = Math.ceil(em / Math.max(perLine, 1));
    if (lines * pt * 1.25 / 72 <= h - 0.04) return pt;
  }
  return min;
}

function title(slide, text, sub) {
  slide.addText(text, { x: 0.5, y: 0.25, w: 9, h: 0.6, fontFace: F, fontSize: 24, bold: true, color: NAVY, margin: 0, isTextBox: true });
  if (sub) slide.addText(sub, { x: 0.5, y: 0.82, w: 9, h: 0.3, fontFace: F, fontSize: 11, color: MUTED, margin: 0, isTextBox: true });
}

function coverSlide(pres, theme, total) {
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 6.6, y: -1.2, w: 5, h: 5, fill: { color: TEAL, transparency: 55 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: 7.8, y: 3.0, w: 3.2, h: 3.2, fill: { color: TEAL, transparency: 75 }, line: { type: "none" } });
  s.addText("作業療法士国家試験対策　基礎医学（神経系）", { x: 0.6, y: 1.2, w: 8, h: 0.4, fontFace: F, fontSize: 14, color: "BFD7EA", margin: 0, isTextBox: true });
  s.addText(`テーマ ${theme.id}`, { x: 0.6, y: 1.8, w: 8, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
  s.addText(theme.title, { x: 0.6, y: 2.3, w: 8.5, h: 1.0, fontFace: F, fontSize: 40, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
  s.addText(`解説スライド　全${total}問`, { x: 0.6, y: 3.5, w: 8, h: 0.4, fontFace: F, fontSize: 16, color: "BFD7EA", margin: 0, isTextBox: true });
}

function dividerSlide(pres, label, sub) {
  const s = pres.addSlide();
  s.background = { color: TEAL };
  s.addText(label, { x: 0.6, y: 2.0, w: 8.8, h: 0.9, fontFace: F, fontSize: 40, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
  if (sub) s.addText(sub, { x: 0.6, y: 2.9, w: 8.8, h: 0.5, fontFace: F, fontSize: 16, color: "E6F2F7", margin: 0, isTextBox: true });
}

// まとめスライド：表形式
function tableSlide(pres, sm) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  title(s, sm.title, sm.sub);
  const hdr = sm.header.map((h) => ({ text: h, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, align: "center" } }));
  const rows = sm.rows.map((r, i) => r.map((c, j) => ({
    text: c, options: { fill: { color: i % 2 ? "FFFFFF" : ICE }, bold: j === 0, color: j === 0 ? NAVY : INK },
  })));
  s.addTable([hdr, ...rows], {
    x: 0.5, y: 1.2, w: 9, colW: sm.colW, fontFace: F, fontSize: sm.fontSize || 11,
    border: { type: "solid", pt: 0.75, color: "C9D6E0" }, valign: "middle", rowH: sm.rowH || 0.3, margin: 0.04,
  });
  if (sm.note) s.addText(sm.note, { x: 0.5, y: 5.05, w: 9, h: 0.35, fontFace: F, fontSize: 11, color: ACC, bold: true, margin: 0, isTextBox: true });
}

// まとめスライド：カード形式
function cardSlide(pres, sm) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  title(s, sm.title, sm.sub);
  const n = sm.cards.length, gap = 0.3, w = (9 - gap * (n - 1)) / n;
  sm.cards.forEach((c, i) => {
    const x = 0.5 + i * (w + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.25, w, h: 3.75, rectRadius: 0.1, fill: { color: ICE }, line: { color: "C9D6E0", width: 0.75 } });
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: 1.4, w: 0.5, h: 0.5, fill: { color: TEAL }, line: { type: "none" } });
    s.addText(String(i + 1), { x: x + 0.2, y: 1.4, w: 0.5, h: 0.5, align: "center", valign: "middle", fontFace: "Arial", fontSize: 16, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
    s.addText(c.head, { x: x + 0.8, y: 1.4, w: w - 0.95, h: 0.5, valign: "middle", fontFace: F, fontSize: 15, bold: true, color: NAVY, margin: 0, isTextBox: true });
    s.addText(c.items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < c.items.length - 1 } })), {
      x: x + 0.15, y: 2.05, w: w - 0.3, h: 2.85, valign: "top", fontFace: F, fontSize: sm.fontSize || 12, color: INK, paraSpaceAfter: 5, isTextBox: true,
    });
  });
  if (sm.note) s.addText(sm.note, { x: 0.5, y: 5.1, w: 9, h: 0.35, fontFace: F, fontSize: 11, color: ACC, bold: true, margin: 0, isTextBox: true });
}

function qHeader(s, pres, q, n, tag) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.3, w: 0.95, h: 0.42, rectRadius: 0.08, fill: { color: tag === "解説" ? ACC : NAVY }, line: { type: "none" } });
  s.addText(`${tag}`, { x: 0.5, y: 0.3, w: 0.95, h: 0.42, align: "center", valign: "middle", fontFace: F, fontSize: 13, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
  s.addText(`[${n}]`, { x: 1.55, y: 0.3, w: 0.8, h: 0.42, valign: "middle", fontFace: "Arial", fontSize: 20, bold: true, color: NAVY, margin: 0, isTextBox: true });
  s.addText(q.src, { x: 6.5, y: 0.3, w: 3.0, h: 0.42, align: "right", valign: "middle", fontFace: F, fontSize: 11, color: MUTED, margin: 0, isTextBox: true });
  s.addText(q.q, { x: 0.5, y: 0.8, w: 9, h: 0.65, valign: "top", fontFace: F, fontSize: fitSize(q.q, 9, 0.65, 17), bold: true, color: INK, margin: 0, isTextBox: true });
}

async function questionSlide(pres, q, n) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  qHeader(s, pres, q, n, "問題");
  const w = q.fig ? 5.2 : 9;
  q.choices.forEach((c, i) => {
    const y = 1.65 + i * 0.66;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w, h: 0.54, rectRadius: 0.08, fill: { color: ICE }, line: { type: "none" } });
    s.addShape(pres.shapes.OVAL, { x: 0.62, y: y + 0.07, w: 0.4, h: 0.4, fill: { color: "FFFFFF" }, line: { color: NAVY, width: 1.25 } });
    s.addText(String(i + 1), { x: 0.62, y: y + 0.07, w: 0.4, h: 0.4, align: "center", valign: "middle", fontFace: "Arial", fontSize: 14, bold: true, color: NAVY, margin: 0, isTextBox: true });
    s.addText(c, { x: 1.15, y, w: w - 0.75, h: 0.54, valign: "middle", fontFace: F, fontSize: fitSize(c, w - 0.75, 0.54, 15), color: INK, margin: 0, isTextBox: true });
  });
  if (q.fig) {
    const [fw, fh] = FIG.sizes[q.fig];
    const h = 3.2, ww = h * fw / fh;
    s.addImage({ data: "image/png;base64," + (await FIG[q.fig](false)).toString("base64"), x: 5.95 + (3.55 - ww) / 2, y: 1.65, w: ww, h });
  }
}

async function answerSlide(pres, q, n) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  qHeader(s, pres, q, n, "解説");
  const hasFig = !!q.fig;
  const w = hasFig ? 5.9 : 9;
  const rowH = 0.5, gap = 0.08, y0 = 1.55;
  const cw = hasFig ? 2.0 : 3.3;
  q.choices.forEach((c, i) => {
    const ok = q.ans.includes(i + 1);
    const y = y0 + i * (rowH + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w, h: rowH, rectRadius: 0.08, fill: { color: ok ? ACC_BG : "F4F6F8" }, line: ok ? { color: ACC, width: 1.5 } : { type: "none" } });
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.07, w: 0.36, h: 0.36, fill: { color: ok ? ACC : "FFFFFF" }, line: { color: ok ? ACC : "8A96A3", width: 1 } });
    s.addText(String(i + 1), { x: 0.6, y: y + 0.07, w: 0.36, h: 0.36, align: "center", valign: "middle", fontFace: "Arial", fontSize: 13, bold: true, color: ok ? "FFFFFF" : "56616C", margin: 0, isTextBox: true });
    s.addText(c, { x: 1.08, y, w: cw, h: rowH, valign: "middle", fontFace: F, fontSize: fitSize(c, cw, rowH, 12), bold: true, color: ok ? ACC : INK, margin: 0, isTextBox: true });
    s.addText(q.notes[i], { x: 1.08 + cw + 0.1, y, w: w - (cw + 0.75), h: rowH, valign: "middle", fontFace: F, fontSize: fitSize(q.notes[i], w - (cw + 0.75), rowH, 11), color: ok ? INK : MUTED, margin: 0, isTextBox: true });
  });
  if (hasFig) {
    const [fw, fh] = FIG.sizes[q.fig];
    const h = 2.8, ww = h * fw / fh;
    s.addImage({ data: "image/png;base64," + (await FIG[q.fig](true)).toString("base64"), x: 6.55 + (2.95 - ww) / 2, y: 1.55, w: ww, h });
  }
  // 正解と要点
  const yb = 4.55;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: yb, w: 1.6, h: 0.75, rectRadius: 0.08, fill: { color: ACC }, line: { type: "none" } });
  s.addText([{ text: "正解", options: { fontSize: 11, breakLine: true } }, { text: q.ans.join("・"), options: { fontSize: 22, bold: true } }],
    { x: 0.5, y: yb, w: 1.6, h: 0.75, align: "center", valign: "middle", fontFace: F, color: "FFFFFF", margin: 0, isTextBox: true });
  if (q.point) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.25, y: yb, w: 7.25, h: 0.75, rectRadius: 0.08, fill: { color: ICE }, line: { type: "none" } });
    s.addText([{ text: "POINT　", options: { bold: true, color: TEAL } }, { text: q.point, options: { color: INK } }],
      { x: 2.4, y: yb, w: 7.0, h: 0.75, valign: "middle", fontFace: F, fontSize: fitSize("POINT　" + q.point, 7.0, 0.75, 12), margin: 0, isTextBox: true });
  }
}

async function build(theme) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = `${theme.id} ${theme.title} 解説`;
  const total = theme.sections.reduce((a, s) => a + s.qs.length, 0);
  coverSlide(pres, theme, total);
  let n = 0;
  for (const sec of theme.sections) {
    const sms = SUMMARY[theme.id][sec.name || "_"] || [];
    const first = n + 1, last = n + sec.qs.length;
    if (sec.name) dividerSlide(pres, sec.name, `問題 [${first}]〜[${last}]`);
    for (const sm of sms) (sm.cards ? cardSlide : tableSlide)(pres, sm);
    for (const q of sec.qs) { n++; await questionSlide(pres, q, n); await answerSlide(pres, q, n); }
  }
  // 最後：正答一覧
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  title(s, "正答一覧", `テーマ ${theme.id}　${theme.title}`);
  const all = theme.sections.flatMap((x) => x.qs);
  const cols = 5, rows = Math.ceil(all.length / cols);
  const tbl = [];
  for (let r = 0; r < rows; r++) {
    const row = [];
    for (let c = 0; c < cols; c++) {
      const k = c * rows + r;
      row.push({ text: k < all.length ? `[${k + 1}]` : "", options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, align: "center" } });
      row.push({ text: k < all.length ? all[k].ans.join("・") : "", options: { bold: true, color: ACC, fill: { color: r % 2 ? "FFFFFF" : ICE }, align: "center" } });
    }
    tbl.push(row);
  }
  s.addTable(tbl, { x: 0.5, y: 1.3, w: 9, colW: Array(10).fill(0.9), rowH: Math.min(0.5, 3.9 / rows), fontFace: F, fontSize: 13, valign: "middle", border: { type: "solid", pt: 0.75, color: "C9D6E0" } });

  const file = path.join(OUT, `${theme.id}_${theme.title}_解説.pptx`);
  await pres.writeFile({ fileName: file });
  console.log("wrote", file);
}

(async () => { for (const t of DATA) await build(t); })();

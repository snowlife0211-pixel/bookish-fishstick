// 統合版 解説PPT（基礎＋画像問題を1つの流れに）
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const { COURSE, expand } = require("./course_seq");
const L = require("./slides_lib");
const { F, NAVY, TEAL, ICE, ACC, ACC_BG, INK, MUTED, fitSize } = L;

const OUT = process.argv[2] || ".";
const IMG_DIR = path.join(__dirname, "images");
const TEACHER_NOTES = JSON.parse(fs.readFileSync(path.join(__dirname, "img_notes.json"), "utf8"));

async function imgInfo(file) {
  const p = path.join(IMG_DIR, file);
  const m = await sharp(p).metadata();
  const jpg = await sharp(p).flatten({ background: "#ffffff" }).jpeg({ quality: 85 }).toBuffer();
  return { data: "image/jpeg;base64," + jpg.toString("base64"), w: m.width, h: m.height };
}

function contain(info, x, y, w, h) {
  const r = Math.min(w / info.w, h / info.h);
  const ww = info.w * r, hh = info.h * r;
  return { data: info.data, x: x + (w - ww) / 2, y: y + (h - hh) / 2, w: ww, h: hh };
}

function header(s, pres, q, n, tag) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.25, w: 0.95, h: 0.42, rectRadius: 0.08, fill: { color: tag === "解説" ? ACC : NAVY }, line: { type: "none" } });
  s.addText(tag, { x: 0.5, y: 0.25, w: 0.95, h: 0.42, align: "center", valign: "middle", fontFace: F, fontSize: 13, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
  s.addText(`[${n}]`, { x: 1.55, y: 0.25, w: 0.8, h: 0.42, valign: "middle", fontFace: "Arial", fontSize: 20, bold: true, color: NAVY, margin: 0, isTextBox: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.35, y: 0.31, w: 0.75, h: 0.3, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: TEAL, width: 1 } });
  s.addText("画像", { x: 2.35, y: 0.31, w: 0.75, h: 0.3, align: "center", valign: "middle", fontFace: F, fontSize: 10, bold: true, color: TEAL, margin: 0, isTextBox: true });
  s.addText(q.src, { x: 6.5, y: 0.25, w: 3.0, h: 0.42, align: "right", valign: "middle", fontFace: F, fontSize: 11, color: MUTED, margin: 0, isTextBox: true });
}

function numCircle(s, pres, x, y, d, n, ok) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: ok ? ACC : "FFFFFF" }, line: { color: ok ? ACC : NAVY, width: 1 } });
  s.addText(String(n), { x, y, w: d, h: d, align: "center", valign: "middle", fontFace: "Arial", fontSize: 12, bold: true, color: ok ? "FFFFFF" : NAVY, margin: 0, isTextBox: true });
}

async function imgQuestionSlide(pres, q, n) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  header(s, pres, q, n, "問題");
  const info = await imgInfo(q.img);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.85, w: 4.3, h: 4.5, fill: { color: "F4F6F8" }, line: { type: "none" } });
  s.addImage(contain(info, 0.55, 0.9, 4.2, 4.4));
  const rowH = 0.46, gap = 0.06, x = 5.05, w = 4.45;
  const choicesTop = 5.35 - q.choices.length * (rowH + gap) + gap;
  const stemH = choicesTop - 0.85 - 0.12;
  s.addText(q.q, { x, y: 0.85, w, h: stemH, valign: "top", fontFace: F, fontSize: fitSize(q.q, w, stemH, 14), bold: true, color: INK, margin: 0, isTextBox: true });
  q.choices.forEach((c, i) => {
    const y = choicesTop + i * (rowH + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: rowH, rectRadius: 0.06, fill: { color: ICE }, line: { type: "none" } });
    numCircle(s, pres, x + 0.08, y + 0.07, 0.32, i + 1, false);
    s.addText(c, { x: x + 0.5, y, w: w - 0.58, h: rowH, valign: "middle", fontFace: F, fontSize: fitSize(c, w - 0.58, rowH, 13), color: INK, margin: 0, isTextBox: true });
  });
}

async function imgAnswerSlide(pres, q, n, id) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  header(s, pres, q, n, "解説");
  const info = await imgInfo(q.img);
  // 左：画像と画像所見
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.85, w: 3.0, h: 2.35, fill: { color: "F4F6F8" }, line: { type: "none" } });
  s.addImage(contain(info, 0.55, 0.9, 2.9, 2.25));
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 3.3, w: 3.0, h: 1.1, rectRadius: 0.08, fill: { color: ICE }, line: { type: "none" } });
  s.addText("画像所見・着眼点", { x: 0.62, y: 3.35, w: 2.8, h: 0.26, fontFace: F, fontSize: 10, bold: true, color: TEAL, margin: 0, isTextBox: true });
  s.addText(q.find, { x: 0.62, y: 3.62, w: 2.8, h: 0.74, valign: "top", fontFace: F, fontSize: fitSize(q.find, 2.8, 0.74, 11), color: INK, margin: 0, isTextBox: true });
  // 右：選択肢ごとの解説
  const x = 3.7, w = 5.8, rowH = 0.64, gap = 0.07, cw = 1.9;
  q.choices.forEach((c, i) => {
    const ok = q.ans.includes(i + 1);
    const y = 0.85 + i * (rowH + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: rowH, rectRadius: 0.06, fill: { color: ok ? ACC_BG : "F4F6F8" }, line: ok ? { color: ACC, width: 1.5 } : { type: "none" } });
    numCircle(s, pres, x + 0.08, y + (rowH - 0.32) / 2, 0.32, i + 1, ok);
    s.addText(c, { x: x + 0.5, y, w: cw, h: rowH, valign: "middle", fontFace: F, fontSize: fitSize(c, cw, rowH, 12), bold: true, color: ok ? ACC : INK, margin: 0, isTextBox: true });
    const note = q.notes[i];
    if (note) s.addText(note, { x: x + 0.5 + cw + 0.1, y, w: w - (cw + 0.68), h: rowH, valign: "middle", fontFace: F, fontSize: fitSize(note, w - (cw + 0.68), rowH, 11), color: ok ? INK : MUTED, margin: 0, isTextBox: true });
  });
  // 下：正解と POINT
  const yb = 4.55;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: yb, w: 1.6, h: 0.75, rectRadius: 0.08, fill: { color: ACC }, line: { type: "none" } });
  s.addText([{ text: "正解", options: { fontSize: 11, breakLine: true } }, { text: q.ans.join("・"), options: { fontSize: 22, bold: true } }],
    { x: 0.5, y: yb, w: 1.6, h: 0.75, align: "center", valign: "middle", fontFace: F, color: "FFFFFF", margin: 0, isTextBox: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.25, y: yb, w: 7.25, h: 0.75, rectRadius: 0.08, fill: { color: ICE }, line: { type: "none" } });
  s.addText([{ text: "POINT　", options: { bold: true, color: TEAL } }, { text: q.point, options: { color: INK } }],
    { x: 2.4, y: yb, w: 7.0, h: 0.75, valign: "middle", fontFace: F, fontSize: fitSize("POINT　" + q.point, 7.0, 0.75, 12), margin: 0, isTextBox: true });
  // 発表者ノート：元資料の解説メモ
  let notes = "";
  if (TEACHER_NOTES[id]) notes += "【元資料（画像問題編）の解説メモ】\n" + TEACHER_NOTES[id];
  if (q.fix) notes += (notes ? "\n\n" : "") + "【補足】" + q.fix;
  if (notes) s.addNotes(notes);
}

function coverSlide(pres, total, counts) {
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 6.6, y: -1.2, w: 5, h: 5, fill: { color: TEAL, transparency: 55 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: 7.8, y: 3.0, w: 3.2, h: 3.2, fill: { color: TEAL, transparency: 75 }, line: { type: "none" } });
  s.addText(COURSE.course, { x: 0.6, y: 1.1, w: 8, h: 0.4, fontFace: F, fontSize: 14, color: "BFD7EA", margin: 0, isTextBox: true });
  s.addText(COURSE.title, { x: 0.6, y: 1.6, w: 8.5, h: 1.0, fontFace: F, fontSize: 40, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
  s.addText(COURSE.subtitle, { x: 0.6, y: 2.6, w: 8.5, h: 0.5, fontFace: F, fontSize: 18, color: "FFFFFF", margin: 0, isTextBox: true });
  s.addText(`全${total}問（基礎 ${counts["基礎"]}問・画像 ${counts["画像"]}問）`, { x: 0.6, y: 3.4, w: 8, h: 0.4, fontFace: F, fontSize: 16, color: "BFD7EA", margin: 0, isTextBox: true });
}

// ブロックの短い名前（小見出しがあればそれを使う）
const shortLabel = (b) => (b.subs.length ? b.subs.join("・") : b.chapterName + (b.part ? `（${b.part}）` : ""));

function roadmapSlide(pres, lectures) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  L.title(s, "この講義の流れ（全3回）", "各回とも「解く → 解説」を3回くり返す。基礎で覚えた知識を、すぐ次の画像問題で使う");
  const w = 2.8, h = 3.85, gx = 0.3, y = 1.3;
  lectures.forEach((l, i) => {
    const x = 0.5 + i * (w + gx);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: { color: ICE }, line: { color: "C9D6E0", width: 0.75 } });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.85, rectRadius: 0.1, fill: { color: NAVY }, line: { type: "none" } });
    s.addText(`第${l.no}回`, { x: x + 0.2, y: y + 0.08, w: w - 0.4, h: 0.42, fontFace: F, fontSize: 20, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
    s.addText(`問${l.first}〜問${l.last}・約${l.total}分`, { x: x + 0.2, y: y + 0.48, w: w - 0.4, h: 0.3, fontFace: F, fontSize: 11, color: "BFD7EA", margin: 0, isTextBox: true });
    l.blocks.forEach((b, k) => {
      const yy = y + 1.0 + k * 0.92;
      const img = b.kind === "画像";
      s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: yy + 0.05, w: 0.42, h: 0.42, fill: { color: img ? TEAL : NAVY }, line: { type: "none" } });
      s.addText(String(b.no), { x: x + 0.15, y: yy + 0.05, w: 0.42, h: 0.42, align: "center", valign: "middle", fontFace: "Arial", fontSize: 13, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
      s.addText(`${b.kind}｜問${b.first}〜${b.last}`, { x: x + 0.67, y: yy, w: w - 0.8, h: 0.26, fontFace: F, fontSize: 10, bold: true, color: img ? TEAL : NAVY, margin: 0, isTextBox: true });
      const t = shortLabel(b);
      s.addText(t, { x: x + 0.67, y: yy + 0.26, w: w - 0.8, h: 0.56, valign: "top", fontFace: F, fontSize: fitSize(t, w - 0.8, 0.56, 12, 9), color: INK, margin: 0, isTextBox: true });
    });
  });
}

// 各回の最初：その回の流れ
function lectureSlide(pres, l) {
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 7.2, y: -1.0, w: 4, h: 4, fill: { color: TEAL, transparency: 60 }, line: { type: "none" } });
  s.addText(COURSE.course, { x: 0.6, y: 0.45, w: 6, h: 0.35, fontFace: F, fontSize: 13, color: "BFD7EA", margin: 0, isTextBox: true });
  s.addText(`第${l.no}回`, { x: 0.6, y: 0.8, w: 6, h: 0.75, fontFace: F, fontSize: 40, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
  s.addText(`問${l.first}〜問${l.last}（${l.last - l.first + 1}問）・目安 約${l.total}分`, { x: 0.6, y: 1.55, w: 8, h: 0.4, fontFace: F, fontSize: 15, color: "FFFFFF", margin: 0, isTextBox: true });
  const hdr = ["演習", "内容", "問題", "解く", "解説"].map((t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: TEAL }, align: "center" } }));
  const rows = l.blocks.map((b) => [
    { text: `演習${b.no}`, options: { bold: true, align: "center" } },
    { text: `${b.kind}：${shortLabel(b)}` },
    { text: `問${b.first}〜${b.last}`, options: { align: "center" } },
    { text: `${b.solve}分`, options: { align: "center", bold: true } },
    { text: `約${b.explain}分`, options: { align: "center" } },
  ].map((c) => ({ text: c.text, options: { color: INK, fill: { color: "FFFFFF" }, ...(c.options || {}) } })));
  s.addTable([hdr, ...rows], { x: 0.6, y: 2.25, w: 8.8, colW: [1.0, 4.4, 1.4, 1.0, 1.0], rowH: 0.55, fontFace: F, fontSize: 13, valign: "middle", border: { type: "solid", pt: 0.75, color: "C9D6E0" } });
}

function chapterDivider(pres, i, ch, first, last) {
  const s = pres.addSlide();
  s.background = { color: ch.kind === "画像" ? TEAL : NAVY };
  s.addText(`第${i + 1}章　${ch.kind}`, { x: 0.6, y: 1.5, w: 8.8, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: "E6F2F7", margin: 0, isTextBox: true });
  s.addText(ch.name, { x: 0.6, y: 2.0, w: 8.8, h: 1.0, fontFace: F, fontSize: fitSize(ch.name, 8.8, 1.0, 36), bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
  s.addText(ch.lead, { x: 0.6, y: 3.1, w: 8.8, h: 0.45, fontFace: F, fontSize: 16, color: "FFFFFF", margin: 0, isTextBox: true });
  s.addText(`問題 [${first}]〜[${last}]`, { x: 0.6, y: 3.65, w: 8.8, h: 0.4, fontFace: F, fontSize: 14, color: "E6F2F7", margin: 0, isTextBox: true });
}

function subDivider(pres, label, chName) {
  const s = pres.addSlide();
  s.background = { color: ICE };
  s.addText(chName, { x: 0.6, y: 2.0, w: 8.8, h: 0.4, fontFace: F, fontSize: 14, color: TEAL, bold: true, margin: 0, isTextBox: true });
  s.addText(label, { x: 0.6, y: 2.4, w: 8.8, h: 0.8, fontFace: F, fontSize: 30, bold: true, color: NAVY, margin: 0, isTextBox: true });
}

function solveSlide(pres, b, chName) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.5, w: 9, h: 4.6, rectRadius: 0.15, fill: { color: ICE }, line: { type: "none" } });
  s.addText(`第${b.lecture}回　演習 ${b.no}`, { x: 0.9, y: 0.85, w: 6, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: TEAL, margin: 0, isTextBox: true });
  s.addText(chName, { x: 0.9, y: 1.35, w: 8.2, h: 0.4, fontFace: F, fontSize: 14, color: MUTED, margin: 0, isTextBox: true });
  s.addText(`問題 [${b.first}] 〜 [${b.last}] を解きましょう`, { x: 0.9, y: 1.95, w: 8.2, h: 0.9, fontFace: F, fontSize: 32, bold: true, color: NAVY, margin: 0, isTextBox: true });
  const kinds = [b.base ? `基礎 ${b.base}問` : "", b.img ? `画像 ${b.img}問` : ""].filter(Boolean).join("・");
  [["解く", `${b.solve}分`, NAVY, 0.9], ["解説", `約${b.explain}分`, TEAL, 2.75]].forEach(([k, v, c, x]) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 3.2, w: 1.7, h: 1.3, rectRadius: 0.1, fill: { color: c }, line: { type: "none" } });
    s.addText([{ text: k, options: { fontSize: 13, breakLine: true } }, { text: v, options: { fontSize: 26, bold: true } }],
      { x, y: 3.2, w: 1.7, h: 1.3, align: "center", valign: "middle", fontFace: F, color: "FFFFFF", margin: 0, isTextBox: true });
  });
  s.addText([{ text: `${b.count}問（${kinds}）`, options: { bold: true, breakLine: true } }, { text: "配布プリントに解答を記入。終わったら解説に進みます。" }],
    { x: 4.7, y: 3.2, w: 4.4, h: 1.3, valign: "middle", fontFace: F, fontSize: 14, color: INK, margin: 0, isTextBox: true });
}

function answerListSlides(pres, list) {
  const per = 45;
  for (let p = 0; p * per < list.length; p++) {
    const part = list.slice(p * per, (p + 1) * per);
    const s = pres.addSlide();
    s.background = { color: "FFFFFF" };
    L.title(s, `正答一覧（${p + 1}/${Math.ceil(list.length / per)}）`, "★＝画像問題");
    const cols = 5, rows = Math.ceil(part.length / cols), tbl = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const e = part[c * rows + r];
        row.push({ text: e ? `[${e.n}]${e.img ? "★" : ""}` : "", options: { bold: true, color: "FFFFFF", fill: { color: e && e.img ? TEAL : NAVY }, align: "center" } });
        row.push({ text: e ? e.ans : "", options: { bold: true, color: ACC, fill: { color: r % 2 ? "FFFFFF" : ICE }, align: "center" } });
      }
      tbl.push(row);
    }
    s.addTable(tbl, { x: 0.5, y: 1.25, w: 9, colW: Array(10).fill(0.9), rowH: Math.min(0.42, 3.95 / rows), fontFace: F, fontSize: 12, valign: "middle", border: { type: "solid", pt: 0.75, color: "C9D6E0" } });
  }
}

async function build() {
  const { chapters: chs, lectures, total } = expand();
  const counts = { 基礎: 0, 画像: 0 };
  chs.forEach((c) => c.seq.forEach((e) => { if (e.type === "base") counts["基礎"]++; if (e.type === "img") counts["画像"]++; }));

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = `${COURSE.title} 解説`;
  coverSlide(pres, total, counts);
  roadmapSlide(pres, lectures);

  const answers = [];
  for (let i = 0; i < chs.length; i++) {
    const ch = chs[i], label = `第${i + 1}章　${ch.name}`;
    if (ch.lectureBefore) lectureSlide(pres, ch.lectureBefore);
    chapterDivider(pres, i, ch, ch.first, ch.last);
    for (const e of ch.seq) {
      if (e.type === "lecture") lectureSlide(pres, e.lecture);
      else if (e.type === "solve") solveSlide(pres, e.block, label);
      else if (e.type === "sm") (e.sm.cards ? L.cardSlide : L.tableSlide)(pres, e.sm);
      else if (e.type === "sub") subDivider(pres, e.label, label);
      else {
        if (e.type === "base") { await L.questionSlide(pres, e.q, e.n); await L.answerSlide(pres, e.q, e.n); }
        else { await imgQuestionSlide(pres, e.q, e.n); await imgAnswerSlide(pres, e.q, e.n, e.id); }
        answers.push({ n: e.n, img: e.type === "img", ans: e.q.ans.join("・") });
      }
    }
  }
  answerListSlides(pres, answers);

  const file = path.join(OUT, `${COURSE.title}_解説（統合版）.pptx`);
  await pres.writeFile({ fileName: file });
  console.log("wrote", file, "questions:", total, counts);
}

build();

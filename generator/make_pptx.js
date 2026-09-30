// 解説用 PPT（テーマ別）
const path = require("path");
const pptxgen = require("pptxgenjs");
const DATA = require("./data");
const SUMMARY = require("./summary");
const { NAVY, ICE, ACC, F, title, coverSlide, dividerSlide, tableSlide, cardSlide, questionSlide, answerSlide } = require("./slides_lib");

const OUT = process.argv[2] || ".";

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
  s.addNotes(require("./talk").themeAnswerList(theme));
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

// course.js を展開して、通し番号・演習ブロック付きの並びを作る（PPT・Word共通）
const DATA = require("./data");
const IMG = require("./img_data");
const COURSE = require("./course");

function baseQs(themeId, from, to) {
  const all = DATA.find((t) => t.id === themeId).sections.flatMap((s) => s.qs);
  return all.slice(from - 1, to).map((q, i) => ({ ...q, origin: `${themeId}-[${from + i}]` }));
}

function expand() {
  let n = 0, blockNo = 0;
  const blocks = [];
  const chapters = COURSE.chapters.map((ch, ci) => {
    const seq = [];
    let block = null;
    for (const it of ch.items) {
      if (it[0] === "solve") {
        block = { no: ++blockNo, chapter: ci, minutes: it[1], first: null, last: null, img: 0, base: 0 };
        blocks.push(block);
        seq.push({ type: "solve", block });
      } else if (it[0] === "sm") seq.push({ type: "sm", sm: ch.summaries[it[1]] });
      else if (it[0] === "sub") seq.push({ type: "sub", label: it[1] });
      else {
        const qs = it[0] === "base" ? baseQs(it[1], it[2], it[3]).map((q) => ({ type: "base", q }))
          : [{ type: "img", id: it[1], q: IMG[it[1]] }];
        for (const e of qs) {
          e.n = ++n;
          if (block) { if (block.first == null) block.first = e.n; block.last = e.n; block[e.type]++; }
          seq.push(e);
        }
      }
    }
    // sm 指定がない章は、最初の演習ブロックの直後に要点まとめを入れる
    if (!seq.some((e) => e.type === "sm")) {
      const k = seq.findIndex((e) => e.type === "solve");
      seq.splice(k + 1, 0, ...ch.summaries.map((sm) => ({ type: "sm", sm })));
    }
    const qs = seq.filter((e) => e.type === "base" || e.type === "img");
    return { ...ch, seq, count: qs.length, first: qs[0].n, last: qs[qs.length - 1].n };
  });
  return { chapters, blocks, total: n };
}

module.exports = { COURSE, expand };

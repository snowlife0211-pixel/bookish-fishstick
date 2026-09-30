// course.js を展開して、通し番号・講義回・演習ブロック付きの並びを作る（PPT・Word共通）
const DATA = require("./data");
const IMG = require("./img_data");
const COURSE = require("./course");

function baseQs(themeId, from, to) {
  const all = DATA.find((t) => t.id === themeId).sections.flatMap((s) => s.qs);
  return all.slice(from - 1, to).map((q, i) => ({ ...q, origin: `${themeId}-[${from + i}]` }));
}

function expand() {
  const pace = COURSE.pace;
  let n = 0;
  const blocks = [], lectures = [];
  let lecture = null;
  const chapters = COURSE.chapters.map((ch, ci) => {
    const seq = [];
    let block = null, lectureBefore = null;
    ch.items.forEach((it, k) => {
      if (it[0] === "lecture") {
        lecture = { no: it[1], blocks: [] };
        lectures.push(lecture);
        // 章の先頭なら章扉の前に出す
        if (k === 0) lectureBefore = lecture; else seq.push({ type: "lecture", lecture });
      } else if (it[0] === "solve") {
        block = { no: blocks.length + 1, chapter: ci, chapterName: ch.name, kind: ch.kind, lecture: lecture.no,
          fixedSolve: it[1], first: null, last: null, img: 0, base: 0, sm: 0, subs: [] };
        blocks.push(block);
        lecture.blocks.push(block);
        seq.push({ type: "solve", block });
      } else if (it[0] === "sm") {
        seq.push({ type: "sm", sm: ch.summaries[it[1]] });
        if (block) block.sm++;
      } else if (it[0] === "sub") {
        seq.push({ type: "sub", label: it[1] });
        if (block) block.subs.push(it[1]);
      } else {
        const qs = it[0] === "base" ? baseQs(it[1], it[2], it[3]).map((q) => ({ type: "base", q }))
          : [{ type: "img", id: it[1], q: IMG[it[1]] }];
        for (const e of qs) {
          e.n = ++n;
          if (block.first == null) block.first = e.n;
          block.last = e.n;
          block[e.type]++;
          seq.push(e);
        }
      }
    });
    // sm 指定がない章は、最初の演習ブロックの直後に要点まとめを入れる
    if (!seq.some((e) => e.type === "sm") && ch.summaries.length) {
      const k = seq.findIndex((e) => e.type === "solve");
      seq.splice(k + 1, 0, ...ch.summaries.map((sm) => ({ type: "sm", sm })));
      seq[k].block.sm += ch.summaries.length;
    }
    const qs = seq.filter((e) => e.type === "base" || e.type === "img");
    return { ...ch, seq, lectureBefore, count: qs.length, first: qs[0].n, last: qs[qs.length - 1].n };
  });

  // 時間の目安
  for (const b of blocks) {
    b.count = b.last - b.first + 1;
    b.solve = b.fixedSolve || Math.ceil(b.base * pace.solve.base + b.img * pace.solve.img);
    b.explain = Math.ceil(b.base * pace.explain.base + b.img * pace.explain.img + b.sm * pace.explain.sm);
    b.total = b.solve + b.explain;
    b.label = b.subs.length ? `${b.chapterName}（${b.subs.join("・")}）` : b.chapterName;
  }
  // 同じ章を小見出しなしで分けたブロックには（前半）（後半）を付ける
  for (const b of blocks) {
    const same = blocks.filter((x) => x.chapter === b.chapter && !x.subs.length);
    if (!b.subs.length && same.length === 2) b.part = same[0] === b ? "前半" : "後半";
  }
  for (const l of lectures) {
    l.first = l.blocks[0].first;
    l.last = l.blocks[l.blocks.length - 1].last;
    l.total = l.blocks.reduce((a, b) => a + b.total, 0);
  }
  return { chapters, blocks, lectures, total: n };
}

module.exports = { COURSE, expand };

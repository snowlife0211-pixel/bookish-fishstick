// 発表者ノート（喋るメモ）の文面。スライドの内容から自動で作る
const NUM = "①②③④⑤";
const join = (lines) => lines.filter((l) => l != null && l !== "").join("\n");

// 「誤っている」「ないのは」など、逆を問う問題か
const REVERSE = /誤っている|関与しない|含まない|でないのは|にないのは|ないのは|最も低い/;
const reverseWord = (q) => (q.match(REVERSE) || [null])[0];
const isWrongStatement = (q) => /誤っている/.test(q);
const twoChoice = (q) => /２つ選べ|2つ選べ/.test(q);

function choiceLines(q) {
  return q.choices.map((c, i) => {
    const ok = q.ans.includes(i + 1);
    const note = q.notes && q.notes[i] ? `：${q.notes[i]}` : "";
    return `・${i + 1}「${c}」${note}${ok ? "　← 正解" : ""}`;
  });
}

function answerIntro(q) {
  const ans = q.ans.map((a) => `${a}番「${q.choices[a - 1]}」`).join("と");
  return `正解は ${ans} です。`;
}

function modality(text) {
  const m = [["SIAS", "SIASの図"], ["眼球運動", "眼球運動の図"], ["顔面の状態", "顔面の図"], ["兎眼", "脳神経の診察の図"],
    ["拡散強調", "MRIの拡散強調像（DWI）"], ["FLAIR", "MRIのFLAIR像"], ["MRA", "MRIとMRA"], ["T1", "MRIのT1強調像"], ["T2", "MRIのT2強調像"],
    ["MRI", "頭部MRI"], ["脳血管造影", "脳血管造影"], ["CT", "頭部CT"]];
  const hit = m.find(([k]) => text.includes(k));
  return hit ? hit[1] : "図";
}

module.exports = {
  cover: (title, total) => join([
    `【導入】今日から「${title}」の演習をします。全部で${total}問です。`,
    "進め方は、配布プリントの問題をまず自分で解いてもらい、そのあとスライドで解説します。",
    "国家試験でよく出るところを集めているので、間違えたところはメモ欄に理由を書いておきましょう。",
  ]),

  roadmap: (lectures, shortLabel) => join([
    `【全体の流れ】この演習は全${lectures.length}回です。`,
    ...lectures.map((l) => `・第${l.no}回：問${l.first}〜問${l.last}（${l.blocks.map(shortLabel).join("／")}）`),
    "どの回も「解く → 解説」を3回くり返します。",
    "基礎で覚えたことを、すぐ次の画像問題で使うように並べています。基礎の問題もあとで画像を読むための準備だと思って聞いてください。",
  ]),

  lecture: (l, shortLabel) => join([
    `【第${l.no}回の始まり】今日は問${l.first}から問${l.last}までの${l.last - l.first + 1}問をやります。`,
    ...l.blocks.map((b) => `・演習${b.no}：${shortLabel(b)}（問${b.first}〜${b.last}、解く${b.solve}分 → 解説 約${b.explain}分）`),
    `全体で約${l.total}分の予定です。時間が押したら、似た問題は正解の確認だけにします。`,
  ]),

  chapter: (i, ch) => join([
    `【第${i + 1}章　${ch.name}】`,
    ch.lead + "。",
    ch.kind === "画像"
      ? "ここからは画像問題です。まず『CTかMRIか』『白いか黒いか』『どこの高さの断面か』を確認するくせをつけましょう。"
      : "ここは基礎の問題です。あとの画像問題を解くための土台になるので、場所と働きをセットで覚えましょう。",
    `問題は[${ch.first}]〜[${ch.last}]です。`,
  ]),

  sub: (label, chName) => join([
    `【${label}】`,
    `${chName}の中の「${label}」のまとまりです。`,
    "ここから数問は同じテーマなので、共通する見方を意識して解説を聞いてください。",
  ]),

  solve: (b) => join([
    `【演習${b.no}】では、プリントの問${b.first}から問${b.last}までを解いてください。`,
    `${b.count}問で、時間は${b.solve}分が目安です。${b.img ? "画像問題は、先に画像を見て所見を考えてから選択肢を読むと解きやすいです。" : ""}`,
    "解答は巻末の解答欄に書いてください。特に指示のない問題は1つ選びます。",
    "（時間になったら）はい、そこまでにしましょう。解説に進みます。",
    `解説の目安は約${b.explain}分です。`,
  ]),

  summary: (sm) => {
    const lines = [`【要点まとめ】${sm.title}`];
    if (sm.sub) lines.push(sm.sub + "。");
    if (sm.cards) {
      for (const c of sm.cards) lines.push(`■${c.head}：${c.items.join("、")}`);
    } else {
      const h = sm.header;
      for (const r of sm.rows) {
        const vals = r.slice(1);
        const same = vals.length > 1 && vals.every((v) => v === vals[0]);
        const rest = same ? `${vals[0]}（${h.slice(1).join("・")}とも同じ）` : vals.map((v, k) => (h[k + 1] ? `${h[k + 1]}は${v}` : v)).join("、");
        lines.push(`・${r[0]}：${rest}`);
      }
    }
    if (sm.note) lines.push(`最後に、${sm.note}。ここは試験でよく問われます。`);
    lines.push("このあとの問題で、この表のどこが使えるかを確認しながら解説します。");
    return join(lines);
  },

  baseQuestion: (q, n) => join([
    `【問${n}】（${q.src}）`,
    `問題文：${q.q}`,
    reverseWord(q.q) ? `※「${reverseWord(q.q)}」ものを選ぶ問題です。読み飛ばしに注意するよう声をかけます。` : "",
    twoChoice(q.q) ? "※2つ選ぶ問題です。" : "",
    q.fig ? "図の番号の位置を、まず一緒に確認しましょう。" : "",
    "何番を選んだか、手を挙げてもらいます（または近くの人と答え合わせ）。",
  ]),

  baseAnswer: (q, n) => join([
    `【問${n} 解説】`,
    answerIntro(q),
    isWrongStatement(q.q) ? "「誤っている」ものを選ぶ問題なので、正解の選択肢が誤った内容になっています。" : "",
    "選択肢を順に確認します。",
    ...choiceLines(q),
    q.point ? `ポイント：${q.point}` : "",
  ]),

  imgQuestion: (q, n) => {
    const mod = modality(q.q + (q.find || ""));
    const scan = !mod.endsWith("図");
    return join([
    `【問${n}・画像問題】（${q.src}）`,
    `これは${mod}です。${scan ? "まず問題文を読む前に、画像のどこがおかしいかを探してもらいます。" : "まず図をよく見て、何をしている場面かを確認してもらいます。"}`,
    scan ? "「白いところ（黒いところ）はどこか」「右か左か」「どの高さの断面か」を確認しましょう。" : "「左右でどこが違うか」「どこまで動いているか」に注目しましょう。",
    `問題文：${q.q}`,
    reverseWord(q.q) ? `※「${reverseWord(q.q)}」ものを選ぶ問題です。` : "",
    "何番を選んだか確認します。",
  ]);
  },

  imgAnswer: (q, n) => join([
    `【問${n} 解説】`,
    `画像所見：${q.find}。`,
    answerIntro(q),
    "選択肢を順に確認します。",
    ...choiceLines(q),
    `ポイント：${q.point}`,
  ]),

  answerList: (page, pages) => join([
    `【正答一覧 ${page}/${pages}】`,
    "最後に正答を一覧で確認します。★は画像問題です。",
    "正解数を解答欄に書いて、間違えた問題はメモ欄に『なぜ間違えたか』を一言書いておきましょう。",
    page === pages ? "お疲れさまでした。画像問題は、画像を見た瞬間に部位と病態が浮かぶまで繰り返し見直してください。" : "",
  ]),

  themeCover: (theme, total) => join([
    `【導入】テーマ${theme.id}「${theme.title}」の解説です。全${total}問です。`,
    "配布プリントで解いてもらったあと、1問ずつ解説します。",
  ]),

  themeDivider: (label, sub) => join([`【${label}】`, sub ? `${sub}です。` : "", `ここから「${label}」の問題に入ります。`]),

  themeAnswerList: (theme) => join([`【正答一覧】テーマ${theme.id}「${theme.title}」の正答を確認します。`, "正解数を記録して、間違えた問題を見直しましょう。"]),

  NUM,
};

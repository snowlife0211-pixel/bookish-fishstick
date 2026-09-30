# 教材生成スクリプト

`materials/` の Word（学生配布用問題）と PPT（解説用）を生成するスクリプトです。

- `data.js` … 基礎問題（02 大脳・中脳・小脳・延髄／04 交感神経・副交感神経）の問題・正解・解説・POINT
- `img_data.js` … 画像問題25問の問題・正解・画像所見・解説・POINT（画像は `images/`）
- `img_notes.json` … 画像問題編（教員用）の元の解説メモ（統合版の発表者ノートに入る）
- `course.js` … 統合版の章立て（どの問題をどの順番で出すか）と、章ごとの要点まとめ
- `summary.js` … テーマ別版の要点まとめスライド
- `figures.js` … 模式図（Penfield 感覚野・中脳横断面）
- `slides_lib.js` … スライド共通部品
- `make_docx.js` / `make_pptx.js` … テーマ別の Word・PPT を生成
- `course_seq.js` … 章立てを展開して通し番号・演習ブロックを作る（PPT・Word共通）
- `make_course.js` … 統合版 PPT（基礎＋画像問題）を生成
- `make_course_docx.js` … 統合版の学生配布用 Word（問題のみ・解答欄付き）を生成

```bash
cd generator
npm install docx pptxgenjs sharp
node make_docx.js ../materials
node make_pptx.js ../materials
node make_course.js ../materials
node make_course_docx.js ../materials
```

問題や解説を直すときは `data.js` / `img_data.js` を、章の順番・講義回・演習の区切りを変えるときは `course.js` を編集して再生成してください。

時間の目安は `course.js` の `pace`（1問あたりの分数）から自動計算します。

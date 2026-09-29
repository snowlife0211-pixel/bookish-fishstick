# 教材生成スクリプト

`materials/` の Word（学生配布用問題）と PPT（解説用）を生成するスクリプトです。

- `data.js` … 問題・正解・選択肢ごとの解説・POINT
- `summary.js` … セクション冒頭の要点まとめスライド
- `figures.js` … 問題用の模式図（Penfield 感覚野・中脳横断面）
- `make_docx.js` / `make_pptx.js` … 生成本体

```bash
cd generator
npm install docx pptxgenjs sharp
node make_docx.js ../materials
node make_pptx.js ../materials
```

問題や解説を直すときは `data.js` を編集して再生成してください。

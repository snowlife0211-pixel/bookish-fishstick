// 問題用の模式図（SVG → PNG）
const sharp = require("sharp");

function num(x, y, n, hl) {
  const fill = hl ? "#C0392B" : "#FFFFFF";
  const tc = hl ? "#FFFFFF" : "#222222";
  return `<circle cx="${x}" cy="${y}" r="14" fill="${fill}" stroke="#222" stroke-width="2"/>` +
    `<text x="${x}" y="${y + 6}" font-family="Arial" font-size="17" font-weight="bold" text-anchor="middle" fill="${tc}">${n}</text>`;
}
function arrow(x1, y1, x2, y2) {
  return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#222" stroke-width="2" marker-end="url(#ah)"/>`;
}
const defs = `<defs><marker id="ah" markerWidth="10" markerHeight="10" refX="8" refY="5" orient="auto"><path d="M0,0 L10,5 L0,10 z" fill="#222"/></marker></defs>`;

// 感覚野の片側前額断（左半球を正面から、左が内側）
function penfieldSvg(hl) {
  const h = (n) => hl && n === 1;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="420" height="380" viewBox="0 0 420 380">
  <rect width="420" height="380" fill="#FFFFFF"/>${defs}
  <path d="M110,250 L110,80 Q110,60 135,60 L200,60 Q330,60 350,190 Q360,250 330,290 Q300,320 270,300"
        fill="none" stroke="#222" stroke-width="3"/>
  <path d="M110,250 Q130,255 140,270" fill="none" stroke="#222" stroke-width="2"/>
  <path d="M245,300 Q235,330 215,320 M245,300 Q255,340 285,335" fill="none" stroke="#222" stroke-width="2"/>
  <line x1="110" y1="280" x2="110" y2="350" stroke="#222" stroke-width="2"/>
  <text x="60" y="370" font-family="sans-serif" font-size="14" fill="#555">内側（大脳縦裂）</text>
  <text x="300" y="370" font-family="sans-serif" font-size="14" fill="#555">外側</text>
  ${arrow(40, 110, 102, 110)}${num(24, 110, 1, h(1))}
  ${arrow(150, 34, 150, 54)}${num(150, 18, 2, false)}
  ${arrow(240, 34, 240, 56)}${num(240, 18, 3, false)}
  ${arrow(390, 150, 348, 160)}${num(400, 136, 4, false)}
  ${arrow(395, 240, 360, 240)}${num(400, 222, 5, false)}
  ${hl ? `<text x="30" y="170" font-family="sans-serif" font-size="15" font-weight="bold" fill="#C0392B">足・下肢</text>
  <text x="165" y="100" font-family="sans-serif" font-size="13" fill="#555">体幹</text>
  <text x="250" y="110" font-family="sans-serif" font-size="13" fill="#555">上肢・手</text>
  <text x="275" y="200" font-family="sans-serif" font-size="13" fill="#555">顔面</text>
  <text x="225" y="262" font-family="sans-serif" font-size="13" fill="#555">口・舌・咽頭</text>` : ""}
</svg>`;
}

// 中脳横断面（上が背側）
function midbrainSvg(hl) {
  const lbl = hl
    ? [["中脳蓋（上丘）", 30, 40], ["赤核", 20, 150], ["黒質", 20, 290], ["中脳水道", 300, 60], ["大脳脚（錐体路）", 290, 325]]
    : null;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="440" height="360" viewBox="0 0 440 360">
  <rect width="440" height="360" fill="#FFFFFF"/>
  <path d="M170,50 Q190,30 220,40 Q250,30 270,50 Q300,70 300,110 Q320,150 330,200 Q345,250 300,275
           Q260,290 240,275 L230,290 L210,290 L200,275 Q180,290 140,275 Q95,250 110,200 Q120,150 140,110 Q140,70 170,50 Z"
        fill="#D5D8DC" stroke="#555" stroke-width="2"/>
  <circle cx="220" cy="95" r="5" fill="#fff" stroke="#555"/>
  <circle cx="180" cy="160" r="17" fill="#909497" stroke="#444"/>
  <circle cx="260" cy="160" r="17" fill="#909497" stroke="#444"/>
  <path d="M125,215 Q160,240 205,225" fill="none" stroke="#3B3B3B" stroke-width="14" stroke-linecap="round"/>
  <path d="M235,225 Q280,240 315,215" fill="none" stroke="#3B3B3B" stroke-width="14" stroke-linecap="round"/>
  ${hl ? `<path d="M125,240 Q150,275 205,268 L205,250 Q160,255 130,228 Z" fill="#E74C3C" opacity="0.55"/>
         <path d="M315,240 Q290,275 235,268 L235,250 Q280,255 310,228 Z" fill="#E74C3C" opacity="0.55"/>` : ""}
  <line x1="95" y1="45" x2="190" y2="70" stroke="#222" stroke-width="1.5"/>${num(80, 42, 1, false)}
  <line x1="75" y1="160" x2="163" y2="160" stroke="#222" stroke-width="1.5"/>${num(60, 160, 2, false)}
  <line x1="80" y1="275" x2="140" y2="222" stroke="#222" stroke-width="1.5"/>${num(66, 285, 3, false)}
  <line x1="350" y1="45" x2="225" y2="92" stroke="#222" stroke-width="1.5"/>${num(364, 40, 4, false)}
  <line x1="360" y1="300" x2="300" y2="255" stroke="#222" stroke-width="1.5"/>${num(374, 308, 5, !!hl)}
  ${lbl ? lbl.map(([t, x, y], i) => `<text x="${[100, 30, 20, 300, 250][i]}" y="${[20, 195, 318, 25, 345][i]}" font-family="sans-serif" font-size="14" ${i === 4 ? 'font-weight="bold" fill="#C0392B"' : 'fill="#333"'}>${t}</text>`).join("") : ""}
</svg>`;
}

async function png(svg) {
  return sharp(Buffer.from(svg), { density: 200 }).png().toBuffer();
}

module.exports = {
  penfield: (hl) => png(penfieldSvg(hl)),
  midbrain: (hl) => png(midbrainSvg(hl)),
  sizes: { penfield: [420, 380], midbrain: [440, 360] },
};

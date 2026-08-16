const fs = require("fs");
const html = fs.readFileSync("index.html", "utf8");
const m = html.match(/<script[^>]*type=["']text\/babel["'][^>]*>([\s\S]*?)<\/script>/i);
if (!m) {
  console.error("no babel script found");
  process.exit(1);
}
const code = m[1];
const lines = code.split(/\n/);
let bad = 0;
for (let i = 0; i < lines.length; i++) {
  const line = lines[i];
  if (/\|\|"[—–-]\}/.test(line)) {
    bad++;
    console.log("UNCLOSED_EMDASH_LINE", i + 1, line.trim().slice(0, 180));
  }
  const exprs = line.match(/\$\{[^}]*\}/g) || [];
  for (const ex of exprs) {
    const quotes = (ex.match(/"/g) || []).length;
    if (quotes % 2 !== 0) {
      bad++;
      console.log("ODD_QUOTES_IN_EXPR", i + 1, ex);
    }
  }
}
// Package H block sanity: PRINTOS_DEPT through function App
const hStart = code.indexOf("PRINTOS_DEPT_LS_KEY");
const hEnd = code.indexOf("function App()");
if (hStart >= 0 && hEnd > hStart) {
  const h = code.slice(hStart, hEnd);
  const hm = h.match(/\|\|"[—–-]\}/);
  if (hm) {
    bad++;
    console.log("H_BLOCK_STILL_BROKEN", hm[0]);
  } else {
    console.log("H_BLOCK_OK");
  }
}
console.log("bad_total", bad);
process.exit(bad ? 1 : 0);

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
const hOk = code.includes("PRINTOS_DEPT_LS_KEY") ? !/\|\|"[—–-]\}/.test(code.slice(code.indexOf("PRINTOS_DEPT_LS_KEY"), code.indexOf("function App()") || code.length)) : true;
const hardFail = bad > 0 && lines.some((_, i) => {
  // only fail hard on unclosed emdash patterns
  return false;
});
let unclosed = 0;
for (let i = 0; i < lines.length; i++) {
  if (/\|\|"[—–-]\}/.test(lines[i])) unclosed++;
}
if (unclosed > 0) {
  console.error("FAIL: unclosed em-dash string literals:", unclosed);
  process.exit(1);
}
console.log("PASS: no unclosed em-dash quotes (nested-template odd-quote warnings are informational)");
process.exit(0);

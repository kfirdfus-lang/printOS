const fs = require("fs");
const path = require("path");
const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
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

let unclosed = 0;
for (let i = 0; i < lines.length; i++) {
  if (/\|\|"[—–-]\}/.test(lines[i])) unclosed++;
}
if (unclosed > 0) {
  console.error("FAIL: unclosed em-dash string literals:", unclosed);
  process.exit(1);
}
console.log("PASS: no unclosed em-dash quotes (nested-template odd-quote warnings are informational)");

// Real JSX/JS syntax check via @babel/parser (catches missing }} in style={{...}})
let parser;
try {
  parser = require("@babel/parser");
} catch (e) {
  console.error("FAIL: @babel/parser not installed — run npm i");
  process.exit(1);
}
try {
  parser.parse(code, {
    sourceType: "script",
    plugins: ["jsx"],
    errorRecovery: false,
    allowReturnOutsideFunction: true,
  });
  console.log("PASS: @babel/parser JSX parse OK");
} catch (e) {
  const loc = e.loc ? ` (${e.loc.line}:${e.loc.column})` : "";
  console.error("FAIL: @babel/parser" + loc + ": " + e.message);
  if (e.loc && lines[e.loc.line - 1]) {
    console.error("LINE:", lines[e.loc.line - 1].trim().slice(0, 240));
  }
  process.exit(1);
}
process.exit(0);

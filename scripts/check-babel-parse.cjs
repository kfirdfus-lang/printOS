const fs = require("fs");
const parser = require("@babel/parser");
const html = fs.readFileSync("index.html", "utf8");
const m = html.match(/<script[^>]*type=["']text\/babel["'][^>]*>([\s\S]*?)<\/script>/i);
if (!m) {
  console.error("no babel script");
  process.exit(1);
}
try {
  parser.parse(m[1], { sourceType: "script", plugins: ["jsx"] });
  console.log("Babel PASS");
} catch (e) {
  console.error("Babel FAIL", e.message);
  if (e.loc) console.error("at", e.loc.line, e.loc.column);
  process.exit(1);
}

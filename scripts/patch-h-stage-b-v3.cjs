const fs = require("fs");
const path = require("path");
const file = path.join(__dirname, "..", "index.html");
let html = fs.readFileSync(file, "utf8");
const block = fs.readFileSync(path.join(__dirname, "h-stage-b-v3-block.js.txt"), "utf8");
const start = html.indexOf("function UnifiedOrderQuoteModal(");
const end = html.indexOf("function CreateOrderModal(");
if (start < 0 || end < 0 || end <= start) {
  console.error("markers not found", start, end);
  process.exit(1);
}
html = html.slice(0, start) + block + "\n\n" + html.slice(end);
fs.writeFileSync(file, html);
console.log("replaced UnifiedOrderQuoteModal..QuotesLifecycleTab, bytes", block.length);

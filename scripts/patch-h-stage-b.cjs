const fs = require("fs");
const path = require("path");
const file = path.join(__dirname, "..", "index.html");
let html = fs.readFileSync(file, "utf8");
const block = fs.readFileSync(path.join(__dirname, "h-stage-b-block.js.txt"), "utf8");
const marker = "\nfunction CreateOrderModal({";
const idx = html.indexOf(marker);
if (idx < 0) {
  console.error("CreateOrderModal marker not found");
  process.exit(1);
}
if (html.includes("function UnifiedOrderQuoteModal(")) {
  console.log("UnifiedOrderQuoteModal already present — skip splice");
} else {
  html = html.slice(0, idx) + "\n" + block + html.slice(idx);
  fs.writeFileSync(file, html);
  console.log("spliced UnifiedOrderQuoteModal before CreateOrderModal");
}

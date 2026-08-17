const fs = require("fs");
const path = require("path");
const file = path.join(__dirname, "..", "index.html");
let html = fs.readFileSync(file, "utf8");
const neu = fs.readFileSync(path.join(__dirname, "open-document-pdf-v3.js.txt"), "utf8").trim();
const start = html.indexOf("window.openDocumentPDF = function(doc) {");
if (start < 0) {
  console.error("start not found");
  process.exit(1);
}
const after = html.slice(start);
const endRel = after.search(/\};\r?\n\r?\n<\/script>/);
if (endRel < 0) {
  console.error("end not found");
  process.exit(1);
}
const end = start + endRel + 2;
html = html.slice(0, start) + neu + html.slice(end);
fs.writeFileSync(file, html);
console.log("replaced openDocumentPDF");

import fs from "fs";
import path from "path";
import http from "http";
import { fileURLToPath } from "url";
import { createRequire } from "module";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, "..");
const indexPath = path.join(root, "index.html");
const html = fs.readFileSync(indexPath, "utf8");

// 1) Babel compile check (same as browser)
const babelStart = html.indexOf('<script type="text/babel">');
if (babelStart < 0) {
  console.error("FAIL: babel script block not found");
  process.exit(1);
}
const codeStart = babelStart + '<script type="text/babel">'.length;
const codeEnd = html.indexOf("</script>", codeStart);
const code = html.slice(codeStart, codeEnd);
const require = createRequire(import.meta.url);
let babel;
try {
  babel = require("@babel/standalone");
} catch {
  console.error("FAIL: install @babel/standalone (npm install --no-save @babel/standalone@7)");
  process.exit(1);
}
try {
  babel.transform(code, { presets: ["react"], filename: "index.html" });
  console.log("OK: Babel transform succeeded (no JSX/syntax errors)");
} catch (e) {
  console.error("FAIL: Babel transform:", e.message);
  if (e.loc) console.error("  at line", e.loc.line, "col", e.loc.column);
  process.exit(1);
}

// 2) Optional browser load + screenshot
const runBrowser = process.argv.includes("--browser");
if (!runBrowser) process.exit(0);

const { chromium } = await import("playwright");

const mime = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript",
  ".css": "text/css",
  ".png": "image/png",
  ".ico": "image/x-icon",
};

const server = http.createServer((req, res) => {
  const urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
  const filePath = path.join(root, urlPath === "/" ? "index.html" : urlPath.replace(/^\//, ""));
  if (!filePath.startsWith(root) || !fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    res.writeHead(404);
    res.end("Not found");
    return;
  }
  const ext = path.extname(filePath);
  res.writeHead(200, { "Content-Type": mime[ext] || "application/octet-stream" });
  res.end(fs.readFileSync(filePath));
});

await new Promise((resolve) => server.listen(8765, "127.0.0.1", resolve));
const shotPath = path.join(root, "verify-screenshot.png");

const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
const errors = [];
page.on("pageerror", (err) => errors.push(`pageerror: ${err.message}`));
page.on("console", (msg) => {
  if (msg.type() === "error") errors.push(`console: ${msg.text()}`);
});

try {
  await page.goto("http://127.0.0.1:8765/", { waitUntil: "networkidle", timeout: 60000 });
  await page.waitForFunction(
    () => {
      const root = document.getElementById("root");
      return root && root.childElementCount > 0;
    },
    { timeout: 30000 }
  );
  // Wait for login UI or dashboard (Supabase fetch); spinner alone means React parsed OK
  await page
    .waitForFunction(
      () => {
        const t = document.body.innerText || "";
        return t.includes("PrintOS") || t.includes("בחר את הפרופיל") || t.includes("טוען נתונים");
      },
      { timeout: 20000 }
    )
    .catch(() => {});
  await page.waitForTimeout(2000);
  const bodyText = await page.evaluate(() => document.body.innerText.slice(0, 500));
  const rootHtml = await page.evaluate(() => document.getElementById("root")?.innerHTML?.slice(0, 200) || "");
  const hasSpinner = rootHtml.includes("animation") || rootHtml.includes("spin");
  const hasLoginOrApp =
    bodyText.includes("PrintOS") ||
    bodyText.includes("בחר את הפרופיל") ||
    bodyText.includes("טוען נתונים") ||
    bodyText.includes("דשבורד");
  const hasContent = hasLoginOrApp || (rootHtml.length > 80 && hasSpinner);
  await page.screenshot({ path: shotPath, fullPage: false });
  console.log("OK: Page loaded, #root has children");
  console.log("Screenshot:", shotPath);
  console.log("Body preview:", bodyText.replace(/\s+/g, " ").slice(0, 120) + "...");
  if (errors.length) {
    console.warn("WARN: console errors:", errors.join("\n  "));
  }
  if (!hasContent) {
    console.error("FAIL: no React UI detected (possible white screen)");
    process.exit(1);
  }
  console.log(hasLoginOrApp ? "OK: Login or app UI visible" : "OK: React mounted (loading spinner — not a parse-error white screen)");
} catch (e) {
  await page.screenshot({ path: shotPath }).catch(() => {});
  console.error("FAIL: browser load:", e.message);
  if (errors.length) console.error("Errors:", errors.join("\n  "));
  process.exit(1);
} finally {
  await browser.close();
  server.close();
}

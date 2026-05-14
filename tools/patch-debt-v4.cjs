const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const idxPath = path.join(root, "index.html");
const v4Path = "C:/Users/kfird/Downloads/01-DebtDashboardV4.jsx";
const modalPath = "C:/Users/kfird/Downloads/02-DebtActionModal-updated.jsx";

let v4 = fs.readFileSync(v4Path, "utf8").split(/\r?\n/);
const v4Start = v4.findIndex((l) => l.startsWith("const DEBT_ACTION_TYPES"));
if (v4Start === -1) throw new Error("v4: no DEBT_ACTION_TYPES");
v4 = v4.slice(v4Start).join("\n").trimEnd() + "\n";

// sendEmail: days_overdue per invoice (filtered rows may lack inv.daysOverdue)
v4 = v4.replace(
  /const invoicesForEmail = filteredInvoices\.map\(inv => \(\{\s*doc_num: inv\.doc_num,\s*doc_balance: inv\.doc_balance,\s*doc_date: inv\.doc_date,\s*doc_payment_date: inv\.calculated_due_date\s*\?\s*inv\.calculated_due_date\.toISOString\(\)\.split\('T'\)\[0\]\s*:\s*inv\.doc_payment_date,\s*days_overdue: inv\.daysOverdue \|\| 0,\s*\}\)\);/,
  `const invoicesForEmail = filteredInvoices.map((inv) => {
        const dueMs = inv.calculated_due_date ? inv.calculated_due_date.getTime() : 0;
        const dOver = dueMs ? Math.max(0, Math.floor((todayMs - dueMs) / 86400000)) : 0;
        return {
        doc_num: inv.doc_num,
        doc_balance: inv.doc_balance,
        doc_date: inv.doc_date,
        doc_payment_date: inv.calculated_due_date
          ? inv.calculated_due_date.toISOString().split("T")[0]
          : inv.doc_payment_date,
        days_overdue: dOver,
      };});`,
);

v4 = v4.replace(
  /csv \+= 'מס,לקוח,תנאי תשלום,מס חשבוניות,פירוט חשבוניות,איש קשר,טלפון,סה"כ\\n';/,
  "csv += 'מס,לקוח,תנאי תשלום,מס חשבוניות,פירוט חשבוניות,איש קשר,טלפון,סה״כ\\n';",
);
v4 = v4.replace(
  /csv \+= `\\n,,,,,,"סה"כ:",\$\{fmtCurrency\(data\.grandTotal\)\}\\n`;/,
  "csv += `\\n,,,,,,סה״כ:,${fmtCurrency(data.grandTotal)}\\n`;",
);

v4 = v4.replace(
  /<td colSpan="4" style=\{\{ padding: 14, textAlign: 'right', fontSize: 14 \}\}>סה"כ \{clients\.length\} לקוחות:<\/td>/,
  "<td colSpan=\"4\" style={{ padding: 14, textAlign: 'right', fontSize: 14 }}>סה״כ {clients.length} לקוחות:</td>",
);
v4 = v4.replace(
  /<strong style=\{\{ color: '#1E3A52' \}\}>נטלי פתרונות הדפסה בע"מ<\/strong> \| ח\.פ\. 517205332 \| שד' הר ציון 104, תל אביב \| 03-6815703/,
  "<strong style={{ color: '#1E3A52' }}>נטלי פתרונות הדפסה בע״מ</strong> | ח.פ. 517205332 | שד&apos; הר ציון 104, תל אביב | 03-6815703",
);

v4 = v4.replace(
  /await new Promise\(r => setTimeout\(r, 1500\)\);/,
  "await new Promise((r) => setTimeout(r, 1500));",
);
v4 = v4.replace(
  /alert\('שגיאה ברענון: ' \+ \(e\.message \|\| e\)\);/,
  "alert(\"שגיאה ברענון: \" + (e && e.message ? e.message : String(e)));",
);

let modal = fs.readFileSync(modalPath, "utf8").split(/\r?\n/);
const mStart = modal.findIndex((l) => l.startsWith("function fmtCurrencySimple"));
const mDebt = modal.findIndex((l) => l.startsWith("function DebtActionModal"));
if (mDebt === -1) throw new Error("modal: no DebtActionModal");
const mEnd = modal.findIndex((l, i) => i > mDebt && l.startsWith("function fmtCurrencySimple"));
let modalBlock;
if (mStart !== -1 && mStart < mDebt) {
  modalBlock = modal.slice(mStart, mEnd === -1 ? undefined : mEnd).concat(modal.slice(mDebt)).join("\n").trim() + "\n";
} else {
  const fmtStart = modal.findIndex((l) => l.startsWith("function fmtCurrencySimple"));
  const fmtEnd = modal.length;
  modalBlock = modal.slice(mDebt, fmtStart).join("\n").trim() + "\n\n" + modal.slice(fmtStart, fmtEnd).join("\n").trim() + "\n";
}

let html = fs.readFileSync(idxPath, "utf8");

const emailMarker = "\n\n// ========================================\n// Modal לניהול מיילים + שליחה\n// ========================================\nfunction EmailModal";
const debtStart = html.indexOf("const DEBT_ACTION_TYPES = {");
const emailIdx = html.indexOf(emailMarker);
if (debtStart === -1 || emailIdx === -1) throw new Error("index markers for debt→email");

const monthlyStart = html.indexOf("\n\n// ========================================\n// הדפסת רשימה חודשית\n// ========================================\nfunction MonthlyPrintReport");
const debtModalIdx = html.indexOf("\nfunction DebtActionModal({");
if (monthlyStart === -1 || debtModalIdx === -1 || monthlyStart > debtModalIdx) {
  throw new Error("monthly / debt modal markers");
}

const emailBlock = html.slice(emailIdx, monthlyStart);
const afterDebtModal = html.slice(debtModalIdx);
const debtModalEnd = afterDebtModal.indexOf("\n\n// ── MAIN APP ──");
if (debtModalEnd === -1) throw new Error("MAIN APP after DebtActionModal");
const oldDebtModal = afterDebtModal.slice(0, debtModalEnd);
const rest = afterDebtModal.slice(debtModalEnd);

html = html.slice(0, debtStart) + v4 + emailBlock + "\n\n" + modalBlock + rest;

fs.writeFileSync(idxPath, html);
console.log("patched v4: dashboard+PrintReport, kept EmailModal, removed MonthlyPrintReport, updated DebtActionModal");

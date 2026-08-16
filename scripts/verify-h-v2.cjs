const fs = require("fs");
const h = fs.readFileSync("index.html", "utf8");
const checks = {
  MyDayItemDetail: h.includes("function MyDayItemDetail"),
  MyDayScreen: h.includes("function MyDayScreen"),
  WorkFilesAdminTab: h.includes("function WorkFilesAdminTab"),
  department_settings: h.includes('from("department_settings")'),
  old_departments_query: h.includes('from("departments")'),
  old_MyDayTaskDetail: h.includes("function MyDayTaskDetail"),
  App_once: (h.match(/function App\(\)/g) || []).length === 1,
  unclosed_emdash: /\|\|"[—–-]\}/.test(h),
};
console.log(checks);
const ok =
  checks.MyDayItemDetail &&
  checks.MyDayScreen &&
  checks.WorkFilesAdminTab &&
  checks.department_settings &&
  !checks.old_departments_query &&
  !checks.old_MyDayTaskDetail &&
  checks.App_once &&
  !checks.unclosed_emdash;
process.exit(ok ? 0 : 1);

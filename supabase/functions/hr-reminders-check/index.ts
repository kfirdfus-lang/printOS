// HR reminders check - runs on the 1st of every month via pg_cron (see
// 20260801222000_hr_reminders_cron.sql) and can be triggered manually from
// the employees dashboard.
// Alerts:
//   1. Every January - annual minimum-wage / pension-cap review reminder.
//   2. Every even January - biennial contract refresh with a labor lawyer.
//   3. Any month - employees whose contract review is more than 2 years old.
// Sends a single Hebrew HTML email only when there is at least one alert.

// @ts-ignore
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { rejectDisallowedInternalOrigin } from "../_shared/cors.ts";
// @ts-ignore
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const RECIPIENTS = ["kfir.dfus@gmail.com", "natalie.zem@gmail.com"];

serve(async (req) => {
  const originBlock = rejectDisallowedInternalOrigin(req);
  if (originBlock) return originBlock;

  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS_HEADERS });

  try {
    const supabase = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    let dryRun = false;
    try {
      const body = await req.json();
      dryRun = body?.dry_run === true;
    } catch (_) { /* empty body is fine */ }

    const now = new Date();
    const currentMonth = now.getMonth() + 1; // 1-12
    const currentYear = now.getFullYear();
    const currentDay = now.getDate();

    const alerts: string[] = [];
    const notes: string[] = [];

    // 1. January - minimum wage / caps review
    if (currentMonth === 1 && currentDay <= 7) {
      alerts.push(`
        <div class="alert alert-annual">
          <div class="alert-icon">📅</div>
          <div class="alert-content">
            <div class="alert-title">🎯 תזכורת שנתית - ינואר ${currentYear}</div>
            <div class="alert-body">
              <strong>זמן לבדיקת שכר מינימום ותקרות מול חשב השכר!</strong><br>
              נקודות לבדוק:<br>
              • עדכון שכר מינימום שנתי לעובדים שעתיים<br>
              • תקרות פנסיה - שכר מבוטח<br>
              • שכר לעובדי ייצור על פי שינויים בצווי הרחבה
            </div>
          </div>
        </div>
      `);
    }

    // 2. Biennial (even years) - contract refresh with a lawyer
    if (currentMonth === 1 && currentDay <= 7 && currentYear % 2 === 0) {
      alerts.push(`
        <div class="alert alert-biennial">
          <div class="alert-icon">📚</div>
          <div class="alert-content">
            <div class="alert-title">📖 תזכורת דו-שנתית ${currentYear}</div>
            <div class="alert-body">
              <strong>זמן לריענון הסכמי העסקה עם עורך דין!</strong><br>
              בעיקר:<br>
              • שינויים בחוקי עבודה<br>
              • עדכוני צווי הרחבה<br>
              • התאמת נספחי סעיף 14<br>
              • ריענון תקנונים ארגוניים
            </div>
          </div>
        </div>
      `);
    }

    // 3. Per-employee: contract review older than 2 years.
    // Tolerant to the HR columns not existing yet (before the F1 migration runs).
    let employeesNeedingReview: Array<Record<string, unknown>> = [];
    const { data: employees, error: empError } = await supabase
      .from("employees")
      .select("id, full_name, role_title, last_contract_review_date, work_start_date")
      .eq("is_active", true);

    if (empError) {
      notes.push(`employees query failed (HR migration not applied yet?): ${empError.message}`);
    } else {
      const twoYearsAgo = new Date(now);
      twoYearsAgo.setFullYear(twoYearsAgo.getFullYear() - 2);

      employeesNeedingReview = (employees || []).filter((emp: Record<string, unknown>) => {
        if (!emp.last_contract_review_date) {
          if (!emp.work_start_date) return false;
          return new Date(String(emp.work_start_date)) < twoYearsAgo;
        }
        return new Date(String(emp.last_contract_review_date)) < twoYearsAgo;
      });

      if (employeesNeedingReview.length > 0) {
        const employeesList = employeesNeedingReview.map((emp) => `
          <li>
            <strong>${emp.full_name}</strong> (${emp.role_title || "—"}) -
            עדכון אחרון: ${emp.last_contract_review_date
              ? new Date(String(emp.last_contract_review_date)).toLocaleDateString("he-IL")
              : "מעולם לא עודכן"}
          </li>
        `).join("");

        alerts.push(`
          <div class="alert alert-employee-review">
            <div class="alert-icon">👤</div>
            <div class="alert-content">
              <div class="alert-title">👥 עובדים דורשים בדיקת עדכון (${employeesNeedingReview.length})</div>
              <div class="alert-body">
                עברו יותר משנתיים מהעדכון האחרון של ההסכם - כדאי לבדוק:<br>
                <ul>${employeesList}</ul>
              </div>
            </div>
          </div>
        `);
      }
    }

    if (alerts.length === 0) {
      return new Response(
        JSON.stringify({ success: true, message: "No HR alerts today", notes }),
        { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
      );
    }

    const html = `<!DOCTYPE html>
<html dir="rtl" lang="he">
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Heebo', Arial, sans-serif; background: #f9fafb; padding: 20px; margin: 0; direction: rtl; }
    .container { max-width: 700px; margin: 0 auto; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
    .header { background: linear-gradient(135deg, #0d9488 0%, #14b8a6 100%); color: white; padding: 24px 30px; }
    .header h1 { margin: 0; font-size: 22px; }
    .header p { margin: 6px 0 0 0; opacity: 0.9; font-size: 14px; }
    .content { padding: 20px 30px; }
    .alert { display: flex; gap: 16px; padding: 16px; border-radius: 10px; margin-bottom: 12px; border-right: 4px solid; }
    .alert-annual { background: #fef3c7; border-color: #f59e0b; }
    .alert-biennial { background: #dbeafe; border-color: #3b82f6; }
    .alert-employee-review { background: #fed7aa; border-color: #f97316; }
    .alert-icon { font-size: 28px; }
    .alert-title { font-weight: 700; color: #1f2937; margin-bottom: 6px; font-size: 15px; }
    .alert-body { color: #374151; font-size: 14px; line-height: 1.6; }
    .alert-body ul { padding-right: 20px; margin: 6px 0; }
    .footer { background: #f9fafb; padding: 15px 30px; text-align: center; font-size: 12px; color: #6b7280; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>👥 תזכורות HR</h1>
      <p>${now.toLocaleDateString("he-IL", { day: "numeric", month: "long", year: "numeric" })}</p>
    </div>
    <div class="content">
      ${alerts.join("")}
    </div>
    <div class="footer">
      🖨️ נטלי פרינט • תזכורות HR אוטומטיות
    </div>
  </div>
</body>
</html>`;

    if (!dryRun) {
      const emailResponse = await supabase.functions.invoke("send-email", {
        body: {
          to: RECIPIENTS,
          subject: `👥 תזכורות HR - ${now.toLocaleDateString("he-IL")} (${alerts.length})`,
          html,
          from: "PrintOS - HR <orders@natalie-print.com>",
        },
      });
      if (emailResponse.error) throw emailResponse.error;
    }

    return new Response(
      JSON.stringify({
        success: true,
        dry_run: dryRun,
        alerts_count: alerts.length,
        employees_needing_review: employeesNeedingReview.length,
        notes,
      }),
      { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
    );
  } catch (error) {
    return new Response(
      JSON.stringify({ error: error instanceof Error ? error.message : String(error) }),
      { status: 500, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
    );
  }
});

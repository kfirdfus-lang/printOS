// Daily birthday greetings for active employees with email + birth_date.
// Cron: 08:00 Israel (05:00 UTC) — see 20260804011000_birthday_greetings_cron.sql
// Also invokable manually from the employees dashboard.

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

const ADMIN_RECIPIENTS = ["kfir.dfus@gmail.com", "natalie.zem@gmail.com"];

function israelYmd(d = new Date()): { month: number; day: number } {
  // Asia/Jerusalem — reliable across DST
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Jerusalem",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(d);
  const get = (t: string) => Number(parts.find((p) => p.type === t)?.value || 0);
  return { month: get("month"), day: get("day") };
}

function buildBirthdayEmail(employee: { full_name: string }): string {
  const firstName = String(employee.full_name || "").split(" ")[0] || "חבר/ה";
  return `<!DOCTYPE html>
<html dir="rtl" lang="he">
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Heebo', Arial, sans-serif; background: linear-gradient(135deg, #fef3c7 0%, #fecaca 100%); padding: 20px; margin: 0; direction: rtl; }
    .container { max-width: 550px; margin: 0 auto; background: white; border-radius: 20px; overflow: hidden; box-shadow: 0 10px 40px rgba(0,0,0,0.15); }
    .header { padding: 40px 30px; text-align: center; background: linear-gradient(135deg, #ec4899 0%, #f472b6 50%, #f59e0b 100%); color: white; }
    .cake { font-size: 60px; margin-bottom: 10px; }
    .header h1 { margin: 0; font-size: 28px; }
    .header p { margin: 8px 0 0 0; opacity: 0.9; font-size: 15px; }
    .content { padding: 30px; text-align: center; }
    .greeting { font-size: 18px; color: #1f2937; line-height: 1.7; margin-bottom: 20px; }
    .balloons { font-size: 32px; margin: 20px 0; letter-spacing: 8px; }
    .signature { padding: 15px; background: #f0fdfa; border-radius: 10px; margin-top: 20px; }
    .signature-title { font-weight: 700; color: #0d9488; margin-bottom: 4px; }
    .signature-name { color: #6b7280; font-size: 14px; }
    .footer { background: #f9fafb; padding: 15px 30px; text-align: center; font-size: 12px; color: #6b7280; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="cake">🎂</div>
      <h1>יום הולדת שמח!</h1>
      <p>${firstName}, זה היום שלך 🎉</p>
    </div>
    <div class="content">
      <div class="greeting">
        <strong>${firstName} יקר/ה,</strong><br><br>
        אנחנו רוצים לאחל לך יום הולדת שמח מכל הלב! 🎊<br>
        שנה מלאת בריאות, שמחה, הצלחות והגשמת חלומות.<br>
        תודה שאת/ה חלק מהמשפחה שלנו בנטלי פרינט 💚
      </div>
      <div class="balloons">🎈 🎁 🎊 🎉 🎈</div>
      <div class="signature">
        <div class="signature-title">בברכה מכל הלב,</div>
        <div class="signature-name">צוות נטלי פרינט</div>
      </div>
    </div>
    <div class="footer">🖨️ נטלי פרינט • ${new Date().getFullYear()}</div>
  </div>
</body>
</html>`;
}

serve(async (req) => {
  const originBlock = rejectDisallowedInternalOrigin(req);
  if (originBlock) return originBlock;

  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS_HEADERS });

  try {
    const supabase = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    const { month: todayMonth, day: todayDay } = israelYmd();

    const { data: employees, error } = await supabase
      .from("employees")
      .select("id, full_name, email, birth_date, role_title")
      .eq("is_active", true)
      .not("birth_date", "is", null)
      .not("email", "is", null);

    if (error) throw error;

    const birthdayEmployees = (employees || []).filter((emp) => {
      if (!emp.birth_date || !String(emp.email || "").trim()) return false;
      // DATE as YYYY-MM-DD — avoid timezone shift from Date parsing
      const parts = String(emp.birth_date).slice(0, 10).split("-").map(Number);
      if (parts.length < 3 || !parts[1] || !parts[2]) return false;
      return parts[1] === todayMonth && parts[2] === todayDay;
    });

    if (birthdayEmployees.length === 0) {
      return new Response(
        JSON.stringify({ success: true, message: "No birthdays today", birthdays: 0 }),
        { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
      );
    }

    const results: Array<Record<string, unknown>> = [];

    for (const emp of birthdayEmployees) {
      try {
        const emailResponse = await supabase.functions.invoke("send-email", {
          body: {
            to: emp.email,
            subject: `🎂 יום הולדת שמח ${String(emp.full_name).split(" ")[0]}!`,
            html: buildBirthdayEmail(emp),
            from: "נטלי פרינט <orders@natalie-print.com>",
          },
        });
        results.push({
          name: emp.full_name,
          email: emp.email,
          success: !emailResponse.error,
          error: emailResponse.error?.message,
        });
      } catch (err) {
        results.push({
          name: emp.full_name,
          success: false,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }

    const namesHtml = birthdayEmployees.map((emp) =>
      `<li><strong>${emp.full_name}</strong>${emp.role_title ? ` - ${emp.role_title}` : ""}</li>`
    ).join("");

    const adminHtml = `<!DOCTYPE html>
<html dir="rtl" lang="he">
<head><meta charset="UTF-8"><style>
body { font-family: 'Heebo', Arial, sans-serif; background: #f9fafb; padding: 20px; direction: rtl; margin: 0; }
.container { max-width: 500px; margin: 0 auto; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
.header { background: linear-gradient(135deg, #ec4899 0%, #f472b6 100%); color: white; padding: 25px; text-align: center; }
.content { padding: 20px 25px; }
.content ul { padding-right: 20px; }
.content li { padding: 5px 0; }
</style></head>
<body>
  <div class="container">
    <div class="header"><h1>🎂 ימי הולדת היום</h1></div>
    <div class="content">
      <p>היום חוגגים יום הולדת בנטלי פרינט:</p>
      <ul>${namesHtml}</ul>
      <p style="color: #6b7280; font-size: 13px; margin-top: 15px;">💌 המערכת כבר שלחה להם ברכות אישיות במייל.</p>
    </div>
  </div>
</body></html>`;

    await supabase.functions.invoke("send-email", {
      body: {
        to: ADMIN_RECIPIENTS,
        subject: `🎂 היום יש ${birthdayEmployees.length} יום הולדת בנטלי פרינט`,
        html: adminHtml,
        from: "PrintOS - HR <orders@natalie-print.com>",
      },
    });

    return new Response(
      JSON.stringify({ success: true, birthdays: birthdayEmployees.length, results }),
      { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
    );
  } catch (error) {
    return new Response(
      JSON.stringify({ error: error instanceof Error ? error.message : String(error) }),
      { status: 500, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
    );
  }
});

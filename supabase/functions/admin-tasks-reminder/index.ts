// Package J — morning admin-tasks reminder
// Schedule: cron 0 5 * * * (05:00 UTC ≈ 08:00 Israel)
// SQL: supabase/migrations/20260823120000_admin_tasks_reminder_cron.sql
// Sends only when there are overdue or due-today open tasks.

// @ts-ignore
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { rejectDisallowedInternalOrigin } from "../_shared/cors.ts";
// @ts-ignore
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const RECIPIENTS = ["kfir.dfus@gmail.com", "natalie.zem@gmail.com"];

const ASSIGNEE_HE: Record<string, string> = {
  kfir: "כפיר",
  natalie: "נטלי",
  both: "שניהם",
};

function todayISO(): string {
  // Israel ≈ UTC+2/+3 — use Asia/Jerusalem wall date
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Jerusalem",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date());
}

function daysOverdue(due: string, today: string): number {
  const a = new Date(due + "T12:00:00Z").getTime();
  const b = new Date(today + "T12:00:00Z").getTime();
  return Math.max(0, Math.round((b - a) / 864e5));
}

function esc(s: unknown): string {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

serve(async (req) => {
  const originBlock = rejectDisallowedInternalOrigin(req);
  if (originBlock) return originBlock;
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: CORS_HEADERS });
  }

  try {
    const supabase = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    const today = todayISO();
    const { data, error } = await supabase
      .from("admin_tasks")
      .select("id,title,assignee,due_date,status")
      .neq("status", "done")
      .not("due_date", "is", null)
      .lte("due_date", today);

    if (error) throw error;

    const rows = data || [];
    const overdue = rows
      .filter((t) => t.due_date && t.due_date < today)
      .sort((a, b) => String(a.due_date).localeCompare(String(b.due_date)));
    const dueToday = rows
      .filter((t) => t.due_date === today)
      .sort((a, b) => String(a.title).localeCompare(String(b.title), "he"));

    if (!overdue.length && !dueToday.length) {
      return new Response(
        JSON.stringify({
          success: true,
          sent: false,
          overdue: 0,
          today: 0,
          message: "nothing to send",
        }),
        { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
      );
    }

    const subjectParts: string[] = [];
    if (overdue.length) subjectParts.push(`${overdue.length} באיחור`);
    if (dueToday.length) subjectParts.push(`${dueToday.length} להיום`);
    const subject = `📋 מנהלה — ${subjectParts.join(", ")}`;

    const li = (t: { title: string; assignee: string; due_date?: string }, extra?: string) =>
      `<li style="margin:0 0 8px 0;color:#0E3651;font-size:15px;line-height:1.45">` +
      `<strong>${esc(t.title)}</strong> — ${esc(ASSIGNEE_HE[t.assignee] || t.assignee)}` +
      (extra ? ` — <span style="color:#DC2626">${esc(extra)}</span>` : "") +
      `</li>`;

    let bodyHtml = `<p style="margin:0 0 20px;color:#0E3651;font-size:16px">בוקר טוב,</p>`;

    if (overdue.length) {
      bodyHtml += `<div style="margin:0 0 18px">
        <div style="font-weight:800;color:#BE185D;font-size:15px;margin-bottom:8px">🔴 באיחור</div>
        <ul style="margin:0;padding-right:18px">${
          overdue.map((t) =>
            li(t, `באיחור ${daysOverdue(String(t.due_date), today)} ימים`)
          ).join("")
        }</ul>
      </div>`;
    }

    if (dueToday.length) {
      bodyHtml += `<div style="margin:0 0 18px">
        <div style="font-weight:800;color:#0E3651;font-size:15px;margin-bottom:8px">📅 להיום</div>
        <ul style="margin:0;padding-right:18px">${dueToday.map((t) => li(t)).join("")}</ul>
      </div>`;
    }

    const html = `<!DOCTYPE html><html lang="he" dir="rtl"><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f1f5f9;font-family:Arial,Helvetica,sans-serif;direction:rtl">
  <div style="max-width:560px;margin:24px auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #e2e8f0">
    <div style="background:#0E3651;padding:20px 24px">
      <div style="color:#62C7C2;font-weight:800;font-size:18px">📋 מנהלה</div>
      <div style="color:#94a3b8;font-size:13px;margin-top:4px">PrintOS · תזכורת בוקר</div>
    </div>
    <div style="height:4px;background:#62C7C2"></div>
    <div style="padding:24px">${bodyHtml}
      <div style="margin-top:28px;text-align:center;background:#EDF8F7;border-radius:10px;padding:14px;color:#0E3651;font-weight:800">
        פתחו את PrintOS → טאב 📋 מנהלה
      </div>
    </div>
  </div>
</body></html>`;

    const { error: mailErr } = await supabase.functions.invoke("send-email", {
      body: {
        to: RECIPIENTS,
        subject,
        html,
        from: "PrintOS <orders@natalie-print.com>",
      },
    });
    if (mailErr) throw mailErr;

    return new Response(
      JSON.stringify({
        success: true,
        sent: true,
        overdue: overdue.length,
        today: dueToday.length,
      }),
      { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
    );
  } catch (err) {
    console.error("[admin-tasks-reminder]", err);
    return new Response(
      JSON.stringify({ success: false, error: String((err as Error)?.message || err) }),
      {
        status: 500,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
      },
    );
  }
});

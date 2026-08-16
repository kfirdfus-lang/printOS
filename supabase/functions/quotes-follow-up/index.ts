// Package H Stage B — daily follow-up for open quotes older than 14 days
// Schedule: cron 0 6 * * * (06:00 UTC ≈ 09:00 Israel standard)
// SQL: supabase/migrations/20260816210000_quotes_follow_up_cron.sql

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

    const cutoffMs = Date.now() - 14 * 864e5;
    const { data: quotes, error } = await supabase
      .from("quotes")
      .select(
        "id,title,bina_doc_id,bina_cust_name,sent_at,created_at,follow_up_sent_at,quote_status",
      )
      .eq("quote_status", "sent")
      .is("follow_up_sent_at", null);

    if (error) throw error;

    const list = (quotes || []).filter((q) => {
      const base = q.sent_at || q.created_at;
      if (!base) return false;
      return new Date(base).getTime() < cutoffMs;
    });

    if (!list.length) {
      return new Response(
        JSON.stringify({ success: true, count: 0, message: "no quotes" }),
        { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
      );
    }

    const rows = list
      .map((q) => {
        const base = q.sent_at || q.created_at;
        const days = base
          ? Math.floor((Date.now() - new Date(base).getTime()) / 864e5)
          : "?";
        const cust = String(q.bina_cust_name || "—").replace(/</g, "&lt;");
        const title = String(q.title || "—").replace(/</g, "&lt;");
        return `<tr><td>${cust}</td><td>${title}</td><td>#${
          q.bina_doc_id || "—"
        }</td><td>${days}</td></tr>`;
      })
      .join("");

    const html = `<!DOCTYPE html><html dir="rtl" lang="he"><body style="font-family:Arial,sans-serif;direction:rtl;padding:16px">
      <h2>הצעות פתוחות מעל 14 יום (${list.length})</h2>
      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%">
        <thead><tr><th>לקוח</th><th>כותרת</th><th>בינה</th><th>ימים</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </body></html>`;

    const { error: mailErr } = await supabase.functions.invoke("send-email", {
      body: {
        to: ["kfir.dfus@gmail.com", "natalie.zem@gmail.com"],
        subject: `הצעות פתוחות מעל 14 יום (${list.length})`,
        html,
        from: "PrintOS <orders@natalie-print.com>",
      },
    });
    if (mailErr) throw mailErr;

    const ids = list.map((q) => q.id);
    const { error: updErr } = await supabase
      .from("quotes")
      .update({ follow_up_sent_at: new Date().toISOString() })
      .in("id", ids);
    if (updErr) throw updErr;

    return new Response(
      JSON.stringify({ success: true, count: list.length, ids }),
      { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
    );
  } catch (err) {
    console.error(err);
    return new Response(
      JSON.stringify({
        success: false,
        error: err instanceof Error ? err.message : String(err),
      }),
      {
        status: 500,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
      },
    );
  }
});

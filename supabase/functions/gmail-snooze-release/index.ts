// G3 — daily cron: restore snoozed Gmail messages to INBOX.

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import { gmailModifyLabels, getValidAccessToken } from "../_shared/gmail.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-cron-secret",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function cronAuthorized(req: Request): boolean {
  const expected = Deno.env.get("CRON_SECRET")?.trim();
  if (!expected) {
    console.error("[gmail-snooze-release] CRON_SECRET env not set");
    return false;
  }
  return req.headers.get("x-cron-secret")?.trim() === expected;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
  if (!cronAuthorized(req)) {
    return new Response(JSON.stringify({ error: "Unauthorized" }), {
      status: 401,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const url = Deno.env.get("SUPABASE_URL")!;
  const key = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const sb = createClient(url, key);

  const { data: due, error } = await sb.from("gmail_snoozed")
    .select("id,message_id")
    .eq("released", false)
    .lte("snooze_until", new Date().toISOString())
    .limit(200);
  if (error) {
    return new Response(JSON.stringify({ error: error.message }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const { data: conn } = await sb.from("gmail_connections").select("user_id").limit(1).maybeSingle();
  if (!conn?.user_id) {
    return new Response(JSON.stringify({ released: 0, note: "no gmail connection" }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const tok = await getValidAccessToken(sb, conn.user_id);
  if ("error" in tok) {
    return new Response(JSON.stringify({ error: "token failed" }), {
      status: 401,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  let released = 0;
  for (const row of due || []) {
    const res = await gmailModifyLabels(tok.token, row.message_id, ["INBOX"], []);
    if (res.ok) {
      await sb.from("gmail_snoozed").update({ released: true }).eq("id", row.id);
      released++;
    }
  }

  return new Response(JSON.stringify({ released, total: (due || []).length }), {
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
});

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type, x-cron-secret",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const COMPLETED_STATUSES = ["הושלם", "הושלמה"];

function cronAuthorized(req: Request): boolean {
  const expected = Deno.env.get("CRON_SECRET")?.trim();
  if (!expected) {
    console.error("[archive-completed-tasks] CRON_SECRET env not set");
    return false;
  }
  const got = req.headers.get("x-cron-secret")?.trim();
  return got === expected;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed; use POST" }), {
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
  const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const supabase = createClient(url, serviceKey);

  const cutoffIso = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();

  const orClause = COMPLETED_STATUSES.map((s) => `status.eq.${s}`).join(",");

  const { data: updated, error } = await supabase
    .from("tasks")
    .update({ archived_at: new Date().toISOString() })
    .is("archived_at", null)
    .lt("updated_at", cutoffIso)
    .or(orClause)
    .select("id");

  if (error) {
    console.error("[archive-completed-tasks]", error);
    return new Response(JSON.stringify({ error: error.message }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const count = updated?.length ?? 0;

  console.log("[archive-completed-tasks] archived rows:", count);

  return new Response(
    JSON.stringify({ success: true, archivedCount: count }),
    {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    },
  );
});

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type, x-cron-secret",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const COMPLETED_STATUSES = ["הושלם", "הושלמה"];

/** משימות במשלוחים שעדיין לא נמסרו — לא לארכב */
const ACTIVE_DELIVERY_STATUSES = ["ready", "in_transit"];

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

  const { data: candidates, error: candidatesError } = await supabase
    .from("tasks")
    .select("id, delivery_status")
    .is("archived_at", null)
    .lt("updated_at", cutoffIso)
    .or(orClause);

  if (candidatesError) {
    console.error("[archive-completed-tasks] fetch candidates:", candidatesError);
    return new Response(JSON.stringify({ error: candidatesError.message }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const candidateCount = candidates?.length ?? 0;
  console.log(`[archive-completed-tasks] Found ${candidateCount} archive candidates`);

  if (!candidates || candidates.length === 0) {
    return new Response(
      JSON.stringify({ success: true, archivedCount: 0, skippedActiveDelivery: 0 }),
      { headers: { ...corsHeaders, "Content-Type": "application/json" } },
    );
  }

  const blockedTaskIds = new Set(
    candidates
      .filter((t) =>
        t.delivery_status &&
        ACTIVE_DELIVERY_STATUSES.includes(t.delivery_status)
      )
      .map((t) => t.id),
  );

  const toArchive = candidates.filter((t) => !blockedTaskIds.has(t.id));

  console.log(
    `[archive-completed-tasks] ${blockedTaskIds.size} tasks have active delivery (ready/in_transit) — skipping`,
  );
  console.log(`[archive-completed-tasks] Archiving ${toArchive.length} tasks`);

  if (toArchive.length === 0) {
    return new Response(
      JSON.stringify({
        success: true,
        archivedCount: 0,
        skippedActiveDelivery: blockedTaskIds.size,
      }),
      { headers: { ...corsHeaders, "Content-Type": "application/json" } },
    );
  }

  const archiveIds = toArchive.map((t) => t.id);
  const archivedAt = new Date().toISOString();

  const { error: archiveError } = await supabase
    .from("tasks")
    .update({ archived_at: archivedAt })
    .in("id", archiveIds);

  if (archiveError) {
    console.error("[archive-completed-tasks] archive update:", archiveError);
    return new Response(JSON.stringify({ error: archiveError.message }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  console.log(`[archive-completed-tasks] Successfully archived ${archiveIds.length} tasks`);

  return new Response(
    JSON.stringify({
      success: true,
      archivedCount: archiveIds.length,
      skippedActiveDelivery: blockedTaskIds.size,
    }),
    { headers: { ...corsHeaders, "Content-Type": "application/json" } },
  );
});

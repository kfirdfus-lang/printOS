const BINA_API_URL = "https://webfiles.binaw.com/post/PostJsonDocV2.aspx";
const TIMEOUT_MS = 30_000;

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function todayYmd(): string {
  return new Date().toISOString().split("T")[0];
}

function countOrdersFromText(text: string): number {
  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed?.Orders)) return parsed.Orders.length;
    if (Array.isArray(parsed)) return parsed.length;
    return 0;
  } catch {
    return 0;
  }
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

  const startedAt = Date.now();
  const env = {
    QUOTAGUARD_URL: Deno.env.get("QUOTAGUARD_URL") ?? "",
    BINA_TOKEN: Deno.env.get("BINA_TOKEN") ?? "",
  };

  if (!env.QUOTAGUARD_URL.trim()) {
    return new Response(
      JSON.stringify({
        success: false,
        duration_ms: Date.now() - startedAt,
        status: 500,
        ordersReturned: 0,
        error: "QUOTAGUARD_URL secret is not configured",
      }),
      { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } },
    );
  }

  if (!env.BINA_TOKEN.trim()) {
    return new Response(
      JSON.stringify({
        success: false,
        duration_ms: Date.now() - startedAt,
        status: 500,
        ordersReturned: 0,
        error: "BINA_TOKEN secret is not configured",
      }),
      { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } },
    );
  }

  let client: Deno.HttpClient | null = null;
  try {
    client = Deno.createHttpClient({
      proxy: {
        url: env.QUOTAGUARD_URL,
      },
    });

    const response = await fetch(BINA_API_URL, {
      client,
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        tokenId: env.BINA_TOKEN,
        docType: -15,
        fromDate: todayYmd(),
        toDate: todayYmd(),
      }),
      signal: AbortSignal.timeout(TIMEOUT_MS),
    });

    const rawText = await response.text();
    const durationMs = Date.now() - startedAt;

    return new Response(
      JSON.stringify({
        success: response.ok,
        duration_ms: durationMs,
        status: response.status,
        ordersReturned: response.ok ? countOrdersFromText(rawText) : 0,
        error: response.ok ? null : rawText.slice(0, 1200),
      }),
      {
        status: response.ok ? 200 : 502,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      },
    );
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    return new Response(
      JSON.stringify({
        success: false,
        duration_ms: Date.now() - startedAt,
        status: 500,
        ordersReturned: 0,
        error: msg,
      }),
      {
        status: 500,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      },
    );
  } finally {
    client?.close();
  }
});

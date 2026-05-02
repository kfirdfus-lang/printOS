const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
};

const BINA_ROOT_URL = "https://webfiles.binaw.com/";
const FETCH_TIMEOUT_MS = 15_000;

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  if (req.method !== "GET") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort("timeout"), FETCH_TIMEOUT_MS);

  try {
    // Any HTTP response (200, 404, 500, …) means the host is reachable; fetch only throws on network/DNS/timeout.
    await fetch(BINA_ROOT_URL, {
      method: "GET",
      signal: controller.signal,
      redirect: "follow",
    });

    return new Response(JSON.stringify({ status: "online" }), {
      status: 200,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (error) {
    return new Response(
      JSON.stringify({
        status: "offline",
        error: error instanceof Error ? error.message : String(error),
      }),
      {
        status: 200,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      },
    );
  } finally {
    clearTimeout(timeoutId);
  }
});

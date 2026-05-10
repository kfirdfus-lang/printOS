import axios from "npm:axios@1.7.9";
import { HttpsProxyAgent } from "npm:https-proxy-agent@7.0.5";
import { fetch as undiciFetch, ProxyAgent } from "npm:undici@6.19.8";

const BINA_API_URL = "https://webfiles.binaw.com/post/PostJsonDocV2.aspx";
const TIMEOUT_MS = 30_000;

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

type AttemptResult = {
  duration: number;
  status: "ok" | "failed";
  ordersReturned: number;
  error: string | null;
};

function todayYmd(): string {
  return new Date().toISOString().split("T")[0];
}

function countOrders(raw: string): number {
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed?.Orders)) return parsed.Orders.length;
    if (Array.isArray(parsed)) return parsed.length;
    return 0;
  } catch {
    return 0;
  }
}

function formatErr(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

async function axiosAttempt(proxyUrl: string, body: Record<string, unknown>): Promise<AttemptResult> {
  const started = Date.now();
  try {
    const agent = new HttpsProxyAgent(proxyUrl);
    const res = await axios.request<string>({
      method: "POST",
      url: BINA_API_URL,
      data: body,
      headers: { "Content-Type": "application/json" },
      httpsAgent: agent,
      proxy: false,
      timeout: TIMEOUT_MS,
      responseType: "text",
      transformResponse: [(d) => d],
      validateStatus: () => true,
    });
    const text = typeof res.data === "string" ? res.data : String(res.data);
    if (res.status < 200 || res.status >= 300) {
      throw new Error(`HTTP ${res.status}: ${text.slice(0, 300)}`);
    }
    return { duration: Date.now() - started, status: "ok", ordersReturned: countOrders(text), error: null };
  } catch (e) {
    return {
      duration: Date.now() - started,
      status: "failed",
      ordersReturned: 0,
      error: formatErr(e),
    };
  }
}

async function denoNativeAttempt(proxyUrl: string, body: Record<string, unknown>): Promise<AttemptResult> {
  const started = Date.now();
  try {
    if (typeof Deno.createHttpClient !== "function") {
      throw new Error("Deno.createHttpClient is not available in this runtime");
    }

    const client = Deno.createHttpClient({
      proxy: { url: proxyUrl },
    });

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort("timeout"), TIMEOUT_MS);
    try {
      const res = await fetch(BINA_API_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
        signal: controller.signal,
        client,
      });
      const text = await res.text();
      if (!res.ok) {
        throw new Error(`HTTP ${res.status}: ${text.slice(0, 300)}`);
      }
      return { duration: Date.now() - started, status: "ok", ordersReturned: countOrders(text), error: null };
    } finally {
      clearTimeout(timeoutId);
      client.close();
    }
  } catch (e) {
    return {
      duration: Date.now() - started,
      status: "failed",
      ordersReturned: 0,
      error: formatErr(e),
    };
  }
}

async function undiciAttempt(proxyUrl: string, body: Record<string, unknown>): Promise<AttemptResult> {
  const started = Date.now();
  const dispatcher = new ProxyAgent(proxyUrl);
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort("timeout"), TIMEOUT_MS);

  try {
    const res = await undiciFetch(BINA_API_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      dispatcher,
      signal: controller.signal,
    });
    const text = await res.text();
    if (!res.ok) {
      throw new Error(`HTTP ${res.status}: ${text.slice(0, 300)}`);
    }
    return { duration: Date.now() - started, status: "ok", ordersReturned: countOrders(text), error: null };
  } catch (e) {
    return {
      duration: Date.now() - started,
      status: "failed",
      ordersReturned: 0,
      error: formatErr(e),
    };
  } finally {
    clearTimeout(timeoutId);
    try {
      await dispatcher.close();
    } catch {
      // ignore close errors in diagnostics
    }
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

  const token = Deno.env.get("BINA_TOKEN");
  const proxyUrl = Deno.env.get("QUOTAGUARD_URL");
  if (!token?.trim()) {
    return new Response(JSON.stringify({ error: "BINA_TOKEN secret is not configured" }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
  if (!proxyUrl?.trim()) {
    return new Response(JSON.stringify({ error: "QUOTAGUARD_URL secret is not configured" }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const today = todayYmd();
  const body = {
    tokenId: token,
    docType: -15,
    fromDate: today,
    toDate: today,
  };

  const [axiosRes, denoRes, undiciRes] = await Promise.all([
    axiosAttempt(proxyUrl, body),
    denoNativeAttempt(proxyUrl, body),
    undiciAttempt(proxyUrl, body),
  ]);

  return new Response(
    JSON.stringify({
      axios: axiosRes,
      deno_native: denoRes,
      undici: undiciRes,
    }),
    { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } },
  );
});

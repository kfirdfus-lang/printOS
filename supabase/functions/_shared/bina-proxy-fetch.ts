import axios, { type AxiosRequestConfig } from "npm:axios@1.7.9";
import { HttpsProxyAgent } from "npm:https-proxy-agent@7.0.5";

const LEGACY_MAX_ATTEMPTS = 3;
const LEGACY_TIMEOUT_MS = 180_000;
/** After attempt 1 fails → wait before attempt 2. After attempt 2 fails → wait before attempt 3. */
const LEGACY_RETRY_DELAY_MS: [number, number] = [5000, 15_000];

const CLOUDFLARE_BINA_PROXY_DEFAULT = "https://bina-proxy.kfir-dfus.workers.dev";
const WORKER_TIMEOUT_MS = 30_000;
/** First try + one retry after failure */
const WORKER_MAX_ATTEMPTS = 2;
const WORKER_RETRY_DELAY_MS = 3000;

function formatAxiosError(e: unknown): string {
  if (e instanceof Error) return e.message;
  return String(e);
}

function bodyHasNonemptyTokenId(jsonBody: unknown): boolean {
  if (!jsonBody || typeof jsonBody !== "object") return false;
  const tid = (jsonBody as Record<string, unknown>)["tokenId"];
  return typeof tid === "string" && tid.trim().length > 0;
}

/**
 * Routes to Cloudflare Worker (V2 only, token added there) when the declared target is PostJsonDocV2
 * and the caller omitted tokenId (e.g. sync-bina-orders, bina-create-customer).
 * Callers that still send tokenId (e.g. bina-api-test) or use PostJsonDoc.aspx (bina-create-quote) stay on QuotaGuard.
 */
function shouldRouteViaCloudflareWorker(binaUrl: string, jsonBody: unknown): boolean {
  let parsed: URL;
  try {
    parsed = new URL(binaUrl);
  } catch {
    return false;
  }
  if (parsed.hostname !== "webfiles.binaw.com") return false;
  if (!parsed.pathname.endsWith("PostJsonDocV2.aspx")) {
    return false;
  }
  if (bodyHasNonemptyTokenId(jsonBody)) return false;
  return true;
}

async function fetchBinaViaCloudflareWorker(
  jsonBody: unknown,
): Promise<{ ok: boolean; status: number; text: string }> {
  const workerBase = Deno.env.get("BINA_PROXY_WORKER_URL")?.trim() || CLOUDFLARE_BINA_PROXY_DEFAULT;

  console.log("[bina] routing via Cloudflare Worker", {
    workerBase,
    attempts: WORKER_MAX_ATTEMPTS,
    timeoutMs: WORKER_TIMEOUT_MS,
  });

  let lastErr: unknown;

  for (let attempt = 1; attempt <= WORKER_MAX_ATTEMPTS; attempt++) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(new Error(`Worker request timeout (${WORKER_TIMEOUT_MS} ms)`)), WORKER_TIMEOUT_MS);

    try {
      const res = await fetch(workerBase, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(jsonBody),
        signal: controller.signal,
      });
      clearTimeout(timeoutId);

      const text = await res.text();
      const ok = res.ok;
      console.log("[bina] Cloudflare Worker response", { attempt, httpStatus: res.status, ok });

      return { ok, status: res.status, text };
    } catch (e) {
      clearTimeout(timeoutId);
      lastErr = e;
      const detail = formatAxiosError(e);
      console.error("[bina] Cloudflare Worker request failed", { attempt, detail });

      if (attempt >= WORKER_MAX_ATTEMPTS) {
        throw e;
      }
      console.log("[bina] retry after", WORKER_RETRY_DELAY_MS, "ms");
      await new Promise((r) => setTimeout(r, WORKER_RETRY_DELAY_MS));
    }
  }

  throw lastErr instanceof Error ? lastErr : new Error("fetchBinaViaCloudflareWorker: unexpected state");
}

export type FetchBinaViaQuotaGuardOptions = {
  /** Per-request timeout (ms). Legacy (QuotaGuard): defaults to LEGACY_TIMEOUT_MS. Ignored when using Worker (fixed 30s). */
  timeoutMs?: number;
  /** Legacy (QuotaGuard): max attempts. Ignored when using Worker (fixed 2 attempts = 1 retry). */
  maxAttempts?: number;
};

/** POST JSON to Bina (webfiles.binaw.com) — via Cloudflare Worker for V2 without tokenId, else via QuotaGuard + axios. */
export async function fetchBinaViaQuotaGuard(
  binaUrl: string,
  jsonBody: unknown,
  options?: FetchBinaViaQuotaGuardOptions,
): Promise<{ ok: boolean; status: number; text: string }> {
  const host = new URL(binaUrl).hostname;
  if (host !== "webfiles.binaw.com") {
    throw new Error(`fetchBinaViaQuotaGuard: unexpected host (${host}), expected webfiles.binaw.com`);
  }

  if (shouldRouteViaCloudflareWorker(binaUrl, jsonBody)) {
    return await fetchBinaViaCloudflareWorker(jsonBody);
  }

  const proxyRaw = Deno.env.get("QUOTAGUARD_URL");
  if (!proxyRaw?.trim()) {
    console.error("[bina] QUOTAGUARD_URL is missing — Bina calls require QuotaGuard");
    throw new Error("QUOTAGUARD_URL secret is not configured");
  }

  let proxyParsed: URL;
  try {
    proxyParsed = new URL(proxyRaw.trim());
  } catch {
    throw new Error("QUOTAGUARD_URL is not a valid URL");
  }

  const portNum = proxyParsed.port
    ? parseInt(proxyParsed.port, 10)
    : (proxyParsed.protocol === "https:" ? 443 : 80);

  const httpsAgent = new HttpsProxyAgent(proxyRaw.trim());

  console.log("[bina] routing through QuotaGuard (CONNECT via HttpsProxyAgent)", {
    proxyHost: proxyParsed.hostname,
    proxyPort: portNum,
    target: host,
    path: new URL(binaUrl).pathname,
  });

  const timeoutMs = options?.timeoutMs ?? LEGACY_TIMEOUT_MS;
  const attempts = Math.max(1, Math.min(options?.maxAttempts ?? LEGACY_MAX_ATTEMPTS, LEGACY_MAX_ATTEMPTS));

  const config: AxiosRequestConfig = {
    method: "POST",
    url: binaUrl,
    headers: { "Content-Type": "application/json" },
    data: jsonBody,
    httpsAgent,
    proxy: false,
    responseType: "text",
    transformResponse: [(data) => data],
    validateStatus: () => true,
    timeout: timeoutMs,
    maxBodyLength: Infinity,
    maxContentLength: Infinity,
  };

  let res: Awaited<ReturnType<typeof axios.request<string>>> | undefined;

  for (let attempt = 1; attempt <= attempts; attempt++) {
    try {
      res = await axios.request<string>(config);
      break;
    } catch (e) {
      const errorMessage = formatAxiosError(e);
      console.error("[bina] proxy request failed (network)", {
        attempt,
        proxyHost: proxyParsed.hostname,
        detail: errorMessage,
      });
      if (attempt === attempts) {
        throw e;
      }
      const nextAttempt = attempt + 1;
      console.log("[bina] retry attempt", nextAttempt, "of", attempts, "after error:", errorMessage);
      const waitMs = attempt === 1 ? LEGACY_RETRY_DELAY_MS[0] : LEGACY_RETRY_DELAY_MS[1];
      await new Promise((r) => setTimeout(r, waitMs));
    }
  }

  if (!res) {
    throw new Error("fetchBinaViaQuotaGuard: no response after retries");
  }

  const ok = res.status >= 200 && res.status < 300;
  console.log("[bina] CONNECT/proxy request finished", {
    httpStatus: res.status,
    ok,
    viaProxyHost: proxyParsed.hostname,
  });

  return { ok, status: res.status, text: typeof res.data === "string" ? res.data : String(res.data) };
}

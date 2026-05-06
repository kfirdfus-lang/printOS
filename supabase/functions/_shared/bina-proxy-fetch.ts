const TIMEOUT_MS = 30_000;
const MAX_ATTEMPTS = 2;
const RETRY_DELAY_MS = 3000;

function formatError(e: unknown): string {
  if (e instanceof Error) return e.message;
  return String(e);
}

export type FetchBinaViaQuotaGuardOptions = {
  /** Per-request timeout (ms). Defaults to 30 seconds. */
  timeoutMs?: number;
  /** Maximum attempts including first try. Defaults to 2 (one retry). */
  maxAttempts?: number;
};

const ALLOWED_BINA_HOSTS = new Set(["webfiles.binaw.com", "webapps.binaw.com"]);

/** POST JSON to Bina (webfiles / webapps) via QuotaGuard using Deno native HTTP client proxy support. */
export async function fetchBinaViaQuotaGuard(
  binaUrl: string,
  jsonBody: unknown,
  options?: FetchBinaViaQuotaGuardOptions,
): Promise<{ ok: boolean; status: number; text: string }> {
  const host = new URL(binaUrl).hostname;
  if (!ALLOWED_BINA_HOSTS.has(host)) {
    throw new Error(
      `fetchBinaViaQuotaGuard: unexpected host (${host}), expected one of: ${[...ALLOWED_BINA_HOSTS].join(", ")}`,
    );
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

  console.log("[bina] routing through QuotaGuard (Deno.createHttpClient)", {
    proxyHost: proxyParsed.hostname,
    proxyPort: proxyParsed.port || (proxyParsed.protocol === "https:" ? "443" : "80"),
    target: host,
    path: new URL(binaUrl).pathname,
  });

  const timeoutMs = options?.timeoutMs ?? TIMEOUT_MS;
  const attempts = Math.max(1, Math.min(options?.maxAttempts ?? MAX_ATTEMPTS, MAX_ATTEMPTS));
  let lastErr: unknown;

  for (let attempt = 1; attempt <= attempts; attempt++) {
    const client = Deno.createHttpClient({
      proxy: { url: proxyRaw.trim() },
    });
    try {
      // Entire jsonBody is serialized as-is — no field filtering / token rewriting.
      const serializedBody = JSON.stringify(jsonBody);
      console.log("[bina] fetch body verbatim", {
        attempt,
        targetUrl: binaUrl,
        serializedCharLength: serializedBody.length,
        serializedByteLength: new TextEncoder().encode(serializedBody).length,
      });

      const res = await fetch(binaUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: serializedBody,
        signal: AbortSignal.timeout(timeoutMs),
        client,
      });
      const text = await res.text();
      const ok = res.status >= 200 && res.status < 300;
      console.log("[bina] request finished", {
        attempt,
        httpStatus: res.status,
        ok,
        viaProxyHost: proxyParsed.hostname,
      });
      return { ok, status: res.status, text };
    } catch (e) {
      lastErr = e;
      const errorMessage = formatError(e);
      console.error("[bina] proxy request failed", {
        attempt,
        proxyHost: proxyParsed.hostname,
        detail: errorMessage,
      });
      if (attempt === attempts) {
        throw e;
      }
      console.log("[bina] retry after", RETRY_DELAY_MS, "ms");
      await new Promise((r) => setTimeout(r, RETRY_DELAY_MS));
    } finally {
      client.close();
    }
  }

  throw lastErr instanceof Error ? lastErr : new Error("fetchBinaViaQuotaGuard: no response after retries");
}

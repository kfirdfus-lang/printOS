import axios, { type AxiosRequestConfig } from "npm:axios@1.7.9";
import { HttpsProxyAgent } from "npm:https-proxy-agent@7.0.5";

/** POST JSON to Bina (webfiles.binaw.com only) via QuotaGuard HTTP proxy using CONNECT tunneling. */
export async function fetchBinaViaQuotaGuard(
  binaUrl: string,
  jsonBody: unknown,
): Promise<{ ok: boolean; status: number; text: string }> {
  const host = new URL(binaUrl).hostname;
  if (host !== "webfiles.binaw.com") {
    throw new Error(`fetchBinaViaQuotaGuard: unexpected host (${host}), expected webfiles.binaw.com`);
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

  // Axios `proxy: { ... }` does not reliably issue HTTP CONNECT for HTTPS targets on Edge/Deno.
  // HttpsProxyAgent performs CONNECT tunneling to the HTTPS origin.
  const httpsAgent = new HttpsProxyAgent(proxyRaw.trim());

  console.log("[bina] routing through QuotaGuard (CONNECT via HttpsProxyAgent)", {
    proxyHost: proxyParsed.hostname,
    proxyPort: portNum,
    target: host,
    path: new URL(binaUrl).pathname,
  });

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
    timeout: 120_000,
    maxBodyLength: Infinity,
    maxContentLength: Infinity,
  };

  let res;
  try {
    res = await axios.request<string>(config);
  } catch (e) {
    console.error("[bina] proxy request failed (network)", {
      proxyHost: proxyParsed.hostname,
      detail: e instanceof Error ? e.message : String(e),
    });
    throw e;
  }

  const ok = res.status >= 200 && res.status < 300;
  console.log("[bina] CONNECT/proxy request finished", {
    httpStatus: res.status,
    ok,
    viaProxyHost: proxyParsed.hostname,
  });

  return { ok, status: res.status, text: typeof res.data === "string" ? res.data : String(res.data) };
}

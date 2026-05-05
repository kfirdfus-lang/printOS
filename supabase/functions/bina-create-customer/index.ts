import { fetchBinaViaQuotaGuard } from "../_shared/bina-proxy-fetch.ts";

const BINA_API_URL = "https://webfiles.binaw.com/post/PostJsonDocV2.aspx";
const DOC_TYPE = -26;
const TIMEOUT_MS = 45_000;
const MAX_ATTEMPTS = 2;

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

const REQUIRED_FIELDS = [
  "custName",
  "custTz",
  "custCity",
  "custAddress",
  "custGroup",
  "custTel",
  "custCel",
  "custEmail",
  "termsofPayment",
] as const;

type CustomerInput = Record<(typeof REQUIRED_FIELDS)[number], string>;

function isNonEmptyString(v: unknown): v is string {
  return typeof v === "string" && v.trim().length > 0;
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

  let body: Record<string, unknown>;
  try {
    body = await req.json();
  } catch {
    return new Response(JSON.stringify({ error: "Invalid JSON body" }), {
      status: 400,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const missing: string[] = [];
  for (const key of REQUIRED_FIELDS) {
    if (!isNonEmptyString(body[key])) missing.push(key);
  }
  if (missing.length > 0) {
    return new Response(
      JSON.stringify({
        error: "Missing or empty required field(s)",
        missing,
        hint: "All of the following must be non-empty strings: " + REQUIRED_FIELDS.join(", "),
      }),
      { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } },
    );
  }

  const binaToken = Deno.env.get("BINA_TOKEN");
  if (!binaToken?.trim()) {
    return new Response(JSON.stringify({ error: "BINA_TOKEN secret is not configured" }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const input = body as unknown as CustomerInput;
  const binaRequest = {
    tokenId: binaToken,
    docType: DOC_TYPE,
    custName: input.custName.trim(),
    custCity: input.custCity.trim(),
    custAddress: input.custAddress.trim(),
    custGroup: input.custGroup.trim(),
    custTel: input.custTel.trim(),
    custCel: input.custCel.trim(),
    custEmail: input.custEmail.trim(),
    termsofPayment: input.termsofPayment.trim(),
    custTz: input.custTz.trim(),
  };

  const displayName = binaRequest.custName.slice(0, 120);
  console.log("[bina-create-customer] Creating customer:", displayName);

  try {
    const r = await fetchBinaViaQuotaGuard(BINA_API_URL, binaRequest, {
      timeoutMs: TIMEOUT_MS,
      maxAttempts: MAX_ATTEMPTS,
    });
    console.log("[bina-create-customer] Bina response:", r.status);

    if (!r.ok) {
      return new Response(
        JSON.stringify({
          error: "Bina returned a non-success HTTP status",
          httpStatus: r.status,
          details: r.text.slice(0, 4000),
        }),
        { status: 502, headers: { ...corsHeaders, "Content-Type": "application/json" } },
      );
    }

    let parsed: unknown;
    try {
      parsed = JSON.parse(r.text);
    } catch {
      return new Response(
        JSON.stringify({
          error: "Bina returned a response that is not valid JSON",
          raw: r.text.slice(0, 4000),
        }),
        { status: 502, headers: { ...corsHeaders, "Content-Type": "application/json" } },
      );
    }

    return new Response(JSON.stringify(parsed), {
      status: 200,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error("[bina-create-customer] Request failed:", msg);
    return new Response(
      JSON.stringify({
        error: "Failed to reach Bina via proxy",
        detail: msg,
      }),
      { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } },
    );
  }
});

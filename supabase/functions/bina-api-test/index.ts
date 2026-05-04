import { fetchBinaViaQuotaGuard } from "../_shared/bina-proxy-fetch.ts";

const BINA_API_URL = "https://webfiles.binaw.com/post/PostJsonDocV2.aspx";
const DOC_TYPE = -15;
const PER_REQUEST_TIMEOUT_MS = 30_000;
const DELAY_BETWEEN_MS = 2000;

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function dateYMD(d: Date): string {
  return d.toISOString().split("T")[0];
}

function yesterdayDate(): Date {
  const d = new Date();
  d.setUTCDate(d.getUTCDate() - 1);
  return d;
}

function tomorrowDate(): Date {
  const d = new Date();
  d.setUTCDate(d.getUTCDate() + 1);
  return d;
}

/** YYYY-MM-DDTHH:mm:ss (no timezone suffix) from UTC midnight yesterday */
function yesterdayMidnightIso(): string {
  const d = yesterdayDate();
  d.setUTCHours(0, 0, 0, 0);
  return d.toISOString().slice(0, 19);
}

function nowIsoUtc(): string {
  return new Date().toISOString().slice(0, 19);
}

/** DD/MM/YYYY (UTC calendar components) */
function toIsraeliDMY(d: Date): string {
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const yyyy = d.getUTCFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function countOrdersFromResponseText(text: string): number {
  try {
    const binaData = JSON.parse(text);
    if (binaData.Orders && Array.isArray(binaData.Orders)) return binaData.Orders.length;
    if (Array.isArray(binaData)) return binaData.length;
    return 0;
  } catch {
    return 0;
  }
}

function redactParams(params: Record<string, unknown>): Record<string, unknown> {
  const out = { ...params };
  if ("tokenId" in out) out.tokenId = "***";
  return out;
}

type TestResult = {
  name: string;
  params: Record<string, unknown>;
  status: "success" | "failed";
  ordersReturned: number;
  responseSnippet: string;
  error: string | null;
};

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

  const binaToken = Deno.env.get("BINA_TOKEN");
  if (!binaToken?.trim()) {
    return new Response(JSON.stringify({ error: "BINA_TOKEN secret is not configured" }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const y = dateYMD(yesterdayDate());
  const t = dateYMD(tomorrowDate());
  const createdAfter = yesterdayMidnightIso();
  const createdSince = y;
  const fromDateCreated = y;
  const toDateCreated = nowIsoUtc();
  const modifiedAfter = yesterdayMidnightIso();
  const createDateFrom = toIsraeliDMY(yesterdayDate());
  const createDateTo = toIsraeliDMY(tomorrowDate());

  const testDefinitions: { name: string; params: Record<string, unknown> }[] = [
    {
      name: "Test 1: Standard fromDate/toDate",
      params: { fromDate: y, toDate: t },
    },
    {
      name: "Test 2: createdAfter (ISO)",
      params: { createdAfter },
    },
    {
      name: "Test 3: createdSince",
      params: { createdSince },
    },
    {
      name: "Test 4: fromDateCreated / toDateCreated",
      params: { fromDateCreated, toDateCreated },
    },
    {
      name: "Test 5: modifiedAfter",
      params: { modifiedAfter },
    },
    {
      name: "Test 6: createDateFrom / createDateTo (DD/MM/YYYY)",
      params: { createDateFrom, createDateTo },
    },
  ];

  const tests: TestResult[] = [];

  for (let i = 0; i < testDefinitions.length; i++) {
    if (i > 0) {
      await new Promise((r) => setTimeout(r, DELAY_BETWEEN_MS));
    }

    const def = testDefinitions[i];
    const body: Record<string, unknown> = {
      tokenId: binaToken,
      docType: DOC_TYPE,
      ...def.params,
    };

    try {
      const r = await fetchBinaViaQuotaGuard(BINA_API_URL, body, {
        timeoutMs: PER_REQUEST_TIMEOUT_MS,
      });
      const text = r.text;
      tests.push({
        name: def.name,
        params: redactParams(body),
        status: "success",
        ordersReturned: countOrdersFromResponseText(text),
        responseSnippet: text.slice(0, 200),
        error: null,
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      tests.push({
        name: def.name,
        params: redactParams(body),
        status: "failed",
        ordersReturned: 0,
        responseSnippet: "",
        error: msg,
      });
    }
  }

  const succeeded = tests.filter((x) => x.status === "success").length;
  const payload = {
    summary: `${succeeded} out of 6 requests succeeded`,
    tests,
  };

  return new Response(JSON.stringify(payload), {
    status: 200,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
});

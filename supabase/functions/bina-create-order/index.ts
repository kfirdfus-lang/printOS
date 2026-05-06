import { fetchBinaViaQuotaGuard } from "../_shared/bina-proxy-fetch.ts";

const BINA_ORDER_URL = "https://webapps.binaw.com/PostJsonDoc.aspx";
const DOC_TYPE = 15;

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

type BodyItem = {
  itemId?: string;
  description: string;
  quantity: number;
  unitPrice: number;
  discount?: number;
};

type BodyClient = {
  binaCustomerId: string;
  name: string;
  city: string;
  address: string;
  contactPerson: string;
  phone?: string;
  email?: string;
};

function isNonEmptyString(v: unknown): v is string {
  return typeof v === "string" && v.trim().length > 0;
}

function validateBody(payload: Record<string, unknown>): string | null {
  const client = payload.client;
  if (!client || typeof client !== "object") return "Missing or invalid client object";
  const c = client as Record<string, unknown>;

  if (!isNonEmptyString(c.binaCustomerId)) return "client.binaCustomerId is required";
  if (!isNonEmptyString(c.name)) return "client.name is required";
  if (!isNonEmptyString(c.city)) return "client.city is required";
  if (!isNonEmptyString(c.address)) return "client.address is required";
  if (!isNonEmptyString(c.contactPerson)) return "client.contactPerson is required";

  const items = payload.items;
  if (!Array.isArray(items) || items.length === 0) {
    return "items must be a non-empty array";
  }

  for (let i = 0; i < items.length; i++) {
    const raw = items[i];
    if (!raw || typeof raw !== "object") return `items[${i}] must be an object`;
    const it = raw as Record<string, unknown>;
    if (!isNonEmptyString(it.description)) return `items[${i}].description is required`;
    if (typeof it.quantity !== "number" || !Number.isFinite(it.quantity) || it.quantity <= 0) {
      return `items[${i}].quantity must be a positive number`;
    }
    if (typeof it.unitPrice !== "number" || !Number.isFinite(it.unitPrice) || it.unitPrice < 0) {
      return `items[${i}].unitPrice must be a non-negative number`;
    }
  }

  return null;
}

function parseCustId(binaCustomerId: string): number | null {
  const n = Number(String(binaCustomerId).trim());
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : null;
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
    return new Response(JSON.stringify({ success: false, error: "Invalid JSON body" }), {
      status: 400,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const errMsg = validateBody(body);
  if (errMsg) {
    console.error("[bina-create-order] validation failed:", errMsg);
    return new Response(JSON.stringify({ success: false, error: errMsg }), {
      status: 400,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const binaToken = Deno.env.get("BINA_TOKEN");
  if (!binaToken?.trim()) {
    return new Response(JSON.stringify({ success: false, error: "BINA_TOKEN secret is not configured" }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const client = body.client as BodyClient;
  const items = body.items as BodyItem[];
  const titleOpt = typeof body.title === "string" ? body.title.trim() : "";
  const remarkOpt = typeof body.remark === "string" ? body.remark.trim() : "";

  const custId = parseCustId(client.binaCustomerId);
  if (custId == null) {
    return new Response(
      JSON.stringify({ success: false, error: "client.binaCustomerId must be a positive numeric Bina customer id" }),
      { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } },
    );
  }

  const requestId = Date.now();
  const docTitle = titleOpt || "הזמנה מהפרינטוס";
  const docRemark = remarkOpt;
  const docStatus = "חדש";

  const binaRequest = {
    tokenId: binaToken,
    requestId,
    docType: DOC_TYPE,
    docWithvat: 1,
    docTitle,
    docStatus,
    docRemark,
    Cust: {
      custId,
      custName: client.name.trim().substring(0, 80),
      custCity: client.city.trim().substring(0, 50),
      custAddress: client.address.trim().substring(0, 120),
      custIshKheser: client.contactPerson.trim().substring(0, 80),
      custTel: (client.phone ?? "").trim().substring(0, 30),
      custEmail: (client.email ?? "").trim().substring(0, 80),
    },
    docItems: items.map((it, idx) => ({
      ItemId: (it.itemId && String(it.itemId).trim()) ? String(it.itemId).trim().substring(0, 40) : `item-${idx + 1}`,
      ItemDesc: it.description.trim().substring(0, 200),
      ItemQty: Math.floor(it.quantity),
      UnitPrice: String(it.unitPrice),
      Unitcurrency: "ILS",
      CurValue: 1,
      Discount: typeof it.discount === "number" && Number.isFinite(it.discount) ? it.discount : 0,
    })),
  };

  const tokenTrimmed = binaToken.trim();
  const payloadJson = JSON.stringify(binaRequest);
  const payloadBytes = new TextEncoder().encode(payloadJson).length;

  console.log("[bina-create-order] pre-send diagnostics", {
    BINA_TOKEN_defined: typeof binaToken !== "undefined" && binaToken !== null,
    BINA_TOKEN_typeof: typeof binaToken,
    BINA_TOKEN_trimmed_length: tokenTrimmed.length,
    BINA_TOKEN_first_5_chars: tokenTrimmed.slice(0, 5),
    target_url_full: BINA_ORDER_URL,
    body_json_string_length: payloadJson.length,
    body_utf8_byte_length: payloadBytes,
    requestId,
    custId,
    docTitle,
    item_count: binaRequest.docItems.length,
  });

  console.log("[bina-create-order] dispatching to Bina (subset, no secrets)", {
    requestId,
    custId,
    docTitle,
    itemCount: binaRequest.docItems.length,
  });

  try {
    const r = await fetchBinaViaQuotaGuard(BINA_ORDER_URL, binaRequest);
    console.log("[bina-create-order] HTTP", r.status, "len", r.text.length);

    if (!r.ok) {
      console.error("[bina-create-order] non-2xx body snippet:", r.text.slice(0, 800));
      return new Response(
        JSON.stringify({
          success: false,
          error: `Bina HTTP ${r.status}`,
          binaResponse: r.text,
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
          success: false,
          error: "Bina returned non-JSON",
          binaResponse: r.text,
        }),
        { status: 502, headers: { ...corsHeaders, "Content-Type": "application/json" } },
      );
    }

    const responseObj = Array.isArray(parsed) ? parsed[0] : parsed;
    const ro = responseObj as Record<string, unknown> | null;

    if (!ro || ro.ResCode === undefined) {
      console.error("[bina-create-order] unexpected shape:", JSON.stringify(parsed).slice(0, 500));
      return new Response(
        JSON.stringify({
          success: false,
          error: "Unexpected Bina response shape",
          binaResponse: parsed,
        }),
        { status: 502, headers: { ...corsHeaders, "Content-Type": "application/json" } },
      );
    }

    const resCode = ro.ResCode;
    const okCode = resCode === 0 || resCode === "0";

    if (!okCode) {
      const msg = String(ro.ResMsg ?? ro.ResMsgHe ?? "Unknown error");
      console.error("[bina-create-order] ResCode", resCode, "ResMsg", msg);
      return new Response(
        JSON.stringify({
          success: false,
          error: msg,
          binaResponse: parsed,
        }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } },
      );
    }

    const docId = ro.docId ?? ro.DocId;
    const binaOrderId = docId != null ? docId : null;
    console.log("[bina-create-order] success docId", binaOrderId);

    return new Response(
      JSON.stringify({
        success: true,
        binaOrderId,
        message: "הזמנה נוצרה",
      }),
      { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } },
    );
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error("[bina-create-order] thrown:", msg);
    return new Response(
      JSON.stringify({ success: false, error: msg }),
      { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } },
    );
  }
});

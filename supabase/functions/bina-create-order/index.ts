import { fetchBinaViaQuotaGuard } from "../_shared/bina-proxy-fetch.ts";

const BINA_ORDER_URL = "https://webfiles.binaw.com/post/PostJsonDoc.aspx";
const DOC_TYPE = 15;
const DOC_TITLE_DEFAULT = "הזמנה מהפרינטוס";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

type IncomingItem = {
  itemId?: string;
  description: string;
  quantity: number;
  unitPrice: number;
  discount?: number;
};

type IncomingClient = {
  binaCustomerId: string | number;
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

function parseCustId(raw: unknown): number | null {
  if (typeof raw === "number" && Number.isFinite(raw) && raw > 0) {
    return Math.floor(raw);
  }
  if (typeof raw === "string") {
    const t = raw.trim();
    if (!t) return null;
    const n = Number(t);
    return Number.isFinite(n) && n > 0 ? Math.floor(n) : null;
  }
  return null;
}

function jsonErr(status: number, message: string) {
  return new Response(JSON.stringify({ success: false, error: message }), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}

function validate(body: Record<string, unknown>): string | null {
  const client = body.client;
  if (!client || typeof client !== "object") return "client is required";
  const c = client as Record<string, unknown>;

  const custId = parseCustId(c.binaCustomerId);
  if (custId === null) return "client.binaCustomerId must be a positive integer";

  if (!isNonEmptyString(c.name)) return "client.name is required";
  if (!isNonEmptyString(c.city)) return "client.city is required";
  if (!isNonEmptyString(c.address)) return "client.address is required";
  if (!isNonEmptyString(c.contactPerson)) return "client.contactPerson is required";

  const items = body.items;
  if (!Array.isArray(items) || items.length === 0) return "items must contain at least one item";

  for (let i = 0; i < items.length; i++) {
    const raw = items[i];
    if (!raw || typeof raw !== "object") return `items[${i}] is invalid`;
    const it = raw as Record<string, unknown>;
    if (!isNonEmptyString(it.description)) return `items[${i}].description is required`;

    const qty = it.quantity;
    if (typeof qty !== "number" || !Number.isInteger(qty) || qty <= 0) {
      return `items[${i}].quantity must be a positive integer`;
    }

    const up = it.unitPrice;
    if (typeof up !== "number" || !Number.isFinite(up) || up < 0) {
      return `items[${i}].unitPrice must be a non-negative number`;
    }
  }

  return null;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  if (req.method !== "POST") {
    return jsonErr(405, "Method not allowed; use POST");
  }

  let body: Record<string, unknown>;
  try {
    body = await req.json();
  } catch {
    return jsonErr(400, "Invalid JSON body");
  }

  const v = validate(body);
  if (v) return jsonErr(400, v);

  const token = Deno.env.get("BINA_TOKEN")?.trim();
  if (!token) return jsonErr(500, "BINA_TOKEN secret is not configured");

  const clientRaw = body.client as Record<string, unknown>;
  const custId = parseCustId(clientRaw.binaCustomerId)!;
  const client = clientRaw as unknown as IncomingClient;
  const items = body.items as IncomingItem[];

  const titleRaw = typeof body.title === "string" ? body.title.trim() : "";
  const remarkRaw = typeof body.remark === "string" ? body.remark.trim() : "";

  const binaPayload = {
    tokenId: token,
    requestId: Date.now(),
    docType: DOC_TYPE,
    docWithvat: 1,
    docTitle: titleRaw || DOC_TITLE_DEFAULT,
    docStatus: "חדש",
    docRemark: remarkRaw,
    Cust: {
      custId,
      custName: client.name.trim().substring(0, 120),
      custCity: client.city.trim().substring(0, 80),
      custAddress: client.address.trim().substring(0, 200),
      custIshKheser: client.contactPerson.trim().substring(0, 120),
      custTel: typeof client.phone === "string" ? client.phone.trim().substring(0, 40) : "",
      custEmail: typeof client.email === "string" ? client.email.trim().substring(0, 120) : "",
    },
    docItems: items.map((it, idx) => {
      const idPart = typeof it.itemId === "string" && it.itemId.trim()
        ? it.itemId.trim()
        : `line-${idx + 1}`;
      const disc =
        typeof it.discount === "number" && Number.isFinite(it.discount) ? it.discount : 0;
      return {
        ItemId: idPart,
        ItemDesc: it.description.trim().substring(0, 240),
        ItemQty: Math.floor(it.quantity),
        UnitPrice: String(it.unitPrice),
        Unitcurrency: "ILS",
        CurValue: 1,
        Discount: disc,
      };
    }),
  };

  try {
    const res = await fetchBinaViaQuotaGuard(BINA_ORDER_URL, binaPayload);

    if (!res.ok) {
      return jsonErr(502, `Bina HTTP ${res.status}`);
    }

    let parsed: unknown;
    try {
      parsed = JSON.parse(res.text);
    } catch {
      return jsonErr(502, "Bina returned invalid JSON");
    }

    const responseObj = Array.isArray(parsed) ? parsed[0] : parsed;
    const ro = responseObj as Record<string, unknown> | null;

    if (!ro || ro.ResCode === undefined) {
      return jsonErr(502, "Unexpected Bina response shape");
    }

    const rc = ro.ResCode;
    if (rc !== 0 && rc !== "0") {
      const errText = String(ro.ResMsg ?? ro.ResMsgHe ?? "Unknown error");
      return new Response(JSON.stringify({ success: false, error: errText }), {
        status: 400,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const docId = ro.docId ?? ro.DocId;
    return new Response(
      JSON.stringify({ success: true, binaOrderId: docId }),
      {
        status: 200,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      },
    );
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    return jsonErr(500, msg);
  }
});

// Package I — one-shot historical import from Bina (docType -15).
// Manual only. Pass { from_date, to_date } as YYYY-MM-DD (max ~40 days).
// Creates tasks with is_archive=true. Never updates existing rows.

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import { fetchBinaViaQuotaGuard } from "../_shared/bina-proxy-fetch.ts";
import { rejectDisallowedInternalOrigin } from "../_shared/cors.ts";

const BINA_API_URL = "https://webfiles.binaw.com/post/PostJsonDocV2.aspx";
const DEFAULT_DEPT = "חדש";
const MAX_RANGE_DAYS = 40;

const DEPARTMENT_CODES: Record<string, string> = {
  "2": "ביגוד ומוצרי פרסום",
  "3": "דיגיטלי צבעוני",
  "4": "דיגיטלי שחור לבן",
  "5": "אופסט",
  "6": "עבודות חוץ",
  "7": "מתקני תצוגה ומוצרים נלווים",
  "8": "פורמט רחב",
};

const VALID_DEPTS = [
  "פורמט רחב",
  "דיגיטלי צבעוני",
  "דיגיטלי שחור לבן",
  "אופסט",
  "עבודות חוץ",
  "משלוחים",
  "תזכורות",
  "ביגוד ומוצרי פרסום",
  "מתקני תצוגה ומוצרים נלווים",
];

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}

function parseHebrewDateForDB(dateStr: unknown): string | null {
  if (!dateStr || typeof dateStr !== "string") return null;
  const parts = dateStr.trim().split("/");
  if (parts.length !== 3) return null;
  const [dd, mm, yyyy] = parts;
  if (!dd || !mm || !yyyy) return null;
  return `${yyyy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}`;
}

function num(v: unknown): number | null {
  if (v === undefined || v === null || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function normalizeBinaSalesAgentServer(val: unknown): string | null {
  const v = String(val ?? "").trim();
  if (!v) return null;
  if (v === "כפיר" || /^כפיר\b/.test(v)) return "כפיר צמח";
  if (/^נטלי\b/.test(v)) return "נטלי";
  if (/^ברק\b/.test(v)) return "ברק";
  return v;
}

function extractDept(title: string): { dept: string; cleanTitle: string } {
  if (!title) return { dept: DEFAULT_DEPT, cleanTitle: "" };
  const match = title.match(/^\s*\[([^\]]+)\]\s*(.*)/);
  if (match) {
    const extractedDept = match[1].trim();
    const cleanTitle = match[2].trim();
    if (VALID_DEPTS.includes(extractedDept)) {
      return { dept: extractedDept, cleanTitle: cleanTitle || title };
    }
    return { dept: DEFAULT_DEPT, cleanTitle: title };
  }
  return { dept: DEFAULT_DEPT, cleanTitle: title };
}

function deptFromOrderItems(order: Record<string, unknown>): string | null {
  const nested = order.Order as { items?: unknown[] } | undefined;
  const items = nested?.items;
  if (!Array.isArray(items) || !items.length) return null;
  const counts = new Map<string, number>();
  for (const raw of items) {
    const code = String((raw as { itemId?: unknown }).itemId ?? "").trim();
    const dept = DEPARTMENT_CODES[code];
    if (!dept) continue;
    counts.set(dept, (counts.get(dept) || 0) + 1);
  }
  let best: string | null = null;
  let n = 0;
  for (const [d, c] of counts) {
    if (c > n) {
      best = d;
      n = c;
    }
  }
  return best;
}

function ymdDays(from: string, to: string): number {
  const a = Date.parse(from + "T00:00:00Z");
  const b = Date.parse(to + "T00:00:00Z");
  if (!Number.isFinite(a) || !Number.isFinite(b)) return Infinity;
  return Math.round((b - a) / 86400000) + 1;
}

async function insertArchiveItems(
  supabase: ReturnType<typeof createClient>,
  taskId: string,
  binaOrderId: number,
  order: Record<string, unknown>,
): Promise<boolean> {
  const nested = order.Order as { items?: unknown[] } | undefined;
  const rawItems = nested?.items && Array.isArray(nested.items) ? nested.items : [];
  if (!rawItems.length) return true;

  const rows: Record<string, unknown>[] = [];
  let fallbackLine = 0;
  for (const raw of rawItems) {
    fallbackLine += 1;
    const item = raw as Record<string, unknown>;
    const itemCode = String(item.itemId ?? "").trim();
    if (!itemCode || !DEPARTMENT_CODES[itemCode]) continue;
    const ln = Number(item.itemLineNumber);
    rows.push({
      task_id: taskId,
      bina_order_id: binaOrderId,
      line_number: Number.isFinite(ln) ? ln : fallbackLine,
      bina_item_code: itemCode,
      department: DEPARTMENT_CODES[itemCode],
      description: String(item.itemDesc ?? "").trim() || "—",
      quantity: Number(item.itemQty) || 0,
      price: Number(item.itemPrice) || 0,
      total: Number(item.itemTotal) || 0,
      status: "מוכן",
    });
  }
  if (!rows.length) return true;
  const { error } = await supabase.from("task_items").upsert(rows, {
    onConflict: "bina_order_id,line_number",
  });
  if (error) {
    console.error("[bina-import-history] task_items:", error.message);
    return false;
  }
  await supabase.from("tasks").update({ has_items: true }).eq("id", taskId);
  return true;
}

Deno.serve(async (req) => {
  const originBlock = rejectDisallowedInternalOrigin(req);
  if (originBlock) return originBlock;
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);

  let body: Record<string, unknown> = {};
  try {
    body = await req.json();
  } catch {
    return json({ error: "Invalid JSON" }, 400);
  }

  const fromDate = String(body.from_date || "").trim();
  const toDate = String(body.to_date || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fromDate) || !/^\d{4}-\d{2}-\d{2}$/.test(toDate)) {
    return json({ error: "from_date and to_date required as YYYY-MM-DD" }, 400);
  }
  const days = ymdDays(fromDate, toDate);
  if (days < 1 || days > MAX_RANGE_DAYS) {
    return json({
      error: `טווח מקסימלי ${MAX_RANGE_DAYS} ימים (קיבלנו ${days}). הרץ חודש/שבועיים בכל פעם.`,
      from_date: fromDate,
      to_date: toDate,
    }, 400);
  }

  const binaToken = Deno.env.get("BINA_TOKEN");
  if (!binaToken) return json({ error: "BINA_TOKEN missing" }, 500);

  const supabase = createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  );

  const requestBody = {
    tokenId: binaToken,
    docType: -15,
    fromDate,
    toDate,
  };

  console.log("[bina-import-history] range", fromDate, "→", toDate);

  let binaResponse;
  try {
    binaResponse = await fetchBinaViaQuotaGuard(BINA_API_URL, requestBody);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    return json({ error: "Bina fetch failed", details: msg, from_date: fromDate, to_date: toDate }, 502);
  }

  if (!binaResponse.ok) {
    return json({
      error: "Bina HTTP error",
      status: binaResponse.status,
      details: String(binaResponse.text || "").slice(0, 800),
      from_date: fromDate,
      to_date: toDate,
      hint: binaResponse.status === 502 ? "נסה טווח קצר יותר (שבועיים / שבוע)" : undefined,
    }, 502);
  }

  let binaData: unknown;
  try {
    binaData = JSON.parse(binaResponse.text);
  } catch {
    return json({ error: "Bina invalid JSON", details: String(binaResponse.text).slice(0, 400) }, 502);
  }

  if (Array.isArray(binaData) && (binaData as { ResCode?: number }[])[0]?.ResCode !== undefined &&
    (binaData as { ResCode?: number }[])[0]?.ResCode !== 0) {
    return json({ error: "Bina API error", binaResponse: binaData }, 400);
  }

  let orders: Record<string, unknown>[] = [];
  const bd = binaData as { Orders?: unknown[] };
  if (bd.Orders && Array.isArray(bd.Orders)) orders = bd.Orders as Record<string, unknown>[];
  else if (Array.isArray(binaData)) orders = binaData as Record<string, unknown>[];

  let fetched = orders.length;
  let created = 0;
  let skipped = 0;
  let failed = 0;
  const errorDetails: string[] = [];
  const nowIso = new Date().toISOString();

  for (const order of orders) {
    try {
      const binaOrderId = order.orderId;
      if (!binaOrderId) {
        failed++;
        errorDetails.push("order without orderId");
        continue;
      }

      const { data: existing } = await supabase
        .from("tasks")
        .select("id")
        .eq("bina_order_id", binaOrderId)
        .maybeSingle();

      if (existing?.id) {
        skipped++;
        continue;
      }

      const titleRaw = String(order.orderSubject || order.orderTitle || "").trim();
      const { dept: deptFromTitle, cleanTitle } = extractDept(titleRaw);
      const dept = deptFromOrderItems(order) || deptFromTitle;
      const binaOrderDate = parseHebrewDateForDB(order.orderDate);
      let dueDate: string | null = null;
      if (order.orderDeliveryDate) {
        dueDate = parseHebrewDateForDB(order.orderDeliveryDate);
      }

      const taskData = {
        title: cleanTitle || `הזמנה #${binaOrderId}`,
        dept,
        status: "הושלם",
        priority: "רגיל",
        client_name: String(order.custName || "").trim(),
        contact: String(order.orderTo || "").trim() || null,
        due_date: dueDate,
        notes: null as string | null,
        bina_order_id: binaOrderId,
        bina_cust_id: order.custId || null,
        bina_cust_address: order.custAddress || null,
        bina_cust_city: order.custCity || null,
        bina_synced_at: nowIso,
        source: "bina_history",
        created_by: "bina-import-history",
        total_amount: num(order.orderTotalAfterDiscount) ?? num(order.orderTotal),
        total_inc_vat: num(order.orderTotalIncVat),
        discount_amount: num(order.orderDiscount),
        sales_agent: normalizeBinaSalesAgentServer(order.orderSalesMan),
        bina_order_date: binaOrderDate,
        bina_order_status: order.orderStatus != null ? String(order.orderStatus).trim() || null : null,
        bina_order_state: order.orderState != null ? String(order.orderState).trim() || null : null,
        is_archive: true,
        archive_imported_at: nowIso,
        completed_at: binaOrderDate ? `${binaOrderDate}T12:00:00.000Z` : nowIso,
        has_items: false,
      };

      const { data: inserted, error: insertError } = await supabase
        .from("tasks")
        .insert(taskData)
        .select("id")
        .single();

      if (insertError || !inserted?.id) {
        failed++;
        errorDetails.push(`Order ${binaOrderId}: ${insertError?.message || "insert failed"}`);
        continue;
      }

      created++;
      await insertArchiveItems(supabase, inserted.id as string, Number(binaOrderId), order);
    } catch (e) {
      failed++;
      errorDetails.push(e instanceof Error ? e.message : String(e));
    }
  }

  const summary = {
    success: true,
    from_date: fromDate,
    to_date: toDate,
    fetched,
    created,
    skipped,
    failed,
    errorDetails: failed ? errorDetails.slice(0, 30) : undefined,
  };
  console.log("[bina-import-history] summary", JSON.stringify(summary));
  return json(summary);
});

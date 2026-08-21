// Package I — one-shot historical quotes import from Bina (docType -14).
// Manual only. Pass { from_date, to_date } as YYYY-MM-DD (max ~40 days).
// Creates quotes with is_archive=true. Never updates non-archive rows.

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import { fetchBinaViaQuotaGuard } from "../_shared/bina-proxy-fetch.ts";
import { rejectDisallowedInternalOrigin } from "../_shared/cors.ts";

const BINA_API_URL = "https://webfiles.binaw.com/post/PostJsonDocV2.aspx";
const MAX_RANGE_DAYS = 40;

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

function ymdDays(from: string, to: string): number {
  const a = Date.parse(from + "T00:00:00Z");
  const b = Date.parse(to + "T00:00:00Z");
  if (!Number.isFinite(a) || !Number.isFinite(b)) return Infinity;
  return Math.round((b - a) / 86400000) + 1;
}

function extractDocs(binaData: unknown): Record<string, unknown>[] {
  if (!binaData || typeof binaData !== "object") return [];
  if (Array.isArray(binaData)) return binaData as Record<string, unknown>[];
  const o = binaData as Record<string, unknown>;
  if (Array.isArray(o.docs)) return o.docs as Record<string, unknown>[];
  if (Array.isArray(o.Docs)) return o.Docs as Record<string, unknown>[];
  if (Array.isArray(o.Quotes)) return o.Quotes as Record<string, unknown>[];
  return [];
}

function extractItems(doc: Record<string, unknown>): Record<string, unknown>[] {
  const nested = doc.doc as { items?: unknown[] } | undefined;
  if (Array.isArray(nested?.items)) return nested!.items as Record<string, unknown>[];
  if (Array.isArray(doc.items)) return doc.items as Record<string, unknown>[];
  return [];
}

async function insertArchiveQuoteItems(
  supabase: ReturnType<typeof createClient>,
  quoteId: string,
  doc: Record<string, unknown>,
): Promise<{ ok: boolean; count: number }> {
  const rawItems = extractItems(doc);
  if (!rawItems.length) return { ok: true, count: 0 };

  const rows: Record<string, unknown>[] = [];
  let fallbackLine = 0;
  for (const raw of rawItems) {
    fallbackLine += 1;
    const item = raw as Record<string, unknown>;
    const ln = Number(item.itemLineNumber);
    const itemId = String(item.itemId ?? "").trim();
    rows.push({
      quote_id: quoteId,
      line_number: Number.isFinite(ln) ? ln : fallbackLine,
      // NOT NULL — empty string when Bina itemId is empty; no department guessing
      item_name: itemId,
      description: String(item.itemDesc ?? "").trim() || null,
      quantity: Number(item.itemQty) || 0,
      unit_price: Number(item.itemPrice) || 0,
      discount_pct: num(item.itemDiscount) ?? 0,
      total: Number(item.itemTotal) || 0,
    });
  }

  const { error: delErr } = await supabase.from("quote_items").delete().eq("quote_id", quoteId);
  if (delErr) {
    console.error("[bina-import-quotes-history] quote_items delete:", delErr.message);
    return { ok: false, count: 0 };
  }

  const { error } = await supabase.from("quote_items").insert(rows);
  if (error) {
    console.error("[bina-import-quotes-history] quote_items insert:", error.message);
    return { ok: false, count: 0 };
  }
  return { ok: true, count: rows.length };
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
    docType: -14,
    fromDate,
    toDate,
  };

  console.log("[bina-import-quotes-history] range", fromDate, "→", toDate);

  let binaResponse;
  try {
    binaResponse = await fetchBinaViaQuotaGuard(BINA_API_URL, requestBody, {
      timeoutMs: 90_000,
      maxAttempts: 2,
    });
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

  if (
    Array.isArray(binaData) &&
    (binaData as { ResCode?: number }[])[0]?.ResCode !== undefined &&
    (binaData as { ResCode?: number }[])[0]?.ResCode !== 0
  ) {
    return json({ error: "Bina API error", binaResponse: binaData }, 400);
  }

  const docs = extractDocs(binaData);
  const nowIso = new Date().toISOString();
  let fetched = docs.length;
  let created = 0;
  let skipped = 0;
  let itemsCreated = 0;
  let itemsBackfilled = 0;
  let failed = 0;
  const errorDetails: string[] = [];

  for (const doc of docs) {
    try {
      const docIdRaw = doc.docId;
      if (docIdRaw === undefined || docIdRaw === null || docIdRaw === "") {
        failed++;
        errorDetails.push("doc without docId");
        continue;
      }
      const binaDocId = Number(docIdRaw);
      if (!Number.isFinite(binaDocId)) {
        failed++;
        errorDetails.push(`invalid docId: ${docIdRaw}`);
        continue;
      }

      const { data: existing } = await supabase
        .from("quotes")
        .select("id, is_archive")
        .eq("bina_doc_id", binaDocId)
        .maybeSingle();

      if (existing?.id) {
        if (existing.is_archive) {
          const r = await insertArchiveQuoteItems(supabase, existing.id as string, doc);
          if (!r.ok) {
            failed++;
            errorDetails.push(`Quote ${binaDocId}: items backfill failed`);
          } else {
            skipped++;
            itemsBackfilled += r.count;
          }
        } else {
          skipped++;
        }
        continue;
      }

      const custId = Number(doc.custId);
      if (!Number.isFinite(custId)) {
        failed++;
        errorDetails.push(`Quote ${binaDocId}: missing custId`);
        continue;
      }

      const binaDocDate = parseHebrewDateForDB(doc.docDate);
      const titleRaw = String(doc.docTitle ?? "").trim();
      const title = titleRaw || `הצעה #${binaDocId}`;
      const custName = String(doc.custName ?? "").trim() || "—";
      const docStatus = doc.docStatus != null ? String(doc.docStatus).trim() : "";
      const sentAt = binaDocDate ? `${binaDocDate}T12:00:00.000Z` : nowIso;

      const quoteRow = {
        bina_doc_id: binaDocId,
        bina_cust_id: custId,
        bina_cust_name: custName,
        bina_cust_address: String(doc.custAddress ?? "").trim() || null,
        bina_cust_city: String(doc.custCity ?? "").trim() || null,
        title,
        contact_person: String(doc.docTo ?? "").trim() || null,
        sales_agent: normalizeBinaSalesAgentServer(doc.docSalesMan),
        subtotal: num(doc.docTotalAfterDiscount) ?? num(doc.docTotal) ?? 0,
        vat_amount: num(doc.docVat) ?? 0,
        total: num(doc.docTotalIncVat) ?? num(doc.docTotalAfterDiscount) ?? num(doc.docTotal) ?? 0,
        total_amount: num(doc.docTotalIncVat) ?? num(doc.docTotalAfterDiscount) ?? num(doc.docTotal),
        status: docStatus || "נשלחה",
        quote_status: /בוטל|נדח/.test(docStatus) ? "rejected" : "sent",
        sent_at: sentAt,
        bina_doc_date: binaDocDate,
        bina_synced_at: nowIso,
        created_by: "bina-import-quotes-history",
        is_archive: true,
        archive_imported_at: nowIso,
        created_at: sentAt,
      };

      const { data: inserted, error: insertError } = await supabase
        .from("quotes")
        .insert(quoteRow)
        .select("id")
        .single();

      if (insertError || !inserted?.id) {
        failed++;
        errorDetails.push(`Quote ${binaDocId}: ${insertError?.message || "insert failed"}`);
        continue;
      }

      created++;
      const r = await insertArchiveQuoteItems(supabase, inserted.id as string, doc);
      if (!r.ok) {
        failed++;
        errorDetails.push(`Quote ${binaDocId}: items insert failed`);
      } else {
        itemsCreated += r.count;
      }
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
    items_created: itemsCreated,
    items_backfilled: itemsBackfilled,
    items: itemsCreated + itemsBackfilled,
    failed,
    errorDetails: failed ? errorDetails.slice(0, 30) : undefined,
  };
  console.log("[bina-import-quotes-history] summary", JSON.stringify(summary));
  return json(summary);
});

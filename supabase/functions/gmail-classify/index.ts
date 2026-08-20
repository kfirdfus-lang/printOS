// G2 — classify Gmail messages in the background. Never blocks the inbox.
// Read-only Gmail: GET body text only. No send/delete/modify.

import { rejectDisallowedInternalOrigin } from "../_shared/cors.ts";
import {
  corsHeaders,
  extractAttachments,
  extractBodies,
  gmailGet,
  getValidAccessToken,
  json,
  requireAdmin,
  serviceClient,
} from "../_shared/gmail.ts";

const MODEL = "claude-sonnet-5";
const ALLOWED = new Set([
  "order",
  "quote_request",
  "supplier_invoice",
  "design_approval",
  "general",
  "irrelevant",
]);

const SYSTEM_PROMPT = `אתה מסווג מיילים עבור בית דפוס בישראל (נטלי פתרונות הדפסה).

תפקידך: לקרוא מייל ולסווג אותו לקטגוריה אחת.

הקטגוריות:

order — הזמנת עבודה. הלקוח מבקש לייצר משהו,
  בדרך כלל עם קובץ מצורף וכמות מוגדרת.

quote_request — שאלת מחיר. הלקוח שואל כמה עולה,
  בלי להזמין בפועל.

supplier_invoice — חשבונית מספק. מסמך שנטלי צריכה לשלם.
  שים לב: המסמך מופנה אל נטלי, לא ממנה.

design_approval — אישור או תיקון עיצוב. הלקוח מגיב
  על הדפסת ניסיון או קובץ שנשלח אליו.

general — תקשורת עסקית שאינה אף אחד מהנ״ל.

irrelevant — פרסום, ניוזלטר, ספאם, התראות מערכת.

כללי הכרעה:
- מייל עם קובץ PDF/AI/EPS מלקוח = order, גם בלי מילים מפורשות
- "כמה עולה" / "מחירון" / "הצעת מחיר" = quote_request
- מסמך שמופנה אל נטלי עם סכום = supplier_invoice
- "מאשר" / "תיקון קטן" / "אפשר להדפיס" = design_approval
- ניוזלטר / קידום מכירות / התראת מערכת = irrelevant

כללים:
1. החזר JSON בלבד. בלי טקסט לפני או אחרי, בלי markdown.
2. אם אתה לא בטוח — confidence נמוך. אל תנחש בביטחון.
3. שם הלקוח — כפי שמופיע בחתימה או בכתובת השולח.
4. אם ערך טקסט מכיל גרשיים — הברח אותם כ-\\" (חשוב:
   שמות חברות ישראליות נגמרים ב-בע"מ).

מבנה:
{
  "category": "order",
  "confidence": 0.9,
  "reason": "משפט קצר בעברית",
  "client_name": "שם" או null,
  "extracted_data": {
    "quantity": מספר או null,
    "description": "מה מבקשים" או null,
    "deadline": "YYYY-MM-DD" או null
  }
}`;

Deno.serve(async (req) => {
  const originBlock = rejectDisallowedInternalOrigin(req);
  if (originBlock) return originBlock;
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);

  let body: Record<string, unknown> = {};
  try {
    body = await req.json();
  } catch {
    return json({ error: "Invalid JSON body" }, 400);
  }

  const action = String(body.action || "").trim();
  if (["send", "delete", "modify", "trash"].includes(action)) {
    return json({ error: "gmail-classify is read-only" }, 400);
  }

  try {
    const sb = serviceClient();
    const admin = await requireAdmin(sb, body.user_id);
    if ("error" in admin) return admin.error;
    if (action === "lookup") return await handleLookup(sb, body);
    if (action === "correct") return await handleCorrect(sb, body);
    if (action === "classify") return await handleClassify(sb, admin.user.id, body);
    if (action === "mark_handled") return await handleHandled(sb, admin.user, body, true);
    if (action === "mark_unhandled") return await handleHandled(sb, admin.user, body, false);
    if (action === "save_alias") return await handleSaveAlias(sb, body);
    if (action === "list_aliases") return await handleListAliases(sb);
    return json({ error: "unknown action" }, 400);
  } catch (e) {
    return json({ error: e instanceof Error ? e.message : String(e) }, 500);
  }
});

const CLASS_FIELDS =
  "message_id,thread_id,category,confidence,reason,client_name,extracted_data,user_corrected_category,classified_at,handled,handled_at,handled_by,handled_reason,created_order_id";

async function handleLookup(sb: ReturnType<typeof serviceClient>, body: Record<string, unknown>) {
  const ids = Array.isArray(body.message_ids) ? body.message_ids.map((x) => String(x || "")).filter(Boolean) : [];
  if (!ids.length) return json({ rows: [] });
  const { data, error } = await sb.from("gmail_classifications").select(CLASS_FIELDS)
    .in("message_id", ids.slice(0, 80));
  if (error) return json({ rows: [], error: error.message });
  return json({ rows: data || [] });
}

async function handleCorrect(sb: ReturnType<typeof serviceClient>, body: Record<string, unknown>) {
  const messageId = String(body.message_id || "").trim();
  const category = String(body.category || "").trim();
  if (!messageId || !ALLOWED.has(category)) return json({ error: "invalid message_id or category" }, 400);
  const { data, error } = await sb.from("gmail_classifications").update({
    user_corrected_category: category,
    corrected_at: new Date().toISOString(),
  }).eq("message_id", messageId).select(CLASS_FIELDS).maybeSingle();
  if (error) return json({ error: error.message }, 500);
  return json({ row: data });
}

async function handleClassify(
  sb: ReturnType<typeof serviceClient>,
  userId: string,
  body: Record<string, unknown>,
) {
  const incoming = Array.isArray(body.messages) ? body.messages.slice(0, 10) : [];
  if (!incoming.length) return json({ rows: [] });

  const ids = incoming.map((m) => String((m as { id?: string }).id || "")).filter(Boolean);
  const { data: existing } = await sb.from("gmail_classifications").select(CLASS_FIELDS)
    .in("message_id", ids);
  const have = new Set((existing || []).map((r: { message_id: string }) => r.message_id));
  const todo = incoming.filter((m) => {
    const id = String((m as { id?: string }).id || "");
    return id && !have.has(id);
  });

  const tok = await getValidAccessToken(sb, userId);
  const access = "token" in tok ? tok.token : "";
  const created = [];
  for (const raw of todo) {
    const msg = raw as { id?: string; threadId?: string; from?: string; subject?: string; snippet?: string; body?: string };
    const row = await classifyOne(sb, access, msg);
    if (row) created.push(row);
  }
  return json({ rows: [...(existing || []), ...created] });
}

async function classifyOne(
  sb: ReturnType<typeof serviceClient>,
  accessToken: string,
  msg: { id?: string; threadId?: string; from?: string; subject?: string; snippet?: string; body?: string; attachmentNames?: string[] },
) {
  const id = String(msg.id || "").trim();
  if (!id) return null;
  let bodyText = String(msg.body || msg.snippet || "").slice(0, 3000);
  let attachmentNames = Array.isArray(msg.attachmentNames) ? msg.attachmentNames.filter(Boolean) : [];
  if (accessToken && (bodyText.length < 200 || !attachmentNames.length)) {
    try {
      const full = await gmailGet(accessToken, `users/me/messages/${encodeURIComponent(id)}?format=full`);
      if (full.ok) {
        const payload = (full.data as { payload?: Parameters<typeof extractBodies>[0] }).payload;
        const extracted = extractBodies(payload);
        bodyText = (extracted.text || extracted.html.replace(/<[^>]+>/g, " ") || bodyText).slice(0, 3000);
        if (!attachmentNames.length) {
          attachmentNames = extractAttachments(payload).map((a) => a.filename).filter(Boolean);
        }
      }
    } catch {
      // keep snippet
    }
  }

  const fallback = {
    category: "general",
    confidence: 0.2,
    reason: "סיווג נכשל",
    client_name: null as string | null,
    extracted_data: { quantity: null, description: null, deadline: null },
  };
  let parsed = fallback;
  try {
    parsed = await callClaude(msg.from || "", msg.subject || "", bodyText, attachmentNames) || fallback;
  } catch {
    parsed = fallback;
  }
  const category = ALLOWED.has(parsed.category) ? parsed.category : "general";
  const row = {
    message_id: id,
    thread_id: msg.threadId || null,
    category,
    confidence: clampConf(parsed.confidence),
    reason: String(parsed.reason || "").slice(0, 500) || null,
    client_name: parsed.client_name ? String(parsed.client_name).slice(0, 200) : null,
    extracted_data: parsed.extracted_data || null,
    model_used: MODEL,
    classified_at: new Date().toISOString(),
  };
  const { data, error } = await sb.from("gmail_classifications").upsert(row, { onConflict: "message_id" }).select(CLASS_FIELDS).maybeSingle();
  if (error) return { ...row, user_corrected_category: null };
  return data;
}

async function handleHandled(
  sb: ReturnType<typeof serviceClient>,
  user: { id: string; name: string | null },
  body: Record<string, unknown>,
  handled: boolean,
) {
  const messageId = String(body.message_id || "").trim();
  if (!messageId) return json({ error: "missing message_id" }, 400);
  const patch = handled
    ? {
      handled: true,
      handled_at: new Date().toISOString(),
      handled_by: user.name || user.id,
      handled_reason: String(body.reason || "manual").trim() || "manual",
      created_order_id: body.created_order_id ? String(body.created_order_id) : null,
    }
    : {
      handled: false,
      handled_at: null,
      handled_by: null,
      handled_reason: null,
      created_order_id: null,
    };
  const { data, error } = await sb.from("gmail_classifications").update(patch).eq("message_id", messageId)
    .select(CLASS_FIELDS).maybeSingle();
  if (error) return json({ error: error.message }, 500);
  if (data) return json({ row: data });
  const { data: inserted, error: insErr } = await sb.from("gmail_classifications").upsert({
    message_id: messageId,
    category: "general",
    confidence: 0,
    classified_at: new Date().toISOString(),
    ...patch,
  }, { onConflict: "message_id" }).select(CLASS_FIELDS).maybeSingle();
  if (insErr) return json({ error: insErr.message }, 500);
  return json({ row: inserted });
}

async function handleSaveAlias(sb: ReturnType<typeof serviceClient>, body: Record<string, unknown>) {
  const clientId = String(body.client_id || "").trim();
  const email = String(body.email || "").trim().toLowerCase().replace(/\s+/g, "");
  if (!clientId || !email || !email.includes("@")) {
    return json({ error: "invalid client_id or email" }, 400);
  }
  const { error } = await sb.from("client_email_aliases").upsert(
    { client_id: clientId, email },
    { onConflict: "email" },
  );
  if (error) return json({ error: error.message }, 500);
  return json({ ok: true, email, client_id: clientId });
}

async function handleListAliases(sb: ReturnType<typeof serviceClient>) {
  const { data, error } = await sb.from("client_email_aliases").select("email,client_id").limit(5000);
  if (error) return json({ rows: [], error: error.message });
  return json({ rows: data || [] });
}

function clampConf(v: unknown): number {
  const n = Number(v);
  if (!Number.isFinite(n)) return 0.3;
  return Math.max(0, Math.min(1, Math.round(n * 100) / 100));
}

async function callClaude(from: string, subject: string, body: string, attachmentNames: string[] = []) {
  const apiKey = Deno.env.get("ANTHROPIC_API_KEY");
  if (!apiKey) return null;
  const attLine = attachmentNames.length ? `\nAttachments: ${attachmentNames.join(", ")}` : "";
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 1000,
      system: SYSTEM_PROMPT,
      messages: [{
        role: "user",
        content: `From: ${from}\nSubject: ${subject}${attLine}\n\n${body}`,
      }],
    }),
  });
  if (!res.ok) return null;
  const data = await res.json();
  const raw = String(data?.content?.[0]?.text || "").trim();
  const cleaned = raw.replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();
  try {
    return JSON.parse(cleaned);
  } catch {
    const m = cleaned.match(/\{[\s\S]*\}/);
    if (!m) return null;
    try {
      return JSON.parse(m[0]);
    } catch {
      return null;
    }
  }
}

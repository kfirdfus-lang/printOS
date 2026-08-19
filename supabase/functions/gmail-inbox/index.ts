// Gmail inbox — list / get / attachment / label modify / snooze.
// No delete, no trash.

import { rejectDisallowedInternalOrigin } from "../_shared/cors.ts";
import {
  corsHeaders,
  extractAttachments,
  extractBodies,
  gmailBatchGet,
  gmailGet,
  gmailModifyLabels,
  getValidAccessToken,
  headerMap,
  json,
  payloadHasAttachments,
  requireAdmin,
  serviceClient,
} from "../_shared/gmail.ts";

const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;
const ATTACHMENT_TOO_LARGE =
  "הקובץ גדול מדי (מעל 25MB) — יש להוריד אותו מג'ימייל";

const READ_ACTIONS = new Set([
  "list", "get", "attachment", "counts", "snooze_list",
]);
const MODIFY_ACTIONS = new Set([
  "mark_read", "mark_unread", "star", "unstar", "archive", "snooze",
]);
const BLOCKED_ACTIONS = new Set(["send", "delete", "trash", "untrash", "insert"]);

const CATEGORY_QUERY: Record<string, string> = {
  primary: "category:primary",
  promotions: "category:promotions",
  social: "category:social",
  updates: "category:updates",
  starred: "is:starred",
  snoozed: "label:Snoozed",
};

const CATEGORY_LABEL: Record<string, string> = {
  CATEGORY_PERSONAL: "primary",
  CATEGORY_PROMOTIONS: "promotions",
  CATEGORY_SOCIAL: "social",
  CATEGORY_UPDATES: "updates",
};

function buildListQuery(body: Record<string, unknown>): string {
  const cat = String(body.category || "primary").trim();
  const catQ = CATEGORY_QUERY[cat] || CATEGORY_QUERY.primary;
  const userQ = String(body.q || "").trim();
  if (!userQ) return catQ;
  if (/\bcategory:|is:starred|label:/.test(userQ)) return userQ;
  return `${catQ} ${userQ}`;
}

async function unreadByCategory(token: string): Promise<Record<string, number>> {
  const counts = { primary: 0, promotions: 0, social: 0, updates: 0 };
  const res = await gmailGet(token, "users/me/labels");
  if (!res.ok) return counts;
  const labels = ((res.data as { labels?: { id?: string; messagesUnread?: number }[] }).labels) || [];
  for (const lab of labels) {
    const key = CATEGORY_LABEL[String(lab.id || "")];
    if (!key) continue;
    counts[key as keyof typeof counts] = Number(lab.messagesUnread) || 0;
  }
  return counts;
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
    return json({ error: "Invalid JSON body" }, 400);
  }

  const action = String(body.action || "").trim();
  if (BLOCKED_ACTIONS.has(action)) {
    return json({ error: "delete and trash are not allowed" }, 400);
  }
  if (!READ_ACTIONS.has(action) && !MODIFY_ACTIONS.has(action)) {
    return json({ error: "unknown action" }, 400);
  }

  try {
    const sb = serviceClient();
    const admin = await requireAdmin(sb, body.user_id);
    if ("error" in admin) return admin.error;
    const tok = await getValidAccessToken(sb, admin.user.id);
    if ("error" in tok) return tok.error;

    if (action === "list") return await handleList(tok.token, body);
    if (action === "counts") {
      const unread = await unreadByCategory(tok.token);
      return json({ unread, primaryUnread: unread.primary || 0 });
    }
    if (action === "get") return await handleGet(tok.token, body, sb);
    if (action === "attachment") return await handleAttachment(tok.token, body);
    if (action === "mark_read") return await handleModify(tok.token, body, [], ["UNREAD"]);
    if (action === "mark_unread") return await handleModify(tok.token, body, ["UNREAD"], []);
    if (action === "star") return await handleModify(tok.token, body, ["STARRED"], []);
    if (action === "unstar") return await handleModify(tok.token, body, [], ["STARRED"]);
    if (action === "archive") return await handleModify(tok.token, body, [], ["INBOX"]);
    if (action === "snooze") return await handleSnooze(tok.token, sb, body);
    if (action === "snooze_list") return await handleSnoozeList(sb);
    return json({ error: "unknown action" }, 400);
  } catch (e) {
    return json({ error: e instanceof Error ? e.message : String(e) }, 500);
  }
});

async function handleList(token: string, body: Record<string, unknown>) {
  const maxResults = Math.min(Math.max(Number(body.maxResults) || 15, 1), 50);
  const pageToken = String(body.pageToken || "").trim();
  const q = buildListQuery(body);
  const qs = new URLSearchParams({ maxResults: String(maxResults), q });
  if (pageToken) qs.set("pageToken", pageToken);

  const listed = await gmailGet(token, `users/me/messages?${qs.toString()}`);
  if (!listed.ok) return json({ error: "gmail list failed", details: listed.data }, listed.status);

  const raw = listed.data as { messages?: { id: string; threadId: string }[]; nextPageToken?: string; resultSizeEstimate?: number };
  const ids = (raw.messages || []).map((m) => m.id).filter(Boolean);
  const headerQs = "format=metadata&metadataHeaders=From&metadataHeaders=Subject&metadataHeaders=Date&metadataHeaders=To";
  const paths = ids.map((id) => `users/me/messages/${encodeURIComponent(id)}?${headerQs}`);
  const metas = await gmailBatchGet(token, paths);
  const messages = metas.map((res, i) => {
    if (!res.ok) return null;
    return mapMeta(res.data, ids[i]);
  });

  return json({
    messages: messages.filter(Boolean),
    nextPageToken: raw.nextPageToken || null,
    resultSizeEstimate: raw.resultSizeEstimate || 0,
  });
}

function mapMeta(raw: unknown, fallbackId: string) {
  const msg = raw as {
    id?: string;
    threadId?: string;
    snippet?: string;
    internalDate?: string;
    labelIds?: string[];
    payload?: { headers?: { name?: string; value?: string }[]; filename?: string; mimeType?: string; body?: { attachmentId?: string; size?: number }; parts?: unknown[] };
  };
  const labels = msg.labelIds || [];
  const h = headerMap(msg.payload);
  return {
    id: msg.id || fallbackId,
    threadId: msg.threadId,
    from: h.from || "",
    to: h.to || "",
    subject: h.subject || "(ללא נושא)",
    date: h.date || "",
    snippet: msg.snippet || "",
    unread: labels.includes("UNREAD"),
    starred: labels.includes("STARRED"),
    hasAttachments: payloadHasAttachments(msg.payload as { filename?: string; body?: { attachmentId?: string }; parts?: unknown[] }),
    internalDate: msg.internalDate || null,
  };
}

async function handleGet(
  token: string,
  body: Record<string, unknown>,
  sb: ReturnType<typeof serviceClient>,
) {
  const messageId = String(body.messageId || "").trim();
  if (!messageId) return json({ error: "missing messageId" }, 400);
  const res = await gmailGet(token, `users/me/messages/${encodeURIComponent(messageId)}?format=full`);
  if (!res.ok) return json({ error: "gmail get failed", details: res.data }, res.status);
  const msg = res.data as {
    id: string;
    threadId: string;
    snippet?: string;
    labelIds?: string[];
    payload?: { headers?: { name?: string; value?: string }[]; mimeType?: string; filename?: string; body?: { data?: string; size?: number; attachmentId?: string }; parts?: unknown[] };
  };
  const h = headerMap(msg.payload);
  const bodies = extractBodies(msg.payload);
  const labels = msg.labelIds || [];

  if (labels.includes("UNREAD") && body.autoMarkRead !== false) {
    await gmailModifyLabels(token, messageId, [], ["UNREAD"]);
  }

  return json({
    id: msg.id,
    threadId: msg.threadId,
    from: h.from || "",
    to: h.to || "",
    cc: h.cc || "",
    subject: h.subject || "(ללא נושא)",
    date: h.date || "",
    snippet: msg.snippet || "",
    unread: false,
    starred: labels.includes("STARRED"),
    messageIdHeader: h["message-id"] || "",
    references: h.references || "",
    text: bodies.text,
    html: bodies.html,
    attachments: extractAttachments(msg.payload),
  });
}

async function handleAttachment(token: string, body: Record<string, unknown>) {
  const messageId = String(body.messageId || "").trim();
  const attachmentId = String(body.attachmentId || "").trim();
  const filename = String(body.filename || "attachment").trim() || "attachment";
  if (!messageId || !attachmentId) return json({ error: "missing messageId or attachmentId" }, 400);
  const declared = Number(body.size) || 0;
  if (declared > MAX_ATTACHMENT_BYTES) return json({ error: ATTACHMENT_TOO_LARGE }, 413);
  const res = await gmailGet(
    token,
    `users/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
  );
  if (!res.ok) return json({ error: "gmail attachment failed", details: res.data }, res.status);
  const att = res.data as { data?: string; size?: number };
  if (!att.data) return json({ error: "empty attachment" }, 404);
  const size = Number(att.size) || Math.ceil((att.data.length * 3) / 4);
  if (size > MAX_ATTACHMENT_BYTES) return json({ error: ATTACHMENT_TOO_LARGE }, 413);
  return json({ filename, data: att.data, size });
}

async function handleModify(
  token: string,
  body: Record<string, unknown>,
  add: string[],
  remove: string[],
) {
  const messageId = String(body.messageId || "").trim();
  if (!messageId) return json({ error: "missing messageId" }, 400);
  const res = await gmailModifyLabels(token, messageId, add, remove);
  if (!res.ok) return json({ error: "gmail modify failed", details: res.data }, res.status);
  return json({ ok: true, id: messageId });
}

async function handleSnooze(
  token: string,
  sb: ReturnType<typeof serviceClient>,
  body: Record<string, unknown>,
) {
  const messageId = String(body.messageId || "").trim();
  const threadId = String(body.threadId || "").trim() || null;
  const until = String(body.snooze_until || "").trim();
  if (!messageId || !until) return json({ error: "missing messageId or snooze_until" }, 400);
  const untilDate = Date.parse(until);
  if (Number.isNaN(untilDate) || untilDate <= Date.now()) {
    return json({ error: "snooze_until must be in the future" }, 400);
  }

  const archived = await gmailModifyLabels(token, messageId, [], ["INBOX"]);
  if (!archived.ok) return json({ error: "archive for snooze failed", details: archived.data }, archived.status);

  const { error } = await sb.from("gmail_snoozed").upsert({
    message_id: messageId,
    thread_id: threadId,
    snooze_until: new Date(untilDate).toISOString(),
    note: String(body.note || "").trim() || null,
    released: false,
  }, { onConflict: "message_id" });
  if (error) return json({ error: error.message }, 500);
  return json({ ok: true, message_id: messageId, snooze_until: new Date(untilDate).toISOString() });
}

async function handleSnoozeList(sb: ReturnType<typeof serviceClient>) {
  const { data, error } = await sb.from("gmail_snoozed")
    .select("message_id,thread_id,snooze_until,note,created_at")
    .eq("released", false)
    .gte("snooze_until", new Date().toISOString())
    .order("snooze_until", { ascending: true })
    .limit(100);
  if (error) return json({ rows: [], error: error.message });
  return json({ rows: data || [] });
}

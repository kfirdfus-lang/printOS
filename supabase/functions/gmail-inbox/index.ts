// Gmail pilot — read-only inbox (list / get / attachment).
// No send, delete, or modify.

import { rejectDisallowedInternalOrigin } from "../_shared/cors.ts";
import {
  corsHeaders,
  extractAttachments,
  extractBodies,
  gmailGet,
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

const READ_ACTIONS = new Set(["list", "get", "attachment"]);

const CATEGORY_QUERY: Record<string, string> = {
  primary: "category:primary",
  promotions: "category:promotions",
  social: "category:social",
  updates: "category:updates",
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
  if (/\bcategory:/.test(userQ)) return userQ;
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
const BLOCKED_ACTIONS = new Set(["send", "delete", "modify", "trash", "untrash", "insert"]);

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
  if (BLOCKED_ACTIONS.has(action) || !READ_ACTIONS.has(action)) {
    return json({ error: "read-only pilot — only list/get/attachment are allowed" }, 400);
  }

  try {
    const sb = serviceClient();
    const admin = await requireAdmin(sb, body.user_id);
    if ("error" in admin) return admin.error;
    const tok = await getValidAccessToken(sb, admin.user.id);
    if ("error" in tok) return tok.error;

    if (action === "list") return await handleList(tok.token, body);
    if (action === "get") return await handleGet(tok.token, body);
    if (action === "attachment") return await handleAttachment(tok.token, body);
    return json({ error: "unknown action" }, 400);
  } catch (e) {
    return json({ error: e instanceof Error ? e.message : String(e) }, 500);
  }
});

async function handleList(token: string, body: Record<string, unknown>) {
  const maxResults = Math.min(Math.max(Number(body.maxResults) || 25, 1), 50);
  const pageToken = String(body.pageToken || "").trim();
  const q = buildListQuery(body);
  const qs = new URLSearchParams({ maxResults: String(maxResults), q });
  if (pageToken) qs.set("pageToken", pageToken);

  const [listed, unread] = await Promise.all([
    gmailGet(token, `users/me/messages?${qs.toString()}`),
    unreadByCategory(token),
  ]);
  if (!listed.ok) return json({ error: "gmail list failed", details: listed.data }, listed.status);

  const raw = listed.data as { messages?: { id: string; threadId: string }[]; nextPageToken?: string; resultSizeEstimate?: number };
  const ids = (raw.messages || []).map((m) => m.id).filter(Boolean);
  const messages = await Promise.all(ids.map((id) => loadMeta(token, id)));

  return json({
    messages: messages.filter(Boolean),
    nextPageToken: raw.nextPageToken || null,
    resultSizeEstimate: raw.resultSizeEstimate || 0,
    unread,
  });
}

async function loadMeta(token: string, id: string) {
  const qs = new URLSearchParams({
    format: "metadata",
    metadataHeaders: "From",
  });
  qs.append("metadataHeaders", "Subject");
  qs.append("metadataHeaders", "Date");
  qs.append("metadataHeaders", "To");
  const res = await gmailGet(token, `users/me/messages/${encodeURIComponent(id)}?${qs.toString()}`);
  if (!res.ok) return null;
  const msg = res.data as {
    id: string;
    threadId: string;
    snippet?: string;
    internalDate?: string;
    labelIds?: string[];
    payload?: { headers?: { name?: string; value?: string }[]; filename?: string; mimeType?: string; body?: { attachmentId?: string; size?: number }; parts?: unknown[] };
  };
  const h = headerMap(msg.payload);
  return {
    id: msg.id,
    threadId: msg.threadId,
    from: h.from || "",
    to: h.to || "",
    subject: h.subject || "(ללא נושא)",
    date: h.date || "",
    snippet: msg.snippet || "",
    unread: (msg.labelIds || []).includes("UNREAD"),
    hasAttachments: payloadHasAttachments(msg.payload as { filename?: string; body?: { attachmentId?: string }; parts?: unknown[] }),
    internalDate: msg.internalDate || null,
  };
}

async function handleGet(token: string, body: Record<string, unknown>) {
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
  return json({
    id: msg.id,
    threadId: msg.threadId,
    from: h.from || "",
    to: h.to || "",
    cc: h.cc || "",
    subject: h.subject || "(ללא נושא)",
    date: h.date || "",
    snippet: msg.snippet || "",
    unread: (msg.labelIds || []).includes("UNREAD"),
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

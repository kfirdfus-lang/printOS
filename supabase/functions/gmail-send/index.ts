// G3 — send Gmail replies/forwards. User-initiated only. Full audit log.

import { rejectDisallowedInternalOrigin } from "../_shared/cors.ts";
import {
  corsHeaders,
  gmailPost,
  getValidAccessToken,
  json,
  requireAdmin,
  serviceClient,
} from "../_shared/gmail.ts";

const MAX_ATTACH_BYTES = 25 * 1024 * 1024;

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
  if (["delete", "trash", "modify"].includes(action)) {
    return json({ error: "gmail-send only supports send and list_sent" }, 400);
  }

  try {
    const sb = serviceClient();
    const admin = await requireAdmin(sb, body.user_id);
    if ("error" in admin) return admin.error;
    if (action === "send") return await handleSend(sb, admin.user, body);
    if (action === "list_sent") return await handleListSent(sb);
    return json({ error: "unknown action" }, 400);
  } catch (e) {
    return json({ error: e instanceof Error ? e.message : String(e) }, 500);
  }
});

async function handleListSent(sb: ReturnType<typeof serviceClient>) {
  const { data, error } = await sb.from("gmail_sent_log")
    .select("id,created_at,sent_by,to_addresses,cc_addresses,subject,body_preview,attachment_names,printos_document_type,printos_document_id,status,error_text,thread_id,gmail_message_id")
    .order("created_at", { ascending: false })
    .limit(100);
  if (error) return json({ rows: [], error: error.message });
  return json({ rows: data || [] });
}

async function handleSend(
  sb: ReturnType<typeof serviceClient>,
  user: { id: string; name: string | null },
  body: Record<string, unknown>,
) {
  const to = normEmails(body.to);
  const cc = normEmails(body.cc);
  const subject = String(body.subject || "").trim();
  const bodyHtml = String(body.body_html || "").trim();
  const attachments = Array.isArray(body.attachments) ? body.attachments : [];
  const inReplyTo = String(body.in_reply_to || "").trim();
  const references = String(body.references || "").trim();
  const threadId = String(body.thread_id || "").trim() || undefined;
  const replyMessageId = String(body.reply_message_id || "").trim() || null;
  const printosType = String(body.printos_document_type || "").trim() || null;
  const printosId = String(body.printos_document_id || "").trim() || null;

  if (!to.length) return json({ error: "חייב להיות לפחות נמען אחד" }, 400);
  if (!bodyHtml) return json({ error: "גוף ההודעה לא יכול להיות ריק" }, 400);

  let attachBytes = 0;
  const attNames: string[] = [];
  for (const raw of attachments) {
    const att = raw as { filename?: string; data_base64?: string; mime_type?: string };
    const name = String(att.filename || "file").trim();
    const b64 = String(att.data_base64 || "").trim();
    if (!b64) continue;
    const size = Math.ceil((b64.length * 3) / 4);
    attachBytes += size;
    attNames.push(name);
  }
  if (attachBytes > MAX_ATTACH_BYTES) {
    return json({ error: "סה\"כ גודל הקבצים מעל 25MB" }, 413);
  }

  const tok = await getValidAccessToken(sb, user.id);
  if ("error" in tok) return tok.error;

  const fromEmail = tok.email || "";
  const logBase = {
    sent_by: user.name || user.id,
    to_addresses: to,
    cc_addresses: cc.length ? cc : null,
    subject,
    body_preview: bodyHtml.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 500),
    in_reply_to_message_id: inReplyTo || null,
    thread_id: threadId || null,
    attachment_names: attNames.length ? attNames : null,
    printos_document_type: printosType,
    printos_document_id: printosId,
    status: "failed" as const,
    error_text: null as string | null,
  };

  const { data: logRow, error: logErr } = await sb.from("gmail_sent_log").insert(logBase).select("id").maybeSingle();
  if (logErr) return json({ error: logErr.message }, 500);

  try {
    const mime = buildMime({
      from: fromEmail,
      to,
      cc,
      subject,
      bodyHtml,
      inReplyTo,
      references,
      attachments: attachments.map((raw) => {
        const att = raw as { filename?: string; data_base64?: string; mime_type?: string };
        return {
          filename: String(att.filename || "file"),
          mimeType: String(att.mime_type || "application/octet-stream"),
          dataBase64: String(att.data_base64 || ""),
        };
      }).filter((a) => a.dataBase64),
    });
    const raw = b64UrlEncode(mime);
    const payload: Record<string, unknown> = { raw };
    if (threadId) payload.threadId = threadId;

    const sent = await gmailPost(tok.token, "users/me/messages/send", payload);
    if (!sent.ok) {
      const errText = JSON.stringify(sent.data).slice(0, 1000);
      await sb.from("gmail_sent_log").update({ status: "failed", error_text: errText }).eq("id", logRow?.id);
      return json({ error: "שליחה נכשלה", details: sent.data }, sent.status);
    }

    const gmailId = String((sent.data as { id?: string })?.id || "");
    await sb.from("gmail_sent_log").update({
      status: "sent",
      gmail_message_id: gmailId,
      error_text: null,
    }).eq("id", logRow?.id);

    if (replyMessageId) {
      await sb.from("gmail_classifications").update({
        handled: true,
        handled_at: new Date().toISOString(),
        handled_by: user.name || user.id,
        handled_reason: "replied",
      }).eq("message_id", replyMessageId);
    }

    return json({ ok: true, gmail_message_id: gmailId, log_id: logRow?.id });
  } catch (e) {
    const errText = e instanceof Error ? e.message : String(e);
    if (logRow?.id) {
      await sb.from("gmail_sent_log").update({ status: "failed", error_text: errText }).eq("id", logRow.id);
    }
    return json({ error: errText }, 500);
  }
}

function normEmails(v: unknown): string[] {
  if (!Array.isArray(v)) return [];
  return v.map((x) => String(x || "").trim().toLowerCase()).filter((e) => e.includes("@"));
}

function b64UrlEncode(raw: string): string {
  const bytes = new TextEncoder().encode(raw);
  let bin = "";
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function b64StdEncode(bytes: Uint8Array): string {
  let bin = "";
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin);
}

function decodeB64(data: string): Uint8Array {
  const pad = data.replace(/-/g, "+").replace(/_/g, "/");
  const bin = atob(pad);
  return Uint8Array.from(bin, (c) => c.charCodeAt(0));
}

function encodeQuotedPrintableHebrew(html: string): string {
  const bytes = new TextEncoder().encode(html);
  let out = "";
  let lineLen = 0;
  for (const b of bytes) {
    let chunk: string;
    if (b === 9 || b === 32) {
      chunk = String.fromCharCode(b);
    } else if (b >= 33 && b <= 126 && b !== 61) {
      chunk = String.fromCharCode(b);
    } else {
      chunk = "=" + b.toString(16).toUpperCase().padStart(2, "0");
    }
    if (lineLen + chunk.length > 75) {
      out += "=\r\n";
      lineLen = 0;
    }
    out += chunk;
    lineLen += chunk.length;
  }
  return out;
}

function buildMime(opts: {
  from: string;
  to: string[];
  cc: string[];
  subject: string;
  bodyHtml: string;
  inReplyTo: string;
  references: string;
  attachments: { filename: string; mimeType: string; dataBase64: string }[];
}): string {
  const boundary = "mixed_" + crypto.randomUUID().replace(/-/g, "");
  const altBoundary = "alt_" + crypto.randomUUID().replace(/-/g, "");
  const lines: string[] = [];
  lines.push(`From: ${opts.from}`);
  lines.push(`To: ${opts.to.join(", ")}`);
  if (opts.cc.length) lines.push(`Cc: ${opts.cc.join(", ")}`);
  lines.push(`Subject: =?UTF-8?B?${btoa(unescape(encodeURIComponent(opts.subject)))}?=`);
  lines.push("MIME-Version: 1.0");
  if (opts.inReplyTo) lines.push(`In-Reply-To: ${opts.inReplyTo}`);
  if (opts.references) lines.push(`References: ${opts.references}`);
  lines.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);
  lines.push("");
  lines.push(`--${boundary}`);
  lines.push(`Content-Type: multipart/alternative; boundary="${altBoundary}"`);
  lines.push("");
  lines.push(`--${altBoundary}`);
  lines.push("Content-Type: text/plain; charset=UTF-8");
  lines.push("Content-Transfer-Encoding: quoted-printable");
  lines.push("");
  lines.push(encodeQuotedPrintableHebrew(opts.bodyHtml.replace(/<[^>]+>/g, " ")));
  lines.push(`--${altBoundary}`);
  lines.push("Content-Type: text/html; charset=UTF-8");
  lines.push("Content-Transfer-Encoding: quoted-printable");
  lines.push("");
  lines.push(encodeQuotedPrintableHebrew(opts.bodyHtml));
  lines.push(`--${altBoundary}--`);

  for (const att of opts.attachments) {
    const bytes = decodeB64(att.dataBase64);
    lines.push(`--${boundary}`);
    lines.push(`Content-Type: ${att.mimeType}; name="${att.filename}"`);
    lines.push("Content-Transfer-Encoding: base64");
    lines.push(`Content-Disposition: attachment; filename="${att.filename}"`);
    lines.push("");
    lines.push(b64StdEncode(bytes).replace(/.{76}/g, "$&\r\n"));
  }
  lines.push(`--${boundary}--`);
  return lines.join("\r\n");
}

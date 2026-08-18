// Shared helpers for the Gmail read-only pilot.
// Never call Gmail send / delete / modify from these functions.

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

export const GMAIL_READONLY_SCOPE = "https://www.googleapis.com/auth/gmail.readonly";

export const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

export function json(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}

export function serviceClient() {
  const url = Deno.env.get("SUPABASE_URL");
  const key = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
  if (!url || !key) throw new Error("SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY is missing");
  return createClient(url, key);
}

type ServiceClient = ReturnType<typeof serviceClient>;

export async function requireAdmin(
  sb: ServiceClient,
  userId: unknown,
): Promise<{ user: { id: string; role: string; name: string | null } } | { error: Response }> {
  const id = String(userId || "").trim();
  if (!id) return { error: json({ error: "missing user_id" }, 400) };
  const { data, error } = await sb.from("users").select("id,role,name").eq("id", id).maybeSingle();
  if (error) return { error: json({ error: error.message }, 500) };
  if (!data || data.role !== "admin") return { error: json({ error: "admin only" }, 403) };
  return { user: data as { id: string; role: string; name: string | null } };
}

export function oauthEnv(): { clientId: string; clientSecret: string; redirectUri: string } | { error: Response } {
  const clientId = (Deno.env.get("GMAIL_CLIENT_ID") || "").trim();
  const clientSecret = (Deno.env.get("GMAIL_CLIENT_SECRET") || "").trim();
  const redirectUri = (Deno.env.get("GMAIL_REDIRECT_URI") || "").trim();
  if (!clientId || !clientSecret || !redirectUri) {
    return { error: json({ error: "Gmail OAuth secrets are not configured" }, 500) };
  }
  return { clientId, clientSecret, redirectUri };
}

type ConnectionRow = {
  user_id: string;
  google_email: string | null;
  access_token: string;
  refresh_token: string | null;
  token_expiry: string | null;
  scope: string | null;
};

async function refreshAccessToken(
  sb: ServiceClient,
  row: ConnectionRow,
  env: { clientId: string; clientSecret: string; redirectUri: string },
): Promise<{ token: string } | { error: Response }> {
  if (!row.refresh_token) return { error: json({ error: "missing refresh_token — reconnect Gmail" }, 401) };
  const body = new URLSearchParams({
    client_id: env.clientId,
    client_secret: env.clientSecret,
    refresh_token: row.refresh_token,
    grant_type: "refresh_token",
  });
  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const data = await res.json();
  if (!res.ok || !data.access_token) {
    return { error: json({ error: "token refresh failed", details: data }, 401) };
  }
  const expiresIn = Number(data.expires_in) || 3600;
  const tokenExpiry = new Date(Date.now() + (expiresIn - 60) * 1000).toISOString();
  const { error } = await sb.from("gmail_connections").update({
    access_token: data.access_token,
    token_expiry: tokenExpiry,
    updated_at: new Date().toISOString(),
  }).eq("user_id", row.user_id);
  if (error) return { error: json({ error: error.message }, 500) };
  return { token: data.access_token as string };
}

export async function getValidAccessToken(
  sb: ServiceClient,
  userId: string,
): Promise<{ token: string; email: string | null } | { error: Response }> {
  const env = oauthEnv();
  if ("error" in env) return env;
  const { data, error } = await sb.from("gmail_connections").select(
    "user_id,google_email,access_token,refresh_token,token_expiry,scope",
  ).eq("user_id", userId).maybeSingle();
  if (error) return { error: json({ error: error.message }, 500) };
  if (!data) return { error: json({ error: "gmail not connected" }, 404) };
  const row = data as ConnectionRow;
  const expiry = row.token_expiry ? Date.parse(row.token_expiry) : 0;
  if (!row.access_token || (expiry && expiry <= Date.now())) {
    const refreshed = await refreshAccessToken(sb, row, env);
    if ("error" in refreshed) return refreshed;
    return { token: refreshed.token, email: row.google_email };
  }
  return { token: row.access_token, email: row.google_email };
}

export async function gmailGet(
  accessToken: string,
  path: string,
): Promise<{ ok: true; data: unknown } | { ok: false; status: number; data: unknown }> {
  const url = path.startsWith("https://")
    ? path
    : `https://gmail.googleapis.com/gmail/v1/${path.replace(/^\/+/, "")}`;
  const res = await fetch(url, {
    method: "GET",
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const text = await res.text();
  let data: unknown = text;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    data = text;
  }
  if (!res.ok) return { ok: false, status: res.status, data };
  return { ok: true, data };
}

export function headerMap(payload: { headers?: { name?: string; value?: string }[] } | null | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  for (const h of payload?.headers || []) {
    if (!h?.name) continue;
    out[h.name.toLowerCase()] = h.value || "";
  }
  return out;
}

export function b64UrlDecode(data: string): string {
  const pad = data.replace(/-/g, "+").replace(/_/g, "/");
  const bin = atob(pad);
  const bytes = Uint8Array.from(bin, (c) => c.charCodeAt(0));
  return new TextDecoder("utf-8").decode(bytes);
}

type MimePart = {
  mimeType?: string;
  filename?: string;
  body?: { data?: string; size?: number; attachmentId?: string };
  parts?: MimePart[];
  headers?: { name?: string; value?: string }[];
};

export function flattenParts(part: MimePart | undefined, acc: MimePart[] = []): MimePart[] {
  if (!part) return acc;
  acc.push(part);
  for (const child of part.parts || []) flattenParts(child, acc);
  return acc;
}

export function extractBodies(payload: MimePart | undefined): { text: string; html: string } {
  const parts = flattenParts(payload);
  let text = "";
  let html = "";
  for (const p of parts) {
    const data = p.body?.data;
    if (!data) continue;
    if (p.mimeType === "text/plain" && !text) text = b64UrlDecode(data);
    if (p.mimeType === "text/html" && !html) html = b64UrlDecode(data);
  }
  if (!text && !html && payload?.body?.data && !payload.filename) {
    const raw = b64UrlDecode(payload.body.data);
    if ((payload.mimeType || "").includes("html")) html = raw;
    else text = raw;
  }
  return { text, html };
}

export function extractAttachments(payload: MimePart | undefined): {
  filename: string;
  mimeType: string;
  size: number;
  attachmentId: string;
}[] {
  return flattenParts(payload)
    .filter((p) => p.body?.attachmentId && (p.filename || p.body.attachmentId))
    .map((p) => ({
      filename: p.filename || "attachment",
      mimeType: p.mimeType || "application/octet-stream",
      size: Number(p.body?.size) || 0,
      attachmentId: p.body!.attachmentId!,
    }));
}

/** True when a metadata/full payload has a named file or attachmentId (list-row paperclip). */
export function payloadHasAttachments(payload: MimePart | undefined): boolean {
  return flattenParts(payload).some((p) => {
    if ((p.filename || "").trim()) return true;
    if (p.body?.attachmentId) return true;
    return false;
  });
}

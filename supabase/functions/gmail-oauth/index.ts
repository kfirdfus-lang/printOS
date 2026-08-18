// Gmail pilot — OAuth start / callback / status / disconnect.
// Read-only scope only. Tokens stay in gmail_connections (service role).

import { rejectDisallowedInternalOrigin } from "../_shared/cors.ts";
import {
  GMAIL_READONLY_SCOPE,
  corsHeaders,
  gmailGet,
  json,
  oauthEnv,
  requireAdmin,
  serviceClient,
} from "../_shared/gmail.ts";

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
  if (["send", "delete", "modify"].includes(action)) {
    return json({ error: "read-only pilot — send/delete/modify are not implemented" }, 400);
  }

  try {
    const sb = serviceClient();
    if (action === "start") return await handleStart(sb, body.user_id);
    if (action === "callback") return await handleCallback(sb, body);
    if (action === "status") return await handleStatus(sb, body.user_id);
    if (action === "disconnect") return await handleDisconnect(sb, body.user_id);
    return json({ error: "unknown action" }, 400);
  } catch (e) {
    return json({ error: e instanceof Error ? e.message : String(e) }, 500);
  }
});

async function handleStart(sb: ReturnType<typeof serviceClient>, userId: unknown) {
  const admin = await requireAdmin(sb, userId);
  if ("error" in admin) return admin.error;
  const env = oauthEnv();
  if ("error" in env) return env.error;

  await sb.from("gmail_oauth_states").delete().lt("expires_at", new Date().toISOString());

  const state = crypto.randomUUID();
  const expiresAt = new Date(Date.now() + 10 * 60 * 1000).toISOString();
  const { error } = await sb.from("gmail_oauth_states").insert({
    state,
    user_id: admin.user.id,
    expires_at: expiresAt,
  });
  if (error) return json({ error: error.message }, 500);

  const url = new URL("https://accounts.google.com/o/oauth2/v2/auth");
  url.searchParams.set("client_id", env.clientId);
  url.searchParams.set("redirect_uri", env.redirectUri);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("scope", GMAIL_READONLY_SCOPE);
  url.searchParams.set("access_type", "offline");
  url.searchParams.set("prompt", "consent");
  url.searchParams.set("include_granted_scopes", "false");
  url.searchParams.set("state", state);

  return json({ url: url.toString() });
}

async function handleCallback(sb: ReturnType<typeof serviceClient>, body: Record<string, unknown>) {
  const code = String(body.code || "").trim();
  const state = String(body.state || "").trim();
  if (!code || !state) return json({ error: "missing code or state" }, 400);

  const env = oauthEnv();
  if ("error" in env) return env.error;

  const { data: st, error: stErr } = await sb.from("gmail_oauth_states").select("state,user_id,expires_at")
    .eq("state", state).maybeSingle();
  if (stErr) return json({ error: stErr.message }, 500);
  if (!st) return json({ error: "invalid or expired state" }, 400);
  if (Date.parse(st.expires_at) <= Date.now()) {
    await sb.from("gmail_oauth_states").delete().eq("state", state);
    return json({ error: "oauth state expired" }, 400);
  }

  const admin = await requireAdmin(sb, st.user_id);
  if ("error" in admin) return admin.error;

  const tokenRes = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      code,
      client_id: env.clientId,
      client_secret: env.clientSecret,
      redirect_uri: env.redirectUri,
      grant_type: "authorization_code",
    }),
  });
  const tokenData = await tokenRes.json();
  if (!tokenRes.ok || !tokenData.access_token) {
    return json({ error: "token exchange failed", details: tokenData }, 400);
  }

  const granted = String(tokenData.scope || "");
  if (granted && !granted.split(/\s+/).includes(GMAIL_READONLY_SCOPE)) {
    return json({ error: "unexpected oauth scope" }, 400);
  }

  const profile = await gmailGet(tokenData.access_token, "users/me/profile");
  const googleEmail = profile.ok
    ? String((profile.data as { emailAddress?: string })?.emailAddress || "")
    : "";

  const expiresIn = Number(tokenData.expires_in) || 3600;
  const tokenExpiry = new Date(Date.now() + (expiresIn - 60) * 1000).toISOString();

  const { data: existing } = await sb.from("gmail_connections").select("refresh_token")
    .eq("user_id", admin.user.id).maybeSingle();
  const refreshToken = tokenData.refresh_token || existing?.refresh_token || null;
  if (!refreshToken) {
    return json({ error: "Google did not return a refresh token. Reconnect and grant access again." }, 400);
  }

  const now = new Date().toISOString();
  const { error: upErr } = await sb.from("gmail_connections").upsert({
    user_id: admin.user.id,
    google_email: googleEmail || null,
    access_token: tokenData.access_token,
    refresh_token: refreshToken,
    token_expiry: tokenExpiry,
    scope: granted || GMAIL_READONLY_SCOPE,
    connected_at: now,
    updated_at: now,
  }, { onConflict: "user_id" });
  if (upErr) return json({ error: upErr.message }, 500);

  await sb.from("gmail_oauth_states").delete().eq("state", state);
  return json({ connected: true, email: googleEmail });
}

async function handleStatus(sb: ReturnType<typeof serviceClient>, userId: unknown) {
  const admin = await requireAdmin(sb, userId);
  if ("error" in admin) return admin.error;
  const { data, error } = await sb.from("gmail_connections")
    .select("google_email,connected_at,scope")
    .eq("user_id", admin.user.id)
    .maybeSingle();
  if (error) return json({ error: error.message }, 500);
  return json({
    connected: !!data,
    email: data?.google_email || null,
    connected_at: data?.connected_at || null,
  });
}

async function handleDisconnect(sb: ReturnType<typeof serviceClient>, userId: unknown) {
  const admin = await requireAdmin(sb, userId);
  if ("error" in admin) return admin.error;
  const { data } = await sb.from("gmail_connections")
    .select("access_token,refresh_token")
    .eq("user_id", admin.user.id)
    .maybeSingle();
  const token = data?.refresh_token || data?.access_token;
  if (token) {
    try {
      await fetch("https://oauth2.googleapis.com/revoke", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({ token }),
      });
    } catch {
      // ignore revoke failures — local row is still removed
    }
  }
  const { error } = await sb.from("gmail_connections").delete().eq("user_id", admin.user.id);
  if (error) return json({ error: error.message }, 500);
  return json({ connected: false });
}

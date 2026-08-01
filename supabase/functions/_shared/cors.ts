// Shared CORS / origin allow-list helpers for Edge Functions.
//
// Two surfaces:
//  - 'public'   (mockup-generate / mockup-finalize): callable from the
//    e-commerce site and the internal board. Allowed origins = built-in
//    defaults + the ALLOWED_ORIGINS_PUBLIC secret (comma separated).
//  - 'internal' (functions only the internal board calls): allowed origins
//    come from the ALLOWED_ORIGINS_INTERNAL secret.
//
// Enforcement is opt-in per surface: until the corresponding secret is set,
// browser origins are not rejected (fail-open) so the internal board keeps
// working before configuration. Requests without an Origin header
// (curl / server-to-server) cannot be identified by CORS and always pass —
// JWT verification and rate limiting are the protections for those.

const PUBLIC_DEFAULT_ORIGINS = [
  'http://localhost:8000',
  'http://localhost:3000',
  'https://natalie-print.co.il',
  'https://www.natalie-print.co.il',
]

const BASE_HEADERS: Record<string, string> = {
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
  'Vary': 'Origin',
}

function parseOriginList(envName: string): string[] {
  return (Deno.env.get(envName) || '')
    .split(',')
    .map((s) => s.trim().replace(/\/+$/, ''))
    .filter(Boolean)
}

export interface CorsResult {
  /** Headers to spread into every response of this request. */
  headers: Record<string, string>
  /** Non-null when the request should be answered immediately (preflight or blocked origin). */
  earlyResponse: Response | null
}

export function applyCors(req: Request, surface: 'public' | 'internal'): CorsResult {
  const envList = parseOriginList(
    surface === 'public' ? 'ALLOWED_ORIGINS_PUBLIC' : 'ALLOWED_ORIGINS_INTERNAL',
  )
  const enforced = envList.length > 0
  const allowed = new Set(
    surface === 'public' ? [...PUBLIC_DEFAULT_ORIGINS, ...envList] : envList,
  )

  const origin = (req.headers.get('origin') || '').replace(/\/+$/, '')
  const blocked = enforced && origin !== '' && !allowed.has(origin)

  const headers: Record<string, string> = {
    ...BASE_HEADERS,
    'Access-Control-Allow-Origin': blocked ? 'null' : origin || '*',
  }

  if (blocked) {
    console.warn(`[cors] blocked origin "${origin}" on ${surface} surface`)
    return {
      headers,
      earlyResponse: new Response(JSON.stringify({ error: 'Origin not allowed' }), {
        status: 403,
        headers: { ...headers, 'Content-Type': 'application/json' },
      }),
    }
  }

  if (req.method === 'OPTIONS') {
    return { headers, earlyResponse: new Response('ok', { headers }) }
  }

  return { headers, earlyResponse: null }
}

/**
 * Guard for internal-only functions (board-facing). Returns a 403 Response when
 * the browser Origin is not in ALLOWED_ORIGINS_INTERNAL; null otherwise.
 * No-Origin callers (cron, curl) and all callers while the secret is unset pass
 * through, so nothing breaks before configuration.
 */
export function rejectDisallowedInternalOrigin(req: Request): Response | null {
  const { earlyResponse } = applyCors(req, 'internal')
  return earlyResponse && earlyResponse.status === 403 ? earlyResponse : null
}

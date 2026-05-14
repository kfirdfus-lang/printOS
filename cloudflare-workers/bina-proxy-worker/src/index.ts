type Env = {
  QUOTAGUARD_URL?: string;
  BINA_TOKEN?: string;
};

function bool(v: unknown): boolean {
  return typeof v === "string" && v.length > 0;
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const qg = env.QUOTAGUARD_URL ?? "";
    const token = env.BINA_TOKEN ?? "";

    const payload = {
      ok: false,
      message: "Diagnostic mode: worker returned env metadata for debugging.",
      request: {
        method: request.method,
        path: new URL(request.url).pathname,
      },
      diagnostics: {
        quotaguard_defined: bool(qg),
        quotaguard_has_valid_scheme: qg.startsWith("http://") || qg.startsWith("https://"),
        bina_token_defined: bool(token),
        quotaguard_url_length: qg.length,
      },
    };

    return new Response(JSON.stringify(payload), {
      status: 500,
      headers: { "content-type": "application/json; charset=UTF-8" },
    });
  },
};

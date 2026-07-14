// @ts-ignore
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";

serve(async (req) => {
  const CORS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  };

  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });

  const results: Record<string, unknown> = {
    timestamp: new Date().toISOString(),
    checks: {} as Record<string, unknown>,
  };
  const checks = results.checks as Record<string, unknown>;

  const resendKey = Deno.env.get("RESEND_API_KEY");
  checks.api_key_exists = !!resendKey;
  checks.api_key_length = resendKey ? resendKey.length : 0;
  checks.api_key_starts_with = resendKey ? resendKey.substring(0, 5) : null;

  if (!resendKey) {
    return new Response(JSON.stringify({ ...results, error: "RESEND_API_KEY not set" }), {
      headers: { ...CORS, "Content-Type": "application/json" },
    });
  }

  try {
    const testResponse = await fetch("https://api.resend.com/domains", {
      headers: {
        Authorization: `Bearer ${resendKey}`,
      },
    });

    checks.domains_api_status = testResponse.status;
    checks.domains_api_ok = testResponse.ok;

    if (testResponse.ok) {
      checks.domains = await testResponse.json();
    } else {
      checks.domains_api_error = await testResponse.text();
    }
  } catch (err) {
    checks.domains_api_exception = err instanceof Error ? err.message : String(err);
  }

  try {
    const sendResponse = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${resendKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        from: "דפוס נטלי <orders@natalie-print.com>",
        to: ["kfir.dfus@gmail.com"],
        subject: "🔬 בדיקת אבחון PrintOS - " + new Date().toLocaleString("he-IL"),
        html: `
          <div dir="rtl" style="font-family: Arial;">
            <h2>בדיקת אבחון המערכת</h2>
            <p>אם קיבלת את המייל הזה - המערכת יכולה לשלוח מיילים.</p>
            <p>זמן שליחה: ${new Date().toISOString()}</p>
          </div>
        `,
      }),
    });

    checks.send_status = sendResponse.status;
    checks.send_ok = sendResponse.ok;

    const sendData = await sendResponse.text();
    try {
      checks.send_response = JSON.parse(sendData);
    } catch {
      checks.send_response_raw = sendData;
    }
  } catch (err) {
    checks.send_exception = err instanceof Error ? err.message : String(err);
  }

  return new Response(JSON.stringify(results, null, 2), {
    headers: { ...CORS, "Content-Type": "application/json" },
  });
});

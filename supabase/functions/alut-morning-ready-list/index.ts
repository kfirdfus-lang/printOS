// @ts-ignore
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
// @ts-ignore
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS_HEADERS });

  try {
    const supabase = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    const { data: items } = await supabase
      .from("alut_order_items")
      .select(`
        id, calendar_type, quantity, status,
        alut_orders (
          order_number,
          company_name,
          contact_name,
          contact_phone,
          delivery_address,
          is_pickup
        )
      `)
      .in("status", ["ready_to_ship", "ready_for_pickup"])
      .order("order_id", { ascending: true });

    if (!items || items.length === 0) {
      return new Response(
        JSON.stringify({ success: true, message: "No ready items - no email sent" }),
        { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
      );
    }

    const shipments = items.filter((i) => i.status === "ready_to_ship");
    const pickups = items.filter((i) => i.status === "ready_for_pickup");

    const typeLabels: Record<string, string> = {
      hard: "קשיח",
      duplex: "דופלקס",
      wall: "קיר",
    };

    function renderTable(list: any[], title: string, isPickup: boolean) {
      if (list.length === 0) return "";

      const rows = list.map((item) => {
        const o = item.alut_orders || {};
        return `
          <tr>
            <td>${o.order_number ?? "?"}</td>
            <td>${o.company_name ?? "—"}</td>
            <td>${Number(item.quantity || 0).toLocaleString("he-IL")} ${typeLabels[item.calendar_type] || item.calendar_type}</td>
            <td>${isPickup ? "🎒 איסוף עצמי" : (o.delivery_address || "—")}</td>
            <td>${o.contact_name ? `${o.contact_name}${o.contact_phone ? ` • ${o.contact_phone}` : ""}` : "—"}</td>
          </tr>
        `;
      }).join("");

      return `
        <div class="section">
          <div class="section-title">${title} (${list.length})</div>
          <table>
            <thead>
              <tr>
                <th>הזמנה</th>
                <th>חברה</th>
                <th>פריט</th>
                <th>${isPickup ? "סוג" : "כתובת"}</th>
                <th>איש קשר</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      `;
    }

    const now = new Date();
    const dateLabel = now.toLocaleDateString("he-IL", {
      timeZone: "Asia/Jerusalem",
      weekday: "long",
      day: "numeric",
      month: "long",
      year: "numeric",
    });
    const dateShort = now.toLocaleDateString("he-IL", { timeZone: "Asia/Jerusalem" });

    const html = `<!DOCTYPE html>
<html dir="rtl" lang="he">
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Heebo', Arial, sans-serif; background: #f9fafb; padding: 20px; direction: rtl; margin: 0; }
    .container { max-width: 800px; margin: 0 auto; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
    .header { background: linear-gradient(135deg, #0d9488 0%, #14b8a6 100%); color: white; padding: 24px 30px; }
    .header h1 { margin: 0; font-size: 22px; }
    .header p { margin: 6px 0 0 0; opacity: 0.9; font-size: 14px; }
    .section { padding: 20px 30px; border-bottom: 1px solid #f3f4f6; }
    .section:last-child { border-bottom: none; }
    .section-title { font-size: 16px; font-weight: 700; color: #0d9488; margin-bottom: 14px; }
    table { width: 100%; border-collapse: collapse; }
    th { background: #f3f4f6; padding: 10px; text-align: right; font-size: 13px; color: #374151; }
    td { padding: 10px; border-bottom: 1px solid #f3f4f6; font-size: 13px; color: #1f2937; }
    tr:last-child td { border-bottom: none; }
    .footer { background: #f9fafb; padding: 15px 30px; text-align: center; font-size: 12px; color: #6b7280; }
    .summary { display: flex; gap: 20px; padding: 20px 30px; background: #f0fdfa; }
    .summary-item { flex: 1; text-align: center; }
    .summary-value { font-size: 28px; font-weight: 900; color: #0d9488; }
    .summary-label { font-size: 12px; color: #6b7280; margin-top: 4px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>☀️ בוקר טוב קארין - רשימת מוכנים</h1>
      <p>${dateLabel}</p>
    </div>
    <div class="summary">
      <div class="summary-item">
        <div class="summary-value">${shipments.length}</div>
        <div class="summary-label">🚚 מוכן למשלוח</div>
      </div>
      <div class="summary-item">
        <div class="summary-value">${pickups.length}</div>
        <div class="summary-label">🎒 מוכן לאיסוף עצמי</div>
      </div>
      <div class="summary-item">
        <div class="summary-value">${items.length}</div>
        <div class="summary-label">📦 סה"כ</div>
      </div>
    </div>
    ${renderTable(shipments, "🚚 מוכן למשלוח", false)}
    ${renderTable(pickups, "🎒 מוכן לאיסוף עצמי", true)}
    <div class="footer">🖨️ נטלי פרינט • רשימה יומית • שעה 9:00 בבוקר</div>
  </div>
</body>
</html>`;

    const emailResponse = await supabase.functions.invoke("send-email", {
      body: {
        to: ["karinha@alut.org.il"],
        subject: `☀️ רשימת מוכנים - ${dateShort} (${items.length} פריטים)`,
        html,
        from: "PrintOS - נטלי פרינט <orders@natalie-print.com>",
      },
    });

    if (emailResponse.error) throw emailResponse.error;

    return new Response(
      JSON.stringify({
        success: true,
        stats: {
          total: items.length,
          shipments: shipments.length,
          pickups: pickups.length,
        },
      }),
      { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
    );
  } catch (error: any) {
    return new Response(
      JSON.stringify({ error: error.message || String(error) }),
      { status: 500, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
    );
  }
});

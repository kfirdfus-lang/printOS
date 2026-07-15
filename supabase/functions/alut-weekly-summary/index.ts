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

    const now = new Date();
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    const todayStr = now.toISOString().split("T")[0];

    // 1. New orders in the last week
    const { data: newOrders } = await supabase
      .from("alut_orders")
      .select("id, order_number, company_name, created_at")
      .gte("created_at", weekAgo.toISOString())
      .order("created_at", { ascending: false });

    // 2. Items delivered in the last week
    const { data: deliveredItems } = await supabase
      .from("alut_order_items")
      .select("id, quantity, calendar_type, delivered_at, alut_orders(order_number, company_name)")
      .gte("delivered_at", weekAgo.toISOString())
      .not("delivered_at", "is", null);

    // 3. Overdue items (deadline passed, not delivered/cancelled)
    const { data: overdueItems } = await supabase
      .from("alut_order_items")
      .select("id, quantity, calendar_type, delivery_deadline, status, alut_orders(order_number, company_name)")
      .not("delivery_deadline", "is", null)
      .lt("delivery_deadline", todayStr)
      .not("status", "in", "(delivered,cancelled)")
      .order("delivery_deadline", { ascending: true });

    // 4. Currently active items
    const { data: activeItems } = await supabase
      .from("alut_order_items")
      .select("id, status")
      .not("status", "in", "(delivered,cancelled)");

    const typeLabels: Record<string, string> = {
      hard: "קשיח",
      duplex: "דופלקס",
      wall: "קיר",
    };

    const newOrdersRows = (newOrders || []).map((o: Record<string, unknown>) =>
      `<tr><td>${o.order_number}</td><td>${o.company_name}</td><td>${new Date(String(o.created_at)).toLocaleDateString("he-IL")}</td></tr>`
    ).join("");

    const deliveredRows = (deliveredItems || []).map((i: Record<string, unknown>) => {
      const order = i.alut_orders as Record<string, unknown> | null;
      return `<tr><td>${order?.order_number ?? "?"}</td><td>${order?.company_name ?? "—"}</td><td>${Number(i.quantity) || 0} ${typeLabels[String(i.calendar_type)] || i.calendar_type}</td></tr>`;
    }).join("");

    const overdueRows = (overdueItems || []).map((i: Record<string, unknown>) => {
      const order = i.alut_orders as Record<string, unknown> | null;
      const deadline = new Date(String(i.delivery_deadline));
      const daysLate = Math.ceil((now.getTime() - deadline.getTime()) / (1000 * 60 * 60 * 24));
      return `<tr><td>${order?.order_number ?? "?"}</td><td>${order?.company_name ?? "—"}</td><td>${Number(i.quantity) || 0} ${typeLabels[String(i.calendar_type)] || i.calendar_type}</td><td style="color: #dc2626; font-weight: bold;">${daysLate} ימים</td></tr>`;
    }).join("");

    const html = `
      <!DOCTYPE html>
      <html dir="rtl" lang="he">
      <head>
        <meta charset="UTF-8">
        <style>
          body { font-family: Arial, sans-serif; direction: rtl; padding: 20px; background: #f9fafb; }
          .container { max-width: 700px; margin: 0 auto; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
          .header { background: linear-gradient(135deg, #0d9488 0%, #14b8a6 100%); color: white; padding: 20px 30px; }
          .header h1 { margin: 0; font-size: 22px; }
          .stat-row { display: flex; padding: 15px 30px; border-bottom: 1px solid #f3f4f6; }
          .stat { flex: 1; text-align: center; }
          .stat-num { font-size: 28px; font-weight: 900; color: #0d9488; }
          .stat-label { font-size: 13px; color: #6b7280; }
          .section { padding: 20px 30px; border-bottom: 1px solid #f3f4f6; }
          .section h2 { margin: 0 0 15px 0; font-size: 16px; color: #0d9488; }
          table { width: 100%; border-collapse: collapse; }
          th { background: #f3f4f6; padding: 8px; text-align: right; font-size: 13px; }
          td { padding: 8px; border-bottom: 1px solid #f3f4f6; font-size: 13px; }
          .empty { color: #9ca3af; font-style: italic; text-align: center; padding: 15px; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>📊 סיכום שבועי - אלוט</h1>
            <p style="margin: 8px 0 0 0; font-size: 13px; opacity: 0.9;">
              ${weekAgo.toLocaleDateString("he-IL")} - ${now.toLocaleDateString("he-IL")}
            </p>
          </div>

          <div class="stat-row">
            <div class="stat">
              <div class="stat-num">${newOrders?.length || 0}</div>
              <div class="stat-label">הזמנות חדשות</div>
            </div>
            <div class="stat">
              <div class="stat-num">${deliveredItems?.length || 0}</div>
              <div class="stat-label">פריטים שנמסרו</div>
            </div>
            <div class="stat">
              <div class="stat-num" style="color: ${overdueItems?.length ? "#dc2626" : "#0d9488"}">${overdueItems?.length || 0}</div>
              <div class="stat-label">באיחור</div>
            </div>
            <div class="stat">
              <div class="stat-num">${activeItems?.length || 0}</div>
              <div class="stat-label">פעילות</div>
            </div>
          </div>

          <div class="section">
            <h2>🆕 הזמנות חדשות השבוע</h2>
            ${newOrdersRows ? `<table><thead><tr><th>הזמנה</th><th>חברה</th><th>תאריך</th></tr></thead><tbody>${newOrdersRows}</tbody></table>` : '<div class="empty">אין הזמנות חדשות השבוע</div>'}
          </div>

          <div class="section">
            <h2>✅ נמסרו השבוע</h2>
            ${deliveredRows ? `<table><thead><tr><th>הזמנה</th><th>חברה</th><th>פריט</th></tr></thead><tbody>${deliveredRows}</tbody></table>` : '<div class="empty">אין פריטים שנמסרו השבוע</div>'}
          </div>

          <div class="section">
            <h2>⚠️ פריטים באיחור</h2>
            ${overdueRows ? `<table><thead><tr><th>הזמנה</th><th>חברה</th><th>פריט</th><th>איחור</th></tr></thead><tbody>${overdueRows}</tbody></table>` : '<div class="empty">אין פריטים באיחור 🎉</div>'}
          </div>

          <div style="padding: 15px 30px; background: #f9fafb; text-align: center; color: #6b7280; font-size: 12px;">
            סיכום שבועי אוטומטי מ-PrintOS
          </div>
        </div>
      </body>
      </html>
    `;

    const recipients = ["kfir.dfus@gmail.com", "natalie.zem@gmail.com"];
    const emailResponse = await supabase.functions.invoke("send-email", {
      body: {
        to: recipients,
        subject: `📊 סיכום שבועי - אלוט (${now.toLocaleDateString("he-IL")})`,
        html,
        from: "PrintOS - סיכומים <orders@natalie-print.com>",
      },
    });

    if (emailResponse.error) throw emailResponse.error;

    return new Response(
      JSON.stringify({
        success: true,
        stats: {
          new_orders: newOrders?.length || 0,
          delivered: deliveredItems?.length || 0,
          overdue: overdueItems?.length || 0,
          active: activeItems?.length || 0,
        },
        recipients,
      }),
      { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
    );
  } catch (error) {
    return new Response(
      JSON.stringify({ error: error instanceof Error ? error.message : String(error) }),
      { status: 500, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
    );
  }
});

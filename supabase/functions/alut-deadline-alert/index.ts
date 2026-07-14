// @ts-ignore
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
// @ts-ignore
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function addBusinessDays(startDate: Date, days: number): Date {
  const result = new Date(startDate);
  let added = 0;
  while (added < days) {
    result.setDate(result.getDate() + 1);
    const dayOfWeek = result.getDay();
    if (dayOfWeek !== 5 && dayOfWeek !== 6) {
      added++;
    }
  }
  return result;
}

serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS_HEADERS });

  try {
    const supabase = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    const today = new Date();
    const targetDate = addBusinessDays(today, 5);
    const targetDateStr = targetDate.toISOString().split("T")[0];
    const todayStr = today.toISOString().split("T")[0];

    const { data: items, error } = await supabase
      .from("alut_order_items")
      .select(`
        id, calendar_type, quantity, delivery_deadline, status,
        alut_orders (order_number, company_name, contact_name)
      `)
      .not("delivery_deadline", "is", null)
      .gte("delivery_deadline", todayStr)
      .lte("delivery_deadline", targetDateStr)
      .not("status", "in", "(shipped,delivered,cancelled)")
      .order("delivery_deadline", { ascending: true });

    if (error) throw error;

    if (!items || items.length === 0) {
      return new Response(
        JSON.stringify({ success: true, message: "No items with upcoming deadlines", items_count: 0 }),
        { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
      );
    }

    const typeLabels: Record<string, string> = {
      hard: "קשיח",
      duplex: "דופלקס",
      wall: "קיר",
    };

    const itemRows = items.map((item: Record<string, unknown>) => {
      const order = item.alut_orders as Record<string, unknown> | null;
      const deadline = new Date(String(item.delivery_deadline));
      const daysLeft = Math.ceil((deadline.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
      const deadlineStr = deadline.toLocaleDateString("he-IL");
      const qty = Number(item.quantity) || 0;

      return `
        <tr>
          <td style="padding:8px;border:1px solid #e5e7eb;">${order?.order_number ?? "?"}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;">${order?.company_name ?? "—"}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;">${typeLabels[String(item.calendar_type)] || item.calendar_type}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;">${qty.toLocaleString("he-IL")}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;">${deadlineStr}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;font-weight:bold;color:${daysLeft <= 2 ? "#dc2626" : "#f59e0b"};">
            ${daysLeft} ימים
          </td>
        </tr>
      `;
    }).join("");

    const html = `
      <div dir="rtl" style="font-family: Arial, sans-serif; max-width: 700px; margin: auto;">
        <div style="background: linear-gradient(135deg, #f59e0b 0%, #ef4444 100%); color: white; padding: 20px; border-radius: 12px 12px 0 0;">
          <h2 style="margin: 0;">⏰ התראת דדליין - אלוט</h2>
          <p style="margin: 8px 0 0 0;">יש ${items.length} פריטים שאספקתם ב-5 ימי עבודה הקרובים</p>
        </div>
        <table style="width: 100%; border-collapse: collapse; margin-top: 20px; background: white;">
          <thead>
            <tr style="background: #f3f4f6;">
              <th style="padding: 10px; border: 1px solid #e5e7eb; text-align: right;">הזמנה</th>
              <th style="padding: 10px; border: 1px solid #e5e7eb; text-align: right;">חברה</th>
              <th style="padding: 10px; border: 1px solid #e5e7eb; text-align: right;">סוג</th>
              <th style="padding: 10px; border: 1px solid #e5e7eb; text-align: right;">כמות</th>
              <th style="padding: 10px; border: 1px solid #e5e7eb; text-align: right;">דדליין</th>
              <th style="padding: 10px; border: 1px solid #e5e7eb; text-align: right;">נותרו</th>
            </tr>
          </thead>
          <tbody>${itemRows}</tbody>
        </table>
        <p style="color: #6b7280; font-size: 12px; margin-top: 30px;">
          התראה אוטומטית מ-PrintOS<br>
          תאריך: ${today.toLocaleDateString("he-IL")}
        </p>
      </div>
    `;

    const recipients = ["kfir.dfus@gmail.com", "natalie.zem@gmail.com"];
    const emailResponse = await supabase.functions.invoke("send-email", {
      body: {
        to: recipients,
        subject: `⏰ ${items.length} פריטי אלוט עם דדליין קרוב`,
        html,
        from: "PrintOS - התראות אלוט <orders@natalie-print.com>",
      },
    });

    if (emailResponse.error) throw emailResponse.error;

    return new Response(
      JSON.stringify({
        success: true,
        items_count: items.length,
        recipients,
        email: emailResponse.data,
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

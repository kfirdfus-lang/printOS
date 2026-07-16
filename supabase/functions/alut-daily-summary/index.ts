// @ts-ignore
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
// @ts-ignore
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function israelDateParts(d: Date) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Jerusalem",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(d);
  const get = (t: string) => parts.find((p) => p.type === t)?.value || "00";
  return {
    dateStr: `${get("year")}-${get("month")}-${get("day")}`,
    hour: Number(get("hour")),
    minute: Number(get("minute")),
  };
}

function addBusinessDays(date: Date, days: number): Date {
  const result = new Date(date);
  let added = 0;
  while (added < days) {
    result.setDate(result.getDate() + 1);
    const day = result.getDay();
    if (day !== 5 && day !== 6) added++;
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

    const now = new Date();
    const { dateStr: todayStr } = israelDateParts(now);
    const todayStart = new Date(`${todayStr}T00:00:00+03:00`);
    const tomorrowStart = new Date(todayStart.getTime() + 24 * 60 * 60 * 1000);

    const { data: allOrders } = await supabase
      .from("alut_orders")
      .select("id, order_number, company_name, total_price, is_pickup");

    const { data: allItems } = await supabase
      .from("alut_order_items")
      .select("id, calendar_type, quantity, status, delivery_deadline, order_id")
      .neq("status", "cancelled");

    const byType: Record<string, { units: number; orders: Set<string | number> }> = {
      hard: { units: 0, orders: new Set() },
      duplex: { units: 0, orders: new Set() },
      wall: { units: 0, orders: new Set() },
    };

    allItems?.forEach((item) => {
      const t = String(item.calendar_type || "hard");
      if (!byType[t]) byType[t] = { units: 0, orders: new Set() };
      byType[t].units += Number(item.quantity) || 0;
      byType[t].orders.add(item.order_id);
    });

    const totalUnits = (byType.hard?.units || 0) + (byType.duplex?.units || 0) + (byType.wall?.units || 0);
    const totalOrders = new Set(allItems?.map((i) => i.order_id)).size;

    const statusCounts = { inPrint: 0, inBinding: 0, inShipment: 0, delivered: 0 };
    allItems?.forEach((item) => {
      if (["new", "design_in_progress", "design_sent", "design_approved", "in_print"].includes(item.status)) {
        statusCounts.inPrint++;
      } else if (["at_davach", "at_itzik"].includes(item.status)) {
        statusCounts.inBinding++;
      } else if (["ready_to_ship", "shipped"].includes(item.status)) {
        statusCounts.inShipment++;
      } else if (item.status === "delivered") {
        statusCounts.delivered++;
      }
    });

    const { data: todayStatusChanges } = await supabase
      .from("alut_status_history")
      .select("to_status, item_id, order_id, created_at")
      .gte("created_at", todayStart.toISOString())
      .lt("created_at", tomorrowStart.toISOString());

    const todayCounts = {
      newOrders: 0,
      designApproved: 0,
      movedToPrint: 0,
      movedToShip: 0,
      delivered: 0,
      sentToDavach: 0,
      sentToItzik: 0,
    };

    const { data: newOrdersToday } = await supabase
      .from("alut_orders")
      .select("id")
      .gte("created_at", todayStart.toISOString())
      .lt("created_at", tomorrowStart.toISOString());

    todayCounts.newOrders = newOrdersToday?.length || 0;

    todayStatusChanges?.forEach((change) => {
      if (change.to_status === "design_approved") todayCounts.designApproved++;
      if (change.to_status === "in_print") todayCounts.movedToPrint++;
      if (change.to_status === "shipped") todayCounts.movedToShip++;
      if (change.to_status === "delivered") todayCounts.delivered++;
      if (change.to_status === "at_davach") todayCounts.sentToDavach++;
      if (change.to_status === "at_itzik") todayCounts.sentToItzik++;
    });

    const overdueItems = allItems?.filter((item) => {
      if (!item.delivery_deadline) return false;
      if (["delivered", "cancelled"].includes(item.status)) return false;
      return String(item.delivery_deadline) < todayStr;
    }) || [];

    const upcomingLimit = addBusinessDays(new Date(`${todayStr}T12:00:00`), 5);
    const upcomingLimitStr = upcomingLimit.toISOString().split("T")[0];
    const upcomingItems = allItems?.filter((item) => {
      if (!item.delivery_deadline) return false;
      if (["delivered", "cancelled"].includes(item.status)) return false;
      const deadline = String(item.delivery_deadline);
      return deadline >= todayStr && deadline <= upcomingLimitStr;
    }) || [];

    const html = buildDailySummaryHtml({
      date: now,
      todayStr,
      totalOrders,
      totalUnits,
      byType,
      statusCounts,
      todayCounts,
      overdueCount: overdueItems.length,
      upcomingCount: upcomingItems.length,
    });

    const emailResponse = await supabase.functions.invoke("send-email", {
      body: {
        to: [
          "kfir.dfus@gmail.com",
          "natalie.zem@gmail.com",
          "karinha@alut.org.il",
        ],
        subject: `📊 סיכום יומי אלוט - ${new Date(todayStr + "T12:00:00").toLocaleDateString("he-IL")}`,
        html,
        from: "PrintOS - אלוט <orders@natalie-print.com>",
      },
    });

    if (emailResponse.error) throw emailResponse.error;

    return new Response(
      JSON.stringify({
        success: true,
        summary: {
          totalOrders,
          totalUnits,
          todayCounts,
          overdue: overdueItems.length,
          upcoming: upcomingItems.length,
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

function buildDailySummaryHtml(data: any): string {
  const { date, todayStr, totalOrders, totalUnits, byType, statusCounts, todayCounts, overdueCount, upcomingCount } = data;
  const dateLabel = new Date(todayStr + "T12:00:00").toLocaleDateString("he-IL", {
    day: "numeric",
    month: "long",
    year: "numeric",
    weekday: "long",
  });
  const timeLabel = date.toLocaleTimeString("he-IL", {
    hour: "2-digit",
    minute: "2-digit",
    timeZone: "Asia/Jerusalem",
  });

  return `<!DOCTYPE html>
<html dir="rtl" lang="he">
<head>
  <meta charset="UTF-8">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Heebo:wght@400;500;700;900&display=swap');
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Heebo', sans-serif;
      background: linear-gradient(135deg, #f0fdfa 0%, #ecfeff 100%);
      padding: 20px;
      color: #1f2937;
    }
    .container {
      max-width: 700px;
      margin: 0 auto;
      background: white;
      border-radius: 16px;
      overflow: hidden;
      box-shadow: 0 10px 30px rgba(13, 148, 136, 0.15);
    }
    .header {
      background: linear-gradient(135deg, #0d9488 0%, #14b8a6 100%);
      color: white;
      padding: 30px;
      text-align: center;
    }
    .header h1 { font-size: 26px; font-weight: 900; margin-bottom: 8px; }
    .header p { opacity: 0.9; font-size: 15px; }
    .section { padding: 24px 30px; border-bottom: 1px solid #f3f4f6; }
    .section:last-child { border-bottom: none; }
    .section-title {
      font-size: 16px; font-weight: 700; color: #0d9488; margin-bottom: 16px;
      display: flex; align-items: center; gap: 8px;
    }
    .overview-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; }
    .overview-card { background: #f9fafb; padding: 16px; border-radius: 10px; border-right: 4px solid; }
    .overview-card.total { border-color: #0d9488; }
    .overview-card.units { border-color: #14b8a6; }
    .overview-label { font-size: 12px; color: #6b7280; margin-bottom: 4px; }
    .overview-value { font-size: 24px; font-weight: 900; color: #0d9488; }
    .types-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
    .type-card { background: #f9fafb; padding: 14px; border-radius: 10px; text-align: center; border-right: 4px solid; }
    .type-card.hard { border-color: #3b82f6; }
    .type-card.duplex { border-color: #22c55e; }
    .type-card.wall { border-color: #f97316; }
    .type-label { font-size: 12px; color: #6b7280; margin-bottom: 6px; }
    .type-units { font-size: 20px; font-weight: 900; color: #1f2937; }
    .type-orders { font-size: 11px; color: #9ca3af; margin-top: 4px; }
    .status-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; }
    .status-card { background: #f9fafb; padding: 12px; border-radius: 8px; text-align: center; }
    .status-card.print { background: #fef3c7; }
    .status-card.binding { background: #f3e8ff; }
    .status-card.shipment { background: #dbeafe; }
    .status-card.done { background: #d1fae5; }
    .status-label { font-size: 11px; color: #6b7280; margin-bottom: 4px; }
    .status-value { font-size: 20px; font-weight: 900; }
    .status-card.print .status-value { color: #d97706; }
    .status-card.binding .status-value { color: #7c3aed; }
    .status-card.shipment .status-value { color: #2563eb; }
    .status-card.done .status-value { color: #059669; }
    .today-list { background: #f9fafb; border-radius: 10px; padding: 16px; }
    .today-item {
      display: flex; justify-content: space-between; padding: 8px 0;
      border-bottom: 1px dashed #e5e7eb; font-size: 14px;
    }
    .today-item:last-child { border-bottom: none; }
    .today-item-label { color: #4b5563; display: flex; gap: 8px; align-items: center; }
    .today-item-value { font-weight: 700; color: #0d9488; }
    .today-empty { text-align: center; color: #9ca3af; padding: 20px; font-style: italic; }
    .alerts-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; }
    .alert-card { padding: 16px; border-radius: 10px; text-align: center; }
    .alert-overdue { background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%); border: 1px solid #fca5a5; }
    .alert-upcoming { background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); border: 1px solid #fcd34d; }
    .alert-label { font-size: 12px; color: #6b7280; margin-bottom: 6px; }
    .alert-value { font-size: 28px; font-weight: 900; }
    .alert-overdue .alert-value { color: #dc2626; }
    .alert-upcoming .alert-value { color: #d97706; }
    .footer { background: #f9fafb; padding: 16px 30px; text-align: center; font-size: 12px; color: #6b7280; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📊 סיכום יומי אלוט</h1>
      <p>${dateLabel}</p>
      <p style="font-size: 13px; margin-top: 4px;">${timeLabel}</p>
    </div>
    <div class="section">
      <div class="section-title">🎯 תמונה כללית של הפרויקט</div>
      <div class="overview-grid">
        <div class="overview-card total">
          <div class="overview-label">סה"כ הזמנות</div>
          <div class="overview-value">${totalOrders}</div>
        </div>
        <div class="overview-card units">
          <div class="overview-label">סה"כ יחידות</div>
          <div class="overview-value">${Number(totalUnits).toLocaleString("he-IL")}</div>
        </div>
      </div>
    </div>
    <div class="section">
      <div class="section-title">📦 חלוקה לפי סוג לוח</div>
      <div class="types-grid">
        <div class="type-card hard">
          <div class="type-label">🔷 קשיח</div>
          <div class="type-units">${(byType.hard?.units || 0).toLocaleString("he-IL")}</div>
          <div class="type-orders">${byType.hard?.orders?.size || 0} הזמנות</div>
        </div>
        <div class="type-card duplex">
          <div class="type-label">📗 דופלקס</div>
          <div class="type-units">${(byType.duplex?.units || 0).toLocaleString("he-IL")}</div>
          <div class="type-orders">${byType.duplex?.orders?.size || 0} הזמנות</div>
        </div>
        <div class="type-card wall">
          <div class="type-label">🖼️ קיר</div>
          <div class="type-units">${(byType.wall?.units || 0).toLocaleString("he-IL")}</div>
          <div class="type-orders">${byType.wall?.orders?.size || 0} הזמנות</div>
        </div>
      </div>
    </div>
    <div class="section">
      <div class="section-title">📈 מצב פריטים כרגע</div>
      <div class="status-grid">
        <div class="status-card print">
          <div class="status-label">בהדפסה</div>
          <div class="status-value">${statusCounts.inPrint}</div>
        </div>
        <div class="status-card binding">
          <div class="status-label">בכריכה</div>
          <div class="status-value">${statusCounts.inBinding}</div>
        </div>
        <div class="status-card shipment">
          <div class="status-label">במשלוח</div>
          <div class="status-value">${statusCounts.inShipment}</div>
        </div>
        <div class="status-card done">
          <div class="status-label">הושלמו</div>
          <div class="status-value">${statusCounts.delivered}</div>
        </div>
      </div>
    </div>
    <div class="section">
      <div class="section-title">✅ מה בוצע היום</div>
      <div class="today-list">${renderTodayActivity(todayCounts)}</div>
    </div>
    <div class="section">
      <div class="section-title">⚠️ התראות</div>
      <div class="alerts-grid">
        <div class="alert-card alert-overdue">
          <div class="alert-label">פריטים באיחור</div>
          <div class="alert-value">${overdueCount}</div>
        </div>
        <div class="alert-card alert-upcoming">
          <div class="alert-label">דדליינים ב-5 ימי עבודה</div>
          <div class="alert-value">${upcomingCount}</div>
        </div>
      </div>
    </div>
    <div class="footer">🖨️ נטלי פרינט • סיכום יומי אוטומטי • שעה 17:00</div>
  </div>
</body>
</html>`;
}

function renderTodayActivity(counts: any): string {
  const items = [
    { icon: "🆕", label: "הזמנות חדשות נקלטו", value: counts.newOrders },
    { icon: "✏️", label: "סקיצות אושרו", value: counts.designApproved },
    { icon: "🖨️", label: "עברו להדפסה", value: counts.movedToPrint },
    { icon: "📦", label: "נשלחו לדבח", value: counts.sentToDavach },
    { icon: "🔧", label: "נשלחו לאיציק", value: counts.sentToItzik },
    { icon: "🚚", label: "יצאו למשלוח", value: counts.movedToShip },
    { icon: "✅", label: "נמסרו ללקוח", value: counts.delivered },
  ].filter((i) => i.value > 0);

  if (items.length === 0) {
    return '<div class="today-empty">לא בוצעו פעולות היום</div>';
  }

  return items.map((i) => `
    <div class="today-item">
      <div class="today-item-label">
        <span style="font-size: 18px;">${i.icon}</span>
        <span>${i.label}</span>
      </div>
      <div class="today-item-value">${i.value}</div>
    </div>
  `).join("");
}

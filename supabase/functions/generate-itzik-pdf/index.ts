// @ts-ignore
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";

const PDFSHIFT_API_URL = "https://api.pdfshift.io/v3/convert/pdf";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function escapeHtml(text: string): string {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function toBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunk = 8192;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

function generateItzikHtml(data: Record<string, unknown>): string {
  const orderNumber = escapeHtml(String(data.order_number ?? ""));
  const itemLetter = escapeHtml(String(data.item_letter ?? ""));
  const generationDate = escapeHtml(String(data.generation_date ?? ""));
  const companyName = escapeHtml(String(data.company_name ?? ""));
  const contactName = escapeHtml(String(data.contact_name ?? "—"));
  const calendarTypeLabel = escapeHtml(String(data.calendar_type_label ?? ""));
  const quantity = Number(data.quantity) || 0;
  const cartonsCount = Number(data.cartons_count) || 0;
  const cartonsSize = Number(data.cartons_size) || 0;
  const notes = data.notes ? escapeHtml(String(data.notes)) : "";

  return `<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
  <meta charset="UTF-8">
  <title>הזמנת עבודה לכריכה - איציק - ${orderNumber}${itemLetter}</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Heebo:wght@400;500;700;900&display=swap');
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Heebo', sans-serif;
      color: #1a1a1a;
      background: #ffffff;
      padding: 40px;
      direction: rtl;
    }
    .header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      border-bottom: 3px solid #ec4899;
      padding-bottom: 20px;
      margin-bottom: 40px;
    }
    .logo { font-size: 28px; font-weight: 900; color: #be185d; }
    .doc-title { font-size: 22px; font-weight: 700; color: #1a1a1a; }
    .meta {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 20px;
      margin-bottom: 40px;
      background: #fdf2f8;
      padding: 20px;
      border-radius: 12px;
      border-right: 4px solid #ec4899;
    }
    .meta-item { font-size: 14px; }
    .meta-label { color: #6b7280; margin-bottom: 4px; }
    .meta-value { font-weight: 700; font-size: 16px; color: #1a1a1a; }
    .section { margin-bottom: 30px; }
    .section-title {
      font-size: 18px;
      font-weight: 700;
      color: #be185d;
      border-bottom: 2px solid #e5e7eb;
      padding-bottom: 8px;
      margin-bottom: 15px;
    }
    .details-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
    .detail-row {
      display: flex;
      justify-content: space-between;
      padding: 12px 0;
      border-bottom: 1px solid #f3f4f6;
    }
    .detail-label { color: #6b7280; font-size: 14px; }
    .detail-value { font-weight: 600; font-size: 15px; color: #1a1a1a; }
    .info-box {
      background: linear-gradient(135deg, #be185d 0%, #ec4899 100%);
      color: white;
      padding: 20px;
      border-radius: 12px;
      text-align: center;
      margin: 30px 0;
    }
    .info-title { font-size: 14px; opacity: 0.9; margin-bottom: 8px; }
    .info-value { font-size: 28px; font-weight: 900; }
    .notes-box {
      background: #fef3c7;
      border-right: 4px solid #f59e0b;
      padding: 15px;
      border-radius: 8px;
      margin-top: 20px;
    }
    .notes-title { font-weight: 700; color: #92400e; margin-bottom: 6px; font-size: 14px; }
    .notes-content { color: #78350f; font-size: 14px; line-height: 1.6; white-space: pre-wrap; }
    .footer {
      margin-top: 60px;
      padding-top: 20px;
      border-top: 2px solid #e5e7eb;
      display: flex;
      justify-content: space-between;
      font-size: 12px;
      color: #6b7280;
    }
    .footer-signature { text-align: center; }
    .footer-signature-line {
      border-bottom: 1px solid #6b7280;
      width: 200px;
      margin-bottom: 4px;
      padding-top: 30px;
    }
  </style>
</head>
<body>
  <div class="header">
    <div class="logo">נטלי פרינט 🖨️</div>
    <div class="doc-title">📋 הזמנת עבודה לכריכה - איציק</div>
  </div>
  <div class="meta">
    <div class="meta-item">
      <div class="meta-label">מספר הזמנה</div>
      <div class="meta-value">${orderNumber}${itemLetter ? ` / ${itemLetter}` : ""}</div>
    </div>
    <div class="meta-item">
      <div class="meta-label">תאריך הזמנה</div>
      <div class="meta-value">${generationDate}</div>
    </div>
  </div>
  <div class="section">
    <div class="section-title">👤 פרטי הלקוח</div>
    <div class="details-grid">
      <div class="detail-row">
        <span class="detail-label">שם החברה:</span>
        <span class="detail-value">${companyName}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">איש קשר:</span>
        <span class="detail-value">${contactName || "—"}</span>
      </div>
    </div>
  </div>
  <div class="section">
    <div class="section-title">📦 פרטי הכריכה</div>
    <div class="details-grid">
      <div class="detail-row">
        <span class="detail-label">סוג לוח:</span>
        <span class="detail-value">${calendarTypeLabel}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">כמות יחידות:</span>
        <span class="detail-value">${quantity.toLocaleString("he-IL")}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">כמות בקרטון:</span>
        <span class="detail-value">${cartonsSize.toLocaleString("he-IL")} יחידות</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">מספר קרטונים:</span>
        <span class="detail-value">${cartonsCount.toLocaleString("he-IL")}</span>
      </div>
    </div>
  </div>
  <div class="info-box">
    <div class="info-title">📦 סה״כ לכריכה</div>
    <div class="info-value">${quantity.toLocaleString("he-IL")} יחידות</div>
  </div>
  ${notes ? `
  <div class="notes-box">
    <div class="notes-title">📝 הערות</div>
    <div class="notes-content">${notes}</div>
  </div>` : ""}
  <div class="footer">
    <div class="footer-signature">
      <div class="footer-signature-line"></div>
      <div>חתימת איציק (קבלה)</div>
    </div>
    <div class="footer-signature">
      <div class="footer-signature-line"></div>
      <div>חתימת מנהל</div>
    </div>
  </div>
</body>
</html>`;
}

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: CORS_HEADERS });
  }

  try {
    const requestData = await req.json();

    const required = [
      "order_number",
      "company_name",
      "calendar_type_label",
      "quantity",
      "cartons_count",
      "cartons_size",
    ];

    for (const field of required) {
      if (requestData[field] === undefined || requestData[field] === null) {
        return new Response(
          JSON.stringify({ error: `Missing required field: ${field}` }),
          { status: 400, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
        );
      }
    }

    if (!requestData.generation_date) {
      requestData.generation_date = new Date().toLocaleDateString("he-IL", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
      });
    }

    const html = generateItzikHtml(requestData);

    const pdfshiftKey = Deno.env.get("PDFSHIFT_API_KEY");
    if (!pdfshiftKey) {
      return new Response(
        JSON.stringify({ error: "PDFSHIFT_API_KEY not configured" }),
        { status: 500, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
      );
    }

    const pdfResponse = await fetch(PDFSHIFT_API_URL, {
      method: "POST",
      headers: {
        "X-API-Key": pdfshiftKey,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        source: html,
        landscape: false,
        format: "A4",
        margin: "10mm",
        use_print: true,
      }),
    });

    if (!pdfResponse.ok) {
      const errorText = await pdfResponse.text();
      return new Response(
        JSON.stringify({ error: "PDFShift API error", details: errorText }),
        { status: 500, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
      );
    }

    const pdfBytes = new Uint8Array(await pdfResponse.arrayBuffer());
    const pdfBase64 = toBase64(pdfBytes);
    const itemLetter = requestData.item_letter || "";

    return new Response(
      JSON.stringify({
        success: true,
        pdf_base64: pdfBase64,
        filename: `itzik_order_${requestData.order_number}${itemLetter}.pdf`,
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

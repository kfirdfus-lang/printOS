// @ts-ignore
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const SYSTEM_PROMPT = `אתה עוזר לפרסר מייל הזמנה של פרויקט "אלוט" - לוחות שנה.

המייל מגיע מקארין מחברת אלוט, ואתה צריך לחלץ ממנו את פרטי ההזמנה בפורמט JSON.

## סוגי לוחות:
- "מנשא קשיח" / "כריכה קשה" → calendar_type: "hard"
- "מנשא רך" / "דופלקס" → calendar_type: "duplex"  
- "לוחות קיר" / "לוח קיר" → calendar_type: "wall"

## פורמט התשובה (JSON בלבד, ללא טקסט נוסף):
{
  "order_number": 62,
  "company_name": "שם החברה",
  "contact_name": "שם איש קשר",
  "contact_phone": "מספר טלפון",
  "contact_email": "מייל או null",
  "delivery_address": "כתובת מלאה או null",
  "is_pickup": false,
  "items": [
    {
      "calendar_type": "hard",
      "quantity": 500,
      "needs_design": true
    },
    {
      "calendar_type": "wall",
      "quantity": 50,
      "needs_design": true
    }
  ]
}

## חוקים חשובים:
1. **needs_design**: אם כתוב "ללא לוגו וברכה" → false. אחרת (או אם כתוב "מצ"ב לוגו וברכה") → true.
2. **is_pickup**: אם הכתובת מכילה "איסוף עצמי" או "בית הדפוס" → true, ואז delivery_address = null.
3. **פריטים מרובים**: אם יש "40 X ו-20 Y" - זה 2 פריטים נפרדים במערך items.
4. **כתובות מרובות**: אם יש חלוקה לכמה כתובות (למשל "100 יחידות ל... 80 יחידות ל...") - שים את הכל בשדה delivery_address כטקסט חופשי.
5. **מספר טלפון**: שמור את הפורמט המקורי (עם מקפים אם יש).
6. **מייל חסר**: אם אין מייל - contact_email: null. אל תמציא.
7. **אם משהו לא ברור** - שים null באותו שדה. אל תנחש.

## דוגמאות:

### דוגמה 1 - הזמנה פשוטה עם פריט אחד:
מייל:
"""
הזמנה מספר 51 :
שם החברה: מהוד הנדסה
כמות: 180 לוחות שולחנים במנשא קשיח מצ"ב לוגו וברכה
שם איש קשר: רונית
טלפון: 054-4595590
מייל: ronit-t@mahod.co.il
כתובת: רחוב יהודה הנחתום 4 באר שבע
"""

תשובה:
{
  "order_number": 51,
  "company_name": "מהוד הנדסה",
  "contact_name": "רונית",
  "contact_phone": "054-4595590",
  "contact_email": "ronit-t@mahod.co.il",
  "delivery_address": "רחוב יהודה הנחתום 4 באר שבע",
  "is_pickup": false,
  "items": [
    {"calendar_type": "hard", "quantity": 180, "needs_design": true}
  ]
}

### דוגמה 2 - שני סוגי לוחות + איסוף עצמי:
מייל:
"""
הזמנה מספר 62 :
שם החברה: אלומות אש
כמות: 40 לוחות שולחנים במנשא רך ו 20 לוחות קיר מצ"ב לוגו וברכה
איש קשר: ז'נט
טלפון: 053-5472786
כתובת: איסוף עצמי מבית הדפוס
"""

תשובה:
{
  "order_number": 62,
  "company_name": "אלומות אש",
  "contact_name": "ז'נט",
  "contact_phone": "053-5472786",
  "contact_email": null,
  "delivery_address": null,
  "is_pickup": true,
  "items": [
    {"calendar_type": "duplex", "quantity": 40, "needs_design": true},
    {"calendar_type": "wall", "quantity": 20, "needs_design": true}
  ]
}

### דוגמה 3 - ללא לוגו וברכה:
מייל:
"""
הזמנה מספר 36 :
שם החברה: ועד אגף החשב- בנק הפועלים
כמות: 95 לוחות שולחנים במנשא רך ללא לוגו וברכה
שם איש קשר: גאנה
טלפון: 054-7766834
כתובת: הנגב 11 ת"א
"""

תשובה:
{
  "order_number": 36,
  "company_name": "ועד אגף החשב- בנק הפועלים",
  "contact_name": "גאנה",
  "contact_phone": "054-7766834",
  "contact_email": null,
  "delivery_address": "הנגב 11 ת\"א",
  "is_pickup": false,
  "items": [
    {"calendar_type": "duplex", "quantity": 95, "needs_design": false}
  ]
}

## חשוב מאוד:
- החזר JSON תקין בלבד, ללא markdown, ללא הסבר, ללא טקסט לפני או אחרי.
- אם אתה לא בטוח משהו - שים null באותו שדה, אל תמציא.`;

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: CORS_HEADERS });
  }

  try {
    const { email_text } = await req.json();

    if (!email_text || typeof email_text !== "string") {
      return new Response(
        JSON.stringify({ error: "Missing email_text" }),
        { status: 400, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
      );
    }

    const apiKey = Deno.env.get("ANTHROPIC_API_KEY");
    if (!apiKey) {
      return new Response(
        JSON.stringify({ error: "ANTHROPIC_API_KEY not configured" }),
        { status: 500, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
      );
    }

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 1024,
        system: SYSTEM_PROMPT,
        messages: [
          {
            role: "user",
            content: `פרסר את המייל הבא:\n\n${email_text}`,
          },
        ],
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return new Response(
        JSON.stringify({ error: "Claude API error", details: errorText }),
        { status: 500, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
      );
    }

    const claudeResponse = await response.json();
    const rawText = claudeResponse.content[0].text.trim();

    let parsedData;
    try {
      const cleanText = rawText
        .replace(/^```json\s*/i, "")
        .replace(/^```\s*/i, "")
        .replace(/\s*```$/i, "")
        .trim();

      parsedData = JSON.parse(cleanText);
    } catch {
      return new Response(
        JSON.stringify({
          error: "Failed to parse Claude response as JSON",
          raw_response: rawText,
        }),
        { status: 500, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
      );
    }

    return new Response(
      JSON.stringify({ success: true, data: parsedData }),
      { headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return new Response(
      JSON.stringify({ error: message }),
      { status: 500, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } },
    );
  }
});

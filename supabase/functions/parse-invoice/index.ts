// Package E — parse supplier invoices via Claude (no Bina API).
// Deploy: supabase functions deploy parse-invoice --no-verify-jwt

// @ts-ignore
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { rejectDisallowedInternalOrigin } from "../_shared/cors.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const MODEL = "claude-sonnet-5";
const VAT_RATE = 0.18;
const OWN_TAX_ID = "517205332"; // ח.פ של נטלי פתרונות הדפסה

const SYSTEM_PROMPT = `אתה מפרסר חשבוניות ספקים ישראליות עבור מערכת הנהלת חשבונות.

תפקידך: לחלץ מהמסמך את השדות המבוקשים ולהחזיר JSON בלבד.

כללים מחייבים:
1. החזר JSON תקין בלבד. בלי טקסט לפני, בלי טקסט אחרי, בלי סימני markdown, בלי \`\`\`.
2. אם שדה לא מופיע במסמך או שאתה לא מצליח לקרוא אותו — החזר null. אל תנחש. אל תמציא.
3. תאריכים בפורמט YYYY-MM-DD בלבד. שים לב: בישראל הפורמט על המסמך הוא DD/MM/YYYY.
   כלומר 22/07/2026 הוא 22 ביולי 2026 ולא 7 בפברואר.
4. מספרים כמספרים — בלי ₪, בלי פסיקים, בלי רווחים. 12,345.67 → 12345.67
5. שים ב-low_confidence_fields את שמות כל השדות שלא היית בטוח בהם
   (טשטוש, חיתוך, כתב יד, חותמת מעל הטקסט, ספרה מעורפלת).
6. אם המסמך אינו חשבונית ספק — החזר null בכל השדות ותכתוב הסבר ב-notes.
7. אם ערך טקסט מכיל תו גרשיים (") — הברח אותו כ-\\" כנדרש
   בתקן JSON. שים לב במיוחד לשמות חברות ישראליות
   שנגמרות ב-בע"מ. ודא שהפלט עובר JSON.parse תקין.

הבהרות על שדות ספציפיים:
- allocation_number = "מספר הקצאה" / "מס' הקצאה" — מספר מרשות המסים, בדרך כלל 9 ספרות.
  שדה חדש יחסית, לא תמיד קיים. אם אינו מופיע — null.
- supplier_tax_id = ח.פ / ע.מ / עוסק מורשה של **הספק המנפיק**, ספרות בלבד, בדרך כלל 9 ספרות.
  קריטי: בחשבונית ישראלית מופיעים לרוב שני מספרי ח.פ —
  של הספק (בדרך כלל בראש המסמך, ליד הלוגו/שם המנפיק)
  ושל הלקוח (ליד "לכבוד" / "לקוח").
  חלץ אך ורק את של הספק המנפיק.
  אם אתה לא בטוח איזה מהם שייך לספק — החזר null
  והוסף את השדה ל-low_confidence_fields.
  אל תנחש.
- invoice_number = מספר החשבונית של הספק.
- invoice_date = תאריך החשבונית. אם יש גם "ת. הדפסה" — היא אינה תאריך החשבונית.
- vat_date = "תאריך מע\\"מ" אם מופיע בנפרד. אם לא — החזר את אותו ערך כמו invoice_date.
- vat_rate = שיעור המע"מ באחוזים כמספר. 18% → 18
- payment_terms = תנאי תשלום כפי שמופיעים ("שוטף+30", "מזומן", "שוטף+60" וכו').
- due_date = תאריך התשלום אם מופיע במסמך במפורש.
  מילות מפתח: "לתשלום עד", "תאריך פרעון", "יש לשלם עד",
  "מועד תשלום", "לתשלום ב-".
  פורמט YYYY-MM-DD. אם לא מופיע במסמך — null.
  אל תחשב אותו מתנאי תשלום, רק חלץ אם כתוב.
- discount_percent / discount_amount = הנחה כללית על החשבונית (לא ברמת שורה).

מבנה ה-JSON המדויק:
{
  "supplier_name": "שם הספק המנפיק" או null,
  "supplier_tax_id": "ספרות בלבד" או null,
  "supplier_address": "רחוב ומספר" או null,
  "supplier_city": "עיר" או null,
  "supplier_phone": "טלפון" או null,
  "allocation_number": "מספר הקצאה" או null,
  "invoice_number": "מספר החשבונית" או null,
  "invoice_date": "YYYY-MM-DD" או null,
  "vat_date": "YYYY-MM-DD" או null,
  "payment_terms": "תנאי תשלום" או null,
  "due_date": "YYYY-MM-DD" או null,
  "currency": "ILS",
  "amount_before_vat": מספר או null,
  "discount_percent": מספר או null,
  "discount_amount": מספר או null,
  "amount_after_discount": מספר או null,
  "vat_rate": מספר או null,
  "vat_amount": מספר או null,
  "total_amount": מספר או null,
  "line_items": [
    {
      "description": "תיאור השורה",
      "quantity": מספר או null,
      "unit_price": מספר או null,
      "discount_percent": מספר או null,
      "total": מספר או null
    }
  ],
  "low_confidence_fields": ["שם_שדה"],
  "notes": "הערה חופשית אם יש משהו חריג" או null
}`;

/** Escape bare " inside JSON string values (e.g. בע"מ). */
function repairUnescapedQuotes(jsonStr: string): string {
  let out = "";
  let inString = false;
  for (let i = 0; i < jsonStr.length; i++) {
    const c = jsonStr[i];
    if (!inString) {
      out += c;
      if (c === '"') inString = true;
      continue;
    }
    if (c === "\\") {
      out += c + (jsonStr[i + 1] ?? "");
      i++;
      continue;
    }
    if (c === '"') {
      const rest = jsonStr.slice(i + 1);
      if (/^\s*[,:}\]]/.test(rest)) {
        inString = false;
        out += c;
      } else {
        out += '\\"';
      }
      continue;
    }
    out += c;
  }
  return out;
}

function tryParseAiJson(raw: string): { ok: true; value: Record<string, unknown> } | { ok: false; raw: string } {
  const cleaned = String(raw || "")
    .replace(/```json/gi, "")
    .replace(/```/g, "")
    .trim();

  const candidates: string[] = [cleaned];
  const start = cleaned.indexOf("{");
  const end = cleaned.lastIndexOf("}");
  if (start >= 0 && end > start) {
    const sliced = cleaned.slice(start, end + 1);
    if (sliced !== cleaned) candidates.push(sliced);
  }

  for (const candidate of candidates) {
    try {
      return { ok: true, value: JSON.parse(candidate) as Record<string, unknown> };
    } catch {
      /* continue */
    }
    try {
      return {
        ok: true,
        value: JSON.parse(repairUnescapedQuotes(candidate)) as Record<string, unknown>,
      };
    } catch {
      /* continue */
    }
  }
  return { ok: false, raw: cleaned };
}

serve(async (req) => {
  const originBlock = rejectDisallowedInternalOrigin(req);
  if (originBlock) return originBlock;

  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  const started = Date.now();
  let rawAiText: string | null = null;

  try {
    const apiKey = Deno.env.get("ANTHROPIC_API_KEY");
    if (!apiKey) throw new Error("ANTHROPIC_API_KEY לא מוגדר ב-secrets");

    const { file_base64, media_type, file_name } = await req.json();
    if (!file_base64 || !media_type) {
      throw new Error("חסרים פרמטרים: file_base64 / media_type");
    }

    const fileBlock =
      media_type === "application/pdf"
        ? {
          type: "document",
          source: {
            type: "base64",
            media_type: "application/pdf",
            data: file_base64,
          },
        }
        : {
          type: "image",
          source: { type: "base64", media_type, data: file_base64 },
        };

    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: 8000,
        system: SYSTEM_PROMPT,
        messages: [
          {
            role: "user",
            content: [
              fileBlock,
              {
                type: "text",
                text: "פרסר את חשבונית הספק הזו והחזר JSON בלבד לפי המבנה שהוגדר.",
              },
            ],
          },
        ],
      }),
    });

    if (!res.ok) {
      const errText = await res.text();
      throw new Error(`Anthropic API ${res.status}: ${errText}`);
    }

    const data = await res.json();
    const stopReason = data.stop_reason || null;

    const text = (data.content || [])
      .filter((b: { type: string }) => b.type === "text")
      .map((b: { text: string }) => b.text)
      .join("\n");
    rawAiText = text;

    console.error("parse-invoice AI response meta", {
      stop_reason: stopReason,
      usage: data.usage || null,
      text_length: text.length,
    });
    console.error("parse-invoice AI raw text FULL:\n" + text);

    const parseResult = tryParseAiJson(text);
    if (!parseResult.ok) {
      const errMsg = stopReason === "max_tokens"
        ? "תשובת ה-AI אינה JSON תקין (התשובה נחתכה — stop_reason=max_tokens)"
        : "תשובת ה-AI אינה JSON תקין";
      console.error("parse-invoice JSON parse failed", {
        stop_reason: stopReason,
        raw_length: parseResult.raw.length,
      });
      return new Response(
        JSON.stringify({
          success: false,
          error: errMsg,
          raw_text: parseResult.raw,
          stop_reason: stopReason,
          model_used: MODEL,
          parse_duration_ms: Date.now() - started,
          usage: data.usage || null,
          file_name: file_name || null,
        }),
        {
          status: 200,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        },
      );
    }
    const parsed = parseResult.value;

    const lowConf: string[] = Array.isArray(parsed.low_confidence_fields)
      ? [...(parsed.low_confidence_fields as string[])]
      : [];
    const warn = (msg: string) => {
      parsed.notes = (parsed.notes ? String(parsed.notes) + " | " : "") + msg;
    };

    const before = Number(parsed.amount_before_vat);
    const vat = Number(parsed.vat_amount);
    const total = Number(parsed.total_amount);
    const disc = Number(parsed.discount_amount) || 0;
    const base = !isNaN(before) ? before - disc : NaN;

    if (!isNaN(base) && !isNaN(vat) && !isNaN(total)) {
      if (Math.abs(base + vat - total) > 0.05) {
        if (!lowConf.includes("total_amount")) lowConf.push("total_amount");
        warn(`⚠️ הסכומים לא מסתדרים: ${base} + ${vat} ≠ ${total}`);
      }
    }

    if (!isNaN(base) && !isNaN(vat) && base > 0) {
      const rate = vat / base;
      if (Math.abs(rate - VAT_RATE) > 0.01) {
        warn(`⚠️ שיעור מע"מ חריג: ${(rate * 100).toFixed(1)}%`);
      }
    }

    if (parsed.supplier_tax_id) {
      const digits = String(parsed.supplier_tax_id).replace(/\D/g, "");
      if (digits === OWN_TAX_ID) {
        parsed.supplier_tax_id = null;
        if (!lowConf.includes("supplier_tax_id")) lowConf.push("supplier_tax_id");
        warn("⚠️ חולץ הח.פ של נטלי במקום של הספק - נוקה");
      } else {
        if (digits.length !== 9) {
          if (!lowConf.includes("supplier_tax_id")) lowConf.push("supplier_tax_id");
          warn(`⚠️ ח.פ באורך חריג (${digits.length} ספרות)`);
        }
        parsed.supplier_tax_id = digits;
      }
    }

    if (parsed.invoice_date) {
      const d = new Date(String(parsed.invoice_date));
      const in30 = new Date(Date.now() + 30 * 864e5);
      if (d > in30) {
        if (!lowConf.includes("invoice_date")) lowConf.push("invoice_date");
        warn("⚠️ תאריך החשבונית עתידי - ייתכן בלבול DD/MM");
      }
    }

    if (!parsed.vat_date && parsed.invoice_date) {
      parsed.vat_date = parsed.invoice_date;
    }

    parsed.low_confidence_fields = lowConf;

    return new Response(
      JSON.stringify({
        success: true,
        parsed,
        model_used: MODEL,
        stop_reason: stopReason,
        parse_duration_ms: Date.now() - started,
        usage: data.usage || null,
        file_name: file_name || null,
      }),
      { headers: { ...corsHeaders, "Content-Type": "application/json" } },
    );
  } catch (err) {
    console.error("parse-invoice error:", err);
    console.error("parse-invoice rawAiText on error FULL:\n" + (rawAiText || ""));
    return new Response(
      JSON.stringify({
        success: false,
        error: String(err instanceof Error ? err.message : err),
        raw_text: rawAiText,
        parse_duration_ms: Date.now() - started,
      }),
      {
        status: 200,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      },
    );
  }
});

// supabase/functions/send-collection-email/index.ts
// שליחת מייל גבייה ללקוח עם פירוט חשבוניות פתוחות
// 2 סוגים: תזכורת רגילה / בקשת בירור (לאיחור)

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

const REPLY_TO_EMAIL = 'kfir.dfus@gmail.com';
const FROM_NAME = 'נטלי פתרונות הדפסה';

interface Invoice {
  doc_num: string | number;
  doc_balance: number;
  doc_date: string;
  doc_payment_date: string;
  days_overdue?: number;
}

interface RequestBody {
  bina_customer_id: string;
  email_type: 'reminder' | 'inquiry'; // תזכורת רגילה / בקשת בירור
  invoices: Invoice[];
  /** אם מועבר — שולח רק לכתובות אלה (מאגר מיילי גבייה) */
  to_emails?: string[];
}

function formatDate(dateStr: string): string {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}

/** מטבע כטקסט בלבד (לוג DB וכו') — ללא HTML */
function formatCurrencyPlain(n: number): string {
  return `₪${Number(n || 0).toLocaleString('he-IL', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function formatCurrency(n: number): string {
  const num = Number(n || 0).toLocaleString('he-IL', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return `<bdi style="unicode-bidi:isolate;direction:ltr">₪${num}</bdi>`;
}

function buildEmailHTML(
  clientName: string,
  contactName: string | null,
  invoices: Invoice[],
  isInquiry: boolean,
): string {
  const total = invoices.reduce((s, inv) => s + Number(inv.doc_balance || 0), 0);

  const greeting = contactName ? `שלום ${contactName},` : 'שלום רב,';

  const introText = isInquiry
    ? `החשבוניות הבאות נמצאות באיחור בתשלום. אנא בדקו את סטטוס התשלום וחזרו אלינו בהקדם להסדרת העניין.`
    : `לקראת מועד הפירעון, מצורפת רשימת חשבוניות לתשלום. נשמח לקבל את התשלום במועד.`;

  const tableRows = invoices.map((inv) => {
    const overdueStr = inv.days_overdue && inv.days_overdue > 0
      ? `<span style="color:#dc2626;font-weight:600">איחור ${inv.days_overdue} ימים</span>`
      : `<span style="color:#16a34a">בתוקף</span>`;

    return `
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid #e5e7eb;text-align:center;font-weight:600;color:#1E3A52">#${inv.doc_num}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #e5e7eb;text-align:center">${formatDate(inv.doc_date)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #e5e7eb;text-align:center">${formatDate(inv.doc_payment_date)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #e5e7eb;text-align:left;font-weight:700;color:#1E3A52;direction:ltr;unicode-bidi:plaintext">${formatCurrency(inv.doc_balance)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #e5e7eb;text-align:center;font-size:12px">${overdueStr}</td>
      </tr>
    `;
  }).join('');

  const headerColor = isInquiry ? '#dc2626' : '#3DB5B1';
  const titleText = isInquiry ? '🔔 בקשה לבדיקת סטטוס תשלום' : '📧 תזכורת לתשלום חשבוניות';

  return `<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<title>${titleText}</title>
</head>
<body style="margin:0;padding:0;background:#f5f5f5;font-family:'Segoe UI',Arial,sans-serif;direction:rtl">
  <div style="max-width:680px;margin:30px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.05)">

    <!-- HEADER -->
    <div style="background:linear-gradient(135deg, ${headerColor} 0%, #1E3A52 100%);padding:30px;color:#fff;text-align:center">
      <div style="font-size:32px;font-weight:900;letter-spacing:2px;margin-bottom:6px">NATALIE</div>
      <div style="font-size:14px;opacity:0.9">פתרונות הדפסה</div>
    </div>

    <!-- BODY -->
    <div style="padding:32px 30px">
      <h2 style="color:#1E3A52;margin:0 0 8px;font-size:20px">${titleText}</h2>
      <p style="color:#6b7280;font-size:14px;margin:0 0 24px">לכבוד: <strong style="color:#1E3A52">${clientName}</strong></p>

      <p style="color:#1E3A52;font-size:15px;line-height:1.7;margin:0 0 20px">${greeting}</p>
      <p style="color:#4b5563;font-size:14px;line-height:1.7;margin:0 0 24px">${introText}</p>

      <!-- TABLE -->
      <div style="overflow-x:auto;margin:24px 0">
        <table style="width:100%;border-collapse:collapse;background:#f8fafa;border-radius:8px;overflow:hidden">
          <thead>
            <tr style="background:#1E3A52;color:#fff">
              <th style="padding:12px;text-align:center;font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px">מס' חשבונית</th>
              <th style="padding:12px;text-align:center;font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px">תאריך הוצאה</th>
              <th style="padding:12px;text-align:center;font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px">תאריך פירעון</th>
              <th style="padding:12px;text-align:left;font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px">סכום</th>
              <th style="padding:12px;text-align:center;font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px">סטטוס</th>
            </tr>
          </thead>
          <tbody>${tableRows}</tbody>
          <tfoot>
            <tr style="background:#3DB5B1;color:#fff;font-weight:700">
              <td colspan="3" style="padding:14px;text-align:right">סה"כ לתשלום:</td>
              <td style="padding:14px;text-align:left;font-size:18px;direction:ltr;unicode-bidi:plaintext">${formatCurrency(total)}</td>
              <td></td>
            </tr>
          </tfoot>
        </table>
      </div>

      <!-- CALL TO ACTION -->
      <div style="background:#f0fdfa;padding:16px 20px;border-radius:8px;border-right:4px solid #3DB5B1;margin:24px 0;color:#1E3A52;font-size:14px;line-height:1.7">
        ${isInquiry
    ? `<strong>נשמח לדעת:</strong> מה סטטוס התשלום? האם יש בעיה? אנא חזרו אלינו ב-${REPLY_TO_EMAIL} או 03-6815703.`
    : `<strong>פרטי תשלום:</strong> ניתן להעביר בהעברה בנקאית או בצ'ק. לפרטים נוספים נא ליצור קשר.`
}
      </div>

      <p style="color:#1E3A52;font-size:14px;line-height:1.7;margin:20px 0 0">בברכה,<br/><strong>${FROM_NAME} בע"מ</strong></p>
    </div>

    <!-- FOOTER -->
    <div style="background:#f8fafa;padding:20px 30px;border-top:2px solid #3DB5B1;text-align:center;color:#6b7280;font-size:12px;line-height:1.8">
      <div style="color:#1E3A52;font-weight:700;font-size:13px;margin-bottom:4px">נטלי פתרונות הדפסה בע"מ</div>
      ח.פ. 517205332 | שד' הר ציון 104, תל אביב | 03-6815703
    </div>
  </div>
</body>
</html>`;
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const resendKey = Deno.env.get('RESEND_API_KEY');
    if (!resendKey) {
      return new Response(JSON.stringify({ error: 'RESEND_API_KEY לא מוגדר' }),
        { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    const supabaseUrl = Deno.env.get('SUPABASE_URL')!;
    const supabaseKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
    const supabase = createClient(supabaseUrl, supabaseKey);

    const body: RequestBody = await req.json();
    const { bina_customer_id, email_type, invoices, to_emails } = body;

    if (!bina_customer_id || !invoices || invoices.length === 0) {
      return new Response(JSON.stringify({ error: 'חסרים נתונים' }),
        { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    // שליפת פרטי הלקוח
    const { data: client, error: clientErr } = await supabase
      .from('clients')
      .select('id, name, contact_name, collection_email_primary, collection_email_secondary')
      .eq('bina_customer_id', String(bina_customer_id))
      .maybeSingle();

    if (clientErr || !client) {
      return new Response(JSON.stringify({ error: 'לקוח לא נמצא' }),
        { status: 404, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    const recipients: string[] = Array.isArray(to_emails)
      ? to_emails.map((e) => String(e).trim().toLowerCase()).filter((e) => e.includes('@'))
      : [];
    if (recipients.length === 0) {
      if (client.collection_email_primary) recipients.push(client.collection_email_primary);
      if (client.collection_email_secondary) recipients.push(client.collection_email_secondary);
    }

    if (recipients.length === 0) {
      return new Response(JSON.stringify({
        error: 'אין כתובת מייל לגבייה ללקוח זה',
        need_email: true,
      }), { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    const isInquiry = email_type === 'inquiry';
    const subject = isInquiry
      ? `בקשה לבדיקת סטטוס תשלום - ${client.name}`
      : `תזכורת לתשלום חשבוניות - ${client.name}`;

    const html = buildEmailHTML(client.name, client.contact_name, invoices, isInquiry);

    // שליחה דרך Resend
    const resendResponse = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${resendKey}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        from: `${FROM_NAME} <gvia@natalie-print.com>`,
        to: recipients,
        replyTo: REPLY_TO_EMAIL,
        subject,
        html,
      }),
    });

    const resendResult = await resendResponse.json();

    if (!resendResponse.ok) {
      console.error('[send-collection-email] Resend rejected:', resendResponse.status, JSON.stringify(resendResult));
      return new Response(JSON.stringify({
        success: false,
        error: 'שגיאה בשליחת מייל',
        details: resendResult,
      }), { status: 502, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    // שמירת פעולה ב-debt_actions
    const totalAmount = invoices.reduce((s, inv) => s + Number(inv.doc_balance || 0), 0);
    await supabase.from('debt_actions').insert({
      bina_customer_id: String(bina_customer_id),
      client_id: client.id,
      client_name: client.name,
      action_type: 'email_sent',
      notes: `נשלח ${isInquiry ? 'בקשת בירור' : 'תזכורת'} עבור ${invoices.length} חשבוניות (${formatCurrencyPlain(totalAmount)}) ל: ${recipients.join(', ')}`,
      amount: totalAmount,
      created_by: 'admin',
    });

    return new Response(JSON.stringify({
      success: true,
      sent_to: recipients,
      email_id: resendResult.id,
      invoices_count: invoices.length,
      total_amount: totalAmount,
    }), { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error('[send-collection-email] חריגה:', msg);
    return new Response(JSON.stringify({ success: false, error: msg }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
  }
});

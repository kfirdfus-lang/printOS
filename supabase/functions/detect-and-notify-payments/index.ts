// supabase/functions/detect-and-notify-payments/index.ts
//
// פונקציה זאת:
// 1. משווה בין snapshot ה-debt האחרון לקודם
// 2. מזהה חשבוניות שנעלמו (שולמו)
// 3. שולחת מייל למורן על כל אחת
// 4. שומרת ב-detected_payments למניעת כפילויות
//
// קורא מ-bina-fetch-debt-report בסוף הריצה שלו

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

const ACCOUNTANT_EMAIL = 'moran.natalieprint@gmail.com';
const FROM_NAME = 'נטלי פתרונות הדפסה';
const FROM_EMAIL = 'gvia@natalie-print.com';
const REPLY_TO = 'kfir.dfus@gmail.com';

// סף מינימלי לסכום - מתחת לזה לא שולחים (להגנה מ"רעש")
const MIN_AMOUNT = 10;

function formatDate(d: string | null): string {
  if (!d) return '';
  const date = new Date(d);
  return `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
}

function formatCurrencyPlain(n: number): string {
  return `₪${Number(n || 0).toLocaleString('he-IL', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function formatCurrencyHTML(n: number): string {
  const num = Number(n || 0).toLocaleString('he-IL', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return `<bdi style="unicode-bidi:isolate;direction:ltr">₪${num}</bdi>`;
}

function buildEmailHTML(payment: any): string {
  return `<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<title>קבלת תשלום</title>
</head>
<body style="margin:0;padding:0;background:#f5f5f5;font-family:'Segoe UI',Arial,sans-serif;direction:rtl">
  <div style="max-width:600px;margin:30px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.05)">
    
    <!-- HEADER -->
    <div style="background:linear-gradient(135deg, #10b981 0%, #047857 100%);padding:24px 30px;color:#fff;text-align:center">
      <div style="font-size:28px;font-weight:900;letter-spacing:2px;margin-bottom:4px">NATALIE</div>
      <div style="font-size:12px;opacity:0.9">פתרונות הדפסה</div>
      <div style="margin-top:12px;font-size:14px;background:rgba(255,255,255,0.2);padding:6px 14px;border-radius:20px;display:inline-block">
        💰 קבלת תשלום זוהתה
      </div>
    </div>

    <!-- BODY -->
    <div style="padding:30px">
      <h2 style="color:#1E3A52;margin:0 0 8px;font-size:18px">שלום מורן,</h2>
      <p style="color:#4b5563;font-size:14px;line-height:1.7;margin:0 0 20px">
        זוהתה קבלת תשלום בבינה. הפרטים:
      </p>
      
      <table style="width:100%;border-collapse:collapse;background:#f8fafa;border-radius:8px;overflow:hidden;margin:16px 0">
        <tr>
          <td style="padding:14px 16px;border-bottom:1px solid #e5e7eb;color:#6b7280;font-size:12px;font-weight:600;width:140px">🏢 לקוח:</td>
          <td style="padding:14px 16px;border-bottom:1px solid #e5e7eb;color:#1E3A52;font-weight:700;font-size:14px">${payment.customer_name}</td>
        </tr>
        <tr>
          <td style="padding:14px 16px;border-bottom:1px solid #e5e7eb;color:#6b7280;font-size:12px;font-weight:600">📄 חשבונית:</td>
          <td style="padding:14px 16px;border-bottom:1px solid #e5e7eb;color:#1E3A52;font-weight:700;font-size:14px">#${payment.doc_num}</td>
        </tr>
        <tr>
          <td style="padding:14px 16px;border-bottom:1px solid #e5e7eb;color:#6b7280;font-size:12px;font-weight:600">📅 תאריך הוצאה:</td>
          <td style="padding:14px 16px;border-bottom:1px solid #e5e7eb;color:#1E3A52;font-size:14px">${formatDate(payment.doc_date)}</td>
        </tr>
        <tr>
          <td style="padding:14px 16px;border-bottom:1px solid #e5e7eb;color:#6b7280;font-size:12px;font-weight:600">📅 תאריך פירעון:</td>
          <td style="padding:14px 16px;border-bottom:1px solid #e5e7eb;color:#1E3A52;font-size:14px">${formatDate(payment.doc_payment_date)}</td>
        </tr>
        <tr style="background:#dcfce7">
          <td style="padding:16px;color:#065f46;font-size:13px;font-weight:700">💰 סכום ששולם:</td>
          <td style="padding:16px;direction:ltr;unicode-bidi:plaintext;text-align:left;color:#047857;font-weight:800;font-size:20px">${formatCurrencyHTML(payment.amount)}</td>
        </tr>
      </table>
      
      <div style="background:#fffbeb;padding:14px 18px;border-radius:8px;border-right:4px solid #f59e0b;margin:20px 0;color:#78350f;font-size:13px;line-height:1.7">
        <strong>📌 לידיעתך:</strong> נא לעדכן בהתאם במערכת. הפרטים זוהו על ידי השוואת דוחות חוב פתוחים מבינה.
      </div>
      
      <p style="color:#1E3A52;font-size:13px;line-height:1.7;margin:20px 0 0">
        בברכה,<br/>
        <strong>${FROM_NAME}</strong>
      </p>
    </div>

    <!-- FOOTER -->
    <div style="background:#f8fafa;padding:16px 30px;border-top:2px solid #10b981;text-align:center;color:#6b7280;font-size:11px;line-height:1.7">
      <div style="color:#1E3A52;font-weight:700;font-size:12px;margin-bottom:4px">נטלי פתרונות הדפסה בע"מ</div>
      ח.פ. 517205332 | 03-6815703 | מייל אוטומטי - אל תשיבי
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
      return new Response(JSON.stringify({ error: 'RESEND_API_KEY חסר' }),
        { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    const supabaseUrl = Deno.env.get('SUPABASE_URL')!;
    const supabaseKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
    const supabase = createClient(supabaseUrl, supabaseKey);

    // 1. מציאת 2 ה-snapshots האחרונים
    const { data: dates, error: datesErr } = await supabase
      .from('debt_snapshots')
      .select('snapshot_date')
      .order('snapshot_date', { ascending: false })
      .limit(2);

    if (datesErr) throw datesErr;

    const uniqueDates = [...new Set((dates || []).map((d) => d.snapshot_date))];

    if (uniqueDates.length < 2) {
      return new Response(JSON.stringify({
        success: true,
        message: 'עדיין אין מספיק נתונים לזיהוי תשלומים (נדרשים לפחות 2 snapshots)',
        snapshots_count: uniqueDates.length,
      }), { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    const latestDate = uniqueDates[0];
    const previousDate = uniqueDates[1];

    console.log(`[detect-payments] השוואה: ${latestDate} ↔ ${previousDate}`);

    // 2. שליפת חשבוניות מ-2 ה-snapshots
    const { data: latestInvoices, error: latestErr } = await supabase
      .from('debt_snapshots')
      .select('bina_customer_id, customer_name, doc_num, doc_balance, doc_date, doc_payment_date')
      .eq('snapshot_date', latestDate);

    const { data: previousInvoices, error: prevErr } = await supabase
      .from('debt_snapshots')
      .select('bina_customer_id, customer_name, doc_num, doc_balance, doc_date, doc_payment_date')
      .eq('snapshot_date', previousDate);

    if (latestErr || prevErr) throw latestErr || prevErr;

    // 3. בניית מפה מהיר של ה-snapshot האחרון
    const latestMap = new Map();
    for (const inv of (latestInvoices || [])) {
      const key = `${inv.bina_customer_id}_${inv.doc_num}`;
      latestMap.set(key, inv);
    }

    // 4. זיהוי חשבוניות שנעלמו מה-snapshot האחרון = שולמו
    const paidInvoices: any[] = [];
    for (const prevInv of (previousInvoices || [])) {
      const key = `${prevInv.bina_customer_id}_${prevInv.doc_num}`;
      const stillExists = latestMap.has(key);

      if (!stillExists) {
        // חשבונית נעלמה - תשלום!
        const amount = Number(prevInv.doc_balance || 0);
        if (amount >= MIN_AMOUNT) {
          paidInvoices.push(prevInv);
        }
      }
    }

    console.log(`[detect-payments] זוהו ${paidInvoices.length} תשלומים`);

    // 5. סינון תשלומים שכבר שלחנו עליהם מייל
    const newPayments: any[] = [];
    for (const inv of paidInvoices) {
      const { data: existing } = await supabase
        .from('detected_payments')
        .select('id')
        .eq('bina_customer_id', String(inv.bina_customer_id))
        .eq('doc_num', String(inv.doc_num))
        .maybeSingle();

      if (!existing) {
        newPayments.push(inv);
      }
    }

    console.log(`[detect-payments] חדשים (לא נשלחו עדיין): ${newPayments.length}`);

    // 6. שליחת מייל + שמירה ב-detected_payments
    const results: any[] = [];
    for (const inv of newPayments) {
      // קודם שומרים שזיהינו - למניעת כפילות אם המייל ישלח כמה פעמים
      const { data: saved, error: saveErr } = await supabase
        .from('detected_payments')
        .insert({
          bina_customer_id: String(inv.bina_customer_id),
          customer_name: inv.customer_name || `לקוח #${inv.bina_customer_id}`,
          doc_num: String(inv.doc_num),
          amount: Number(inv.doc_balance),
          doc_date: inv.doc_date,
          doc_payment_date: inv.doc_payment_date,
          email_status: 'pending',
        })
        .select()
        .single();

      if (saveErr) {
        // אם כשל כי כבר קיים (race condition) - מתעלמים ומתקדמים
        if (saveErr.code === '23505') continue;
        console.error('[detect-payments] שגיאת שמירה:', saveErr);
        results.push({ doc_num: inv.doc_num, status: 'save_failed', error: saveErr.message });
        continue;
      }

      // עכשיו שולחים מייל
      const subject = `💰 תשלום נקלט: ${inv.customer_name} - ${formatCurrencyPlain(Number(inv.doc_balance))}`;
      const html = buildEmailHTML({
        customer_name: inv.customer_name || `לקוח #${inv.bina_customer_id}`,
        doc_num: inv.doc_num,
        amount: Number(inv.doc_balance),
        doc_date: inv.doc_date,
        doc_payment_date: inv.doc_payment_date,
      });

      const resendResponse = await fetch('https://api.resend.com/emails', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${resendKey}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          from: `${FROM_NAME} <${FROM_EMAIL}>`,
          to: [ACCOUNTANT_EMAIL],
          replyTo: REPLY_TO,
          subject,
          html,
        }),
      });

      const resendResult = await resendResponse.json();

      if (!resendResponse.ok) {
        console.error('[detect-payments] Resend rejected:', resendResponse.status, JSON.stringify(resendResult));
        await supabase.from('detected_payments').update({
          email_status: 'failed',
          email_error: JSON.stringify(resendResult),
        }).eq('id', saved.id);
        results.push({ doc_num: inv.doc_num, status: 'email_failed' });
      } else {
        await supabase.from('detected_payments').update({
          email_status: 'sent',
          email_sent_at: new Date().toISOString(),
        }).eq('id', saved.id);
        results.push({ doc_num: inv.doc_num, status: 'sent', amount: Number(inv.doc_balance) });
      }
    }

    return new Response(JSON.stringify({
      success: true,
      compared_dates: { latest: latestDate, previous: previousDate },
      summary: {
        total_paid_detected: paidInvoices.length,
        new_payments: newPayments.length,
        emails_sent: results.filter((r) => r.status === 'sent').length,
        emails_failed: results.filter((r) => r.status === 'email_failed').length,
      },
      results,
    }), { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error('[detect-payments] חריגה:', msg);
    return new Response(JSON.stringify({ success: false, error: msg }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
  }
});

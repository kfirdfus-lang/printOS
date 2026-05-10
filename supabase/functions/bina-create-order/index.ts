// supabase/functions/bina-create-order/index.ts
// יצירת הזמנה בבינה (docType 15)
// משתמש באותו endpoint כמו create-customer שעובד

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';
import { fetchBinaViaQuotaGuard } from '../_shared/bina-proxy-fetch.ts';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

// אותו endpoint כמו bina-create-customer שעובד
const BINA_URL = 'https://webfiles.binaw.com/post/PostJsonDocV2.aspx';

interface OrderClient {
  binaCustomerId: string | number;
  name: string;
  city: string;
  address: string;
  contactPerson: string;
  phone?: string;
  email?: string;
  customerOrderId?: string;
}

interface OrderItem {
  itemId: string;
  description: string;
  quantity: number;
  unitPrice: number;
  discount?: number;
}

interface OrderPayment {
  payDate: string;
  numberOfPayments: number;
  firstPayment: number;
  creditName: string;
  creditDateEnd?: string;
  creditNum?: string;
  receiptNumber?: string;
}

interface CreateOrderRequest {
  client: OrderClient;
  items: OrderItem[];
  title?: string;
  remark?: string;
  status?: string;
  payment?: OrderPayment;
}

interface BinaResponse {
  ResCode: number;
  ResMsg: string;
  docId?: number;
  docIdPayment?: number;
}

// ---------- Helpers ----------

function buildBinaPayload(req: CreateOrderRequest, token: string) {
  const requestId = Math.floor(Math.random() * 2_000_000_000);

  const custIdNum = typeof req.client.binaCustomerId === 'number'
    ? req.client.binaCustomerId
    : parseInt(String(req.client.binaCustomerId), 10);

  if (!custIdNum || Number.isNaN(custIdNum)) {
    throw new Error('binaCustomerId חייב להיות מספר תקין');
  }

  const payload: Record<string, unknown> = {
    tokenId: token,
    requestId,
    docType: 15,
    docWithvat: 0,
    docTitle: req.title || 'הזמנה חדשה',
    docStatus: req.status || 'חדשה',
    docRemark: req.remark || '',
    Cust: {
      custId: custIdNum,
      custName: req.client.name,
      custCity: req.client.city,
      custAddress: req.client.address,
      custIshKheser: req.client.contactPerson,
      custTel: req.client.phone || '',
      custEmail: req.client.email || '',
      custOrderId: req.client.customerOrderId || '',
    },
    docItems: req.items.map((it) => ({
      ItemId: it.itemId,
      ItemDesc: it.description,
      ItemQty: it.quantity,
      UnitPrice: String(it.unitPrice),
      Unitcurrency: 'ILS',
      CurValue: 1,
      Discount: it.discount ?? 0,
    })),
  };

  if (req.payment) {
    payload.Payments = {
      PayDate: req.payment.payDate,
      NumberOfPayments: req.payment.numberOfPayments,
      FirstPayment: req.payment.firstPayment,
      CreditName: req.payment.creditName,
      CreditDateEnd: req.payment.creditDateEnd || '',
      CreditNum: req.payment.creditNum || '',
      ReceiptNumber: req.payment.receiptNumber || '',
    };
  }

  return payload;
}

function validateRequest(req: CreateOrderRequest): string | null {
  if (!req.client) return 'חסר אובייקט client';
  if (!req.client.binaCustomerId) return 'חסר binaCustomerId';
  if (!req.client.name) return 'חסר שם לקוח';
  if (!req.client.city) return 'חסרה עיר';
  if (!req.client.address) return 'חסרה כתובת';
  if (!req.client.contactPerson) return 'חסר איש קשר (custIshKheser - שדה חובה בבינה)';
  if (!Array.isArray(req.items) || req.items.length === 0) return 'חסרים פריטים';

  for (let i = 0; i < req.items.length; i++) {
    const it = req.items[i];
    if (!it.itemId) return `פריט ${i + 1}: חסר itemId`;
    if (!it.description) return `פריט ${i + 1}: חסר תיאור`;
    if (!it.quantity || it.quantity <= 0) return `פריט ${i + 1}: כמות לא תקינה`;
    if (it.unitPrice == null || it.unitPrice < 0) return `פריט ${i + 1}: מחיר לא תקין`;
  }

  return null;
}

/** בינה מחזירים לפעמים array ולפעמים object - נורמליזציה לאובייקט */
function extractBinaResponse(text: string): BinaResponse | null {
  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed)) {
      return parsed.length > 0 ? parsed[0] as BinaResponse : null;
    }
    return parsed as BinaResponse;
  } catch {
    return null;
  }
}

// ---------- Main handler ----------

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response(null, { headers: corsHeaders });
  }

  if (req.method !== 'POST') {
    return new Response(JSON.stringify({ error: 'POST only' }), {
      status: 405,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    });
  }

  try {
    const token = Deno.env.get('BINA_TOKEN');
    if (!token?.trim()) {
      return new Response(
        JSON.stringify({ error: 'BINA_TOKEN לא מוגדר ב-Supabase Secrets' }),
        { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      );
    }

    const body = (await req.json()) as CreateOrderRequest;

    const validationError = validateRequest(body);
    if (validationError) {
      return new Response(JSON.stringify({ error: validationError }), {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });
    }

    // טוקן נשלח as-is, כמו ב-bina-create-customer שעובד
    const payload = buildBinaPayload(body, token.trim());

    // לוג בלי הטוקן
    const { tokenId: _, ...payloadForLog } = payload as { tokenId: string };
    console.log('[bina-create-order] שולח לבינה:', JSON.stringify(payloadForLog));

    const result = await fetchBinaViaQuotaGuard(BINA_URL, payload);

    const sanitizedPreview = result.text.slice(0, 500).replace(/U22g\w+/g, '[TOKEN_REDACTED]')
    console.log('[bina-create-order] תגובה מבינה:', {
      status: result.status,
      ok: result.ok,
      bodyPreview: sanitizedPreview,
    });

    if (!result.ok) {
      return new Response(
        JSON.stringify({
          success: false,
          error: `בינה החזירו HTTP ${result.status}`,
          binaResponse: result.text.slice(0, 500),
        }),
        { status: 502, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      );
    }

    const binaData = extractBinaResponse(result.text);

    if (!binaData) {
      return new Response(
        JSON.stringify({
          success: false,
          error: 'תגובה לא-JSON או ריקה מבינה',
          rawResponse: result.text.slice(0, 500),
        }),
        { status: 502, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      );
    }

    if (binaData.ResCode !== 0) {
      return new Response(
        JSON.stringify({
          success: false,
          error: `בינה דחו: ${binaData.ResMsg || 'ללא הודעה'}`,
          resCode: binaData.ResCode,
          fullResponse: binaData,
        }),
        { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      );
    }

    // הצלחה - יוצרים task מקושר
    const supabaseUrl = Deno.env.get('SUPABASE_URL');
    const supabaseKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY');

    let taskId: string | null = null;
    if (supabaseUrl && supabaseKey && binaData.docId) {
      try {
        const supabase = createClient(supabaseUrl, supabaseKey);
        const { data: task, error: taskErr } = await supabase
          .from('tasks')
          .insert({
            title: body.title || `הזמנה ${binaData.docId}`,
            client_name: body.client.name,
            status: 'חדש',
            dept: 'כללי',
            priority: 'בינונית',
            source: 'manual',
            bina_order_id: String(binaData.docId),
            notes: body.remark || '',
          })
          .select('id')
          .single();

        if (!taskErr && task) {
          taskId = task.id;
        } else if (taskErr) {
          console.error('[bina-create-order] שגיאה ביצירת task:', taskErr);
        }
      } catch (e) {
        console.error('[bina-create-order] חריגה ביצירת task:', e);
      }
    }

    return new Response(
      JSON.stringify({
        success: true,
        binaOrderId: binaData.docId,
        binaPaymentId: binaData.docIdPayment,
        taskId,
      }),
      { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    );
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error('[bina-create-order] חריגה כללית:', msg);
    return new Response(JSON.stringify({ error: msg }), {
      status: 500,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    });
  }
});
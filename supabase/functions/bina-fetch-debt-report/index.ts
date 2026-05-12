// supabase/functions/bina-fetch-debt-report/index.ts
// שולפת דוח חייבים מבינה (docType -900) ושומרת snapshot
// מקושרת לטבלת clients דרך bina_customer_id

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';
import { fetchBinaViaQuotaGuard } from '../_shared/bina-proxy-fetch.ts';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

const BINA_URL = 'https://webfiles.binaw.com/post/PostJsonDocV2.aspx';
const ORDERS_QUERY_URL = 'https://webapps.binaw.com/PostJsonDocQuery.aspx';

interface BinaDebtRecord {
  custId: number;
  docNum: number;
  docDate: string;          // dd/MM/yyyy
  docPaymentDate: string;   // dd/MM/yyyy
  docTotal: number;
  docBalance: number;
}

interface ClientRecord {
  id: string;
  name: string;
  bina_customer_id: string;
}

interface TaskClientHint {
  name: string;
  address: string;
  city: string;
}

// ---------- Helpers ----------

/** ממיר תאריך מ-dd/MM/yyyy ל-yyyy-MM-dd (פורמט SQL) */
function parseHebrewDate(dateStr: string | null | undefined): string | null {
  if (!dateStr || typeof dateStr !== 'string') return null;
  const parts = dateStr.trim().split('/');
  if (parts.length !== 3) return null;
  const [dd, mm, yyyy] = parts;
  if (!dd || !mm || !yyyy) return null;
  return `${yyyy}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`;
}

/** מחשב כמה ימים באיחור (חיובי = באיחור, אפס/שלילי = לא באיחור) */
function calcDaysOverdue(paymentDateSQL: string | null): number {
  if (!paymentDateSQL) return 0;
  const paymentDate = new Date(paymentDateSQL);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const diffMs = today.getTime() - paymentDate.getTime();
  const days = Math.floor(diffMs / (1000 * 60 * 60 * 24));
  return Math.max(0, days);
}

// ---------- Main handler ----------

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response(null, { headers: corsHeaders });
  }

  if (req.method !== 'POST' && req.method !== 'GET') {
    return new Response(JSON.stringify({ error: 'POST or GET only' }), {
      status: 405,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    });
  }

  try {
    const token = Deno.env.get('BINA_TOKEN');
    if (!token?.trim()) {
      return new Response(
        JSON.stringify({ error: 'BINA_TOKEN לא מוגדר' }),
        { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      );
    }

    const supabaseUrl = Deno.env.get('SUPABASE_URL')!;
    const supabaseKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
    const supabase = createClient(supabaseUrl, supabaseKey);

    const tasksClientMap = new Map<string, TaskClientHint>();
    {
      const pageSize = 1000;
      let from = 0;
      for (;;) {
        const { data: taskPage, error: tasksMapErr } = await supabase
          .from('tasks')
          .select('bina_cust_id, client_name, bina_cust_address, bina_cust_city')
          .not('bina_cust_id', 'is', null)
          .not('client_name', 'is', null)
          .range(from, from + pageSize - 1);

        if (tasksMapErr) {
          console.error('[debt-report] שגיאה בטעינת מפת לקוחות מ-tasks:', tasksMapErr.message);
          break;
        }
        const rows = taskPage || [];
        for (const t of rows) {
          const key = String(t.bina_cust_id);
          if (!tasksClientMap.has(key) && t.client_name) {
            tasksClientMap.set(key, {
              name: String(t.client_name).trim(),
              address: (t.bina_cust_address && String(t.bina_cust_address)) || '',
              city: (t.bina_cust_city && String(t.bina_cust_city)) || '',
            });
          }
        }
        if (rows.length < pageSize) break;
        from += pageSize;
      }
      console.log(`[debt-report] טען ${tasksClientMap.size} לקוחות ממשימות היסטוריות`);
    }
    const clientHintsFromTasksDb = new Set(tasksClientMap.keys());

    // אופציה: לקוח ספציפי (אם נשלח בגוף הבקשה)
    let specificCustId: number | null = null;
    if (req.method === 'POST') {
      try {
        const body = await req.json();
        if (body?.custId) {
          specificCustId = parseInt(String(body.custId), 10);
        }
      } catch {
        // אין body - שולפים הכל
      }
    }

    // --- שלב 1: שליפה מבינה ---
    const binaPayload: Record<string, unknown> = {
      tokenId: token.trim(),
      docType: -900,
    };
    if (specificCustId) {
      binaPayload.custId = specificCustId;
    }

    console.log('[debt-report] שולח לבינה:', { docType: -900, custId: specificCustId || 'ALL' });

    const result = await fetchBinaViaQuotaGuard(BINA_URL, binaPayload);

    // סינון הטוקן מהתגובה לפני לוג
    const sanitized = result.text.slice(0, 500).replace(/U22g\w+/g, '[TOKEN_REDACTED]');
    console.log('[debt-report] תגובה מבינה:', {
      status: result.status,
      ok: result.ok,
      bodyPreview: sanitized,
    });

    if (!result.ok) {
      return new Response(
        JSON.stringify({
          success: false,
          error: `בינה החזירו HTTP ${result.status}`,
          details: sanitized,
        }),
        { status: 502, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      );
    }

    let binaDebts: BinaDebtRecord[];
    try {
      const parsed = JSON.parse(result.text);
      // deno-lint-ignore no-explicit-any
      const data: any = parsed;
      console.log(
        '[DEBUG-FIRST-RECORD]',
        JSON.stringify(data[0] || data.Records?.[0] || data.Items?.[0] || data, null, 2).substring(0, 2000),
      );

      // בדיקת שגיאה - בינה מחזירים גם array וגם object, וגם ResCode/resCode (case insensitive)
      const errorCheck = Array.isArray(parsed) ? parsed[0] : parsed;
      const errCode = errorCheck?.ResCode ?? errorCheck?.resCode;
      const errMsg = errorCheck?.ResMsg ?? errorCheck?.resMsg;
      if (errCode !== undefined && errCode !== 0) {
        return new Response(
          JSON.stringify({
            success: false,
            error: `בינה דחו: ${errMsg || 'ללא הודעה'}`,
            resCode: errCode,
          }),
          { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
        );
      }

      // אם בינה החזירו אובייקט שגיאה במקום מערך (לא אמור לקרות אחרי הבדיקה למעלה)
      if (!Array.isArray(parsed)) {
        const pErr = parsed as Record<string, unknown>;
        const pCode = pErr?.ResCode ?? pErr?.resCode;
        const pMsg = pErr?.ResMsg ?? pErr?.resMsg;
        if (pCode !== undefined && pCode !== 0) {
          return new Response(
            JSON.stringify({
              success: false,
              error: `בינה דחו: ${pMsg || 'ללא הודעה'}`,
              resCode: pCode,
            }),
            { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
          );
        }
        return new Response(
          JSON.stringify({
            success: false,
            error: 'תגובה לא צפויה מבינה',
            details: sanitized,
          }),
          { status: 502, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
        );
      }
      binaDebts = parsed;
    } catch {
      return new Response(
        JSON.stringify({
          success: false,
          error: 'תגובה לא-JSON מבינה',
          rawResponse: sanitized,
        }),
        { status: 502, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      );
    }

    console.log(`[debt-report] התקבלו ${binaDebts.length} רשומות מבינה`);

    // --- שלב 2: שליפת לקוחות מ-DB לקישור ---
    const { data: clientsData, error: clientsErr } = await supabase
      .from('clients')
      .select('id, name, bina_customer_id')
      .not('bina_customer_id', 'is', null);

    if (clientsErr) {
      console.error('[debt-report] שגיאה בשליפת לקוחות:', clientsErr);
      return new Response(
        JSON.stringify({ success: false, error: `שגיאה ב-DB: ${clientsErr.message}` }),
        { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      );
    }

    // מפת bina_customer_id -> client record
    const clientsMap = new Map<string, ClientRecord>();
    for (const c of (clientsData || []) as ClientRecord[]) {
      if (c.bina_customer_id) {
        clientsMap.set(String(c.bina_customer_id), c);
      }
    }

    console.log(`[debt-report] נטענו ${clientsMap.size} לקוחות עם bina_customer_id`);

    // --- שלב 3: עיבוד ובניית רשומות לשמירה ---
    const today = new Date().toISOString().split('T')[0]; // yyyy-MM-dd

    const snapshotRecords = [];
    const missingClients = new Map<string, { firstSeen: number; debtTotal: number; invoiceCount: number }>();

    for (const debt of binaDebts) {
      const docPaymentDateSQL = parseHebrewDate(debt.docPaymentDate);
      const daysOverdue = calcDaysOverdue(docPaymentDateSQL);
      const balance = debt.docBalance ?? 0;
      const isOverdue = daysOverdue > 0 && balance > 0;

      // מדלגים על רשומות ריקות לחלוטין (ללא custId וללא יתרה)
      if (!debt.custId && !balance) continue;

      const custIdStr = String(debt.custId);
      const client = clientsMap.get(custIdStr);

      snapshotRecords.push({
        snapshot_date: today,
        bina_customer_id: custIdStr,
        customer_name: client?.name || null,
        doc_num: debt.docNum || null,  // יכול להיות null ביתרת פתיחה
        doc_date: parseHebrewDate(debt.docDate),
        doc_payment_date: docPaymentDateSQL,
        doc_total: debt.docTotal ?? null,
        doc_balance: balance,
        is_overdue: isOverdue,
        days_overdue: daysOverdue,
        client_id: client?.id || null,
        client_exists_in_db: !!client,
      });

      // אם לקוח חסר - לסטטיסטיקה
      if (!client && balance > 0) {
        const existing = missingClients.get(custIdStr);
        if (existing) {
          existing.debtTotal += balance;
          existing.invoiceCount += 1;
        } else {
          missingClients.set(custIdStr, {
            firstSeen: debt.custId,
            debtTotal: balance,
            invoiceCount: 1,
          });
        }
      }
    }

    let namesFromHistoricalOrders = 0;
    const missingIds = Array.from(missingClients.keys())
      .map((id) => parseInt(id, 10))
      .filter((n) => !Number.isNaN(n));
    const missingIdSet = new Set(missingIds);

    if (missingIds.length > 0) {
      console.log(`[debt-report] מנסה למצוא שמות ל-${missingIds.length} לקוחות חסרים`);

      const fromDate5y = new Date();
      fromDate5y.setFullYear(fromDate5y.getFullYear() - 5);
      const toDate6m = new Date();
      toDate6m.setMonth(toDate6m.getMonth() + 6);

      const ordersResult = await fetchBinaViaQuotaGuard(ORDERS_QUERY_URL, {
        tokenId: token.trim(),
        docType: -15,
        fromDate: fromDate5y.toISOString().split('T')[0],
        toDate: toDate6m.toISOString().split('T')[0],
      });

      if (ordersResult.ok) {
        try {
          const ordersData = JSON.parse(ordersResult.text) as Record<string, unknown> | unknown[];
          const orders = Array.isArray(ordersData)
            ? ordersData
            : (Array.isArray((ordersData as Record<string, unknown>)?.Orders)
              ? ((ordersData as Record<string, unknown>).Orders as unknown[])
              : []);
          console.log(`[debt-report] התקבלו ${orders.length} הזמנות היסטוריות מבינה`);

          for (const order of orders) {
            const o = order as Record<string, unknown>;
            const rawId = o.custId ?? o.CustId;
            if (rawId === undefined || rawId === null) continue;
            const cid = String(rawId);
            const n = parseInt(cid, 10);
            if (!missingIdSet.has(n)) continue;

            const custName = (o.custName ?? o.CustName) as string | undefined;
            if (!custName || !String(custName).trim()) continue;

            if (!tasksClientMap.has(cid)) {
              tasksClientMap.set(cid, {
                name: String(custName).trim(),
                address: String(o.custAddress ?? o.CustAddress ?? '') || '',
                city: String(o.custCity ?? o.CustCity ?? '') || '',
              });
              namesFromHistoricalOrders++;
            }
          }
          console.log(
            `[debt-report] נמצאו שמות מבינה (הזמנות היסטוריות) ל-${namesFromHistoricalOrders} לקוחות חסרים`,
          );
        } catch (e) {
          console.error('[debt-report] שגיאה בפענוח הזמנות היסטוריות:', e);
        }
      } else {
        console.warn('[debt-report] שאילתת הזמנות היסטורית נכשלה:', ordersResult.status);
      }
    }

    let createdClients = 0;
    let foundInTasks = 0;

    for (const [custId] of missingClients.entries()) {
      if (clientsMap.has(custId)) continue;

      const { data: existingRow } = await supabase
        .from('clients')
        .select('id, name, bina_customer_id')
        .eq('bina_customer_id', custId)
        .maybeSingle();

      if (existingRow) {
        clientsMap.set(custId, existingRow as ClientRecord);
        continue;
      }

      const taskClient = tasksClientMap.get(custId);
      const addrParts = [taskClient?.city, taskClient?.address].filter(Boolean);
      const fullAddress = addrParts.length ? addrParts.join(', ') : (taskClient?.address || '');

      const newClient = {
        name: taskClient?.name || `לקוח #${custId}`,
        bina_customer_id: String(custId),
        address: fullAddress || null,
        notes: taskClient
          ? 'נוצר אוטומטית מדוח חייבים (שם ממשימה היסטורית)'
          : 'נוצר אוטומטית מדוח חייבים - שם זמני',
      };

      const { data: insData, error: insErr } = await supabase
        .from('clients')
        .insert(newClient)
        .select('id, name, bina_customer_id')
        .maybeSingle();

      if (!insErr && insData) {
        createdClients++;
        if (taskClient && clientHintsFromTasksDb.has(custId)) foundInTasks++;
        clientsMap.set(custId, {
          id: insData.id,
          name: insData.name,
          bina_customer_id: String(insData.bina_customer_id ?? custId),
        });
      } else if (insErr) {
        const dup =
          insErr.code === '23505' ||
          (insErr.message && /duplicate|unique/i.test(insErr.message));
        if (dup) {
          const { data: again } = await supabase
            .from('clients')
            .select('id, name, bina_customer_id')
            .eq('bina_customer_id', custId)
            .maybeSingle();
          if (again) clientsMap.set(custId, again as ClientRecord);
          else console.error(`[debt-report] שגיאה ביצירת לקוח ${custId} (כפילות ללא fetch):`, insErr.message);
        } else {
          console.error(`[debt-report] שגיאה ביצירת לקוח ${custId}:`, insErr.message);
        }
      }
    }

    for (const r of snapshotRecords) {
      if (!r.client_exists_in_db && r.bina_customer_id) {
        const c = clientsMap.get(r.bina_customer_id);
        if (c) {
          r.customer_name = c.name;
          r.client_id = c.id;
          r.client_exists_in_db = true;
        }
      }
    }

    // --- שלב 4: שמירה ב-DB (upsert לפי snapshot_date + custId + docNum) ---
    if (snapshotRecords.length > 0) {
      // מוחקים את ה-snapshot של היום (אם קיים) ושמים מחדש
      const { error: deleteErr } = await supabase
        .from('debt_snapshots')
        .delete()
        .eq('snapshot_date', today);

      if (deleteErr) {
        console.error('[debt-report] שגיאה במחיקת snapshot ישן:', deleteErr);
      }

      // הכנסה במקטעים של 500
      const chunkSize = 500;
      let totalInserted = 0;
      for (let i = 0; i < snapshotRecords.length; i += chunkSize) {
        const chunk = snapshotRecords.slice(i, i + chunkSize);
        const { error: insertErr } = await supabase
          .from('debt_snapshots')
          .insert(chunk);

        if (insertErr) {
          console.error('[debt-report] שגיאה בהכנסת chunk:', insertErr);
          return new Response(
            JSON.stringify({
              success: false,
              error: `שגיאה בשמירה: ${insertErr.message}`,
              insertedSoFar: totalInserted,
            }),
            { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
          );
        }
        totalInserted += chunk.length;
      }

      console.log(`[debt-report] נשמרו ${totalInserted} רשומות`);
    }

    // --- שלב 5: סיכום למסך ---
    const totalDebt = snapshotRecords
      .filter(r => r.doc_balance > 0)
      .reduce((sum, r) => sum + Number(r.doc_balance), 0);

    const overdueDebt = snapshotRecords
      .filter(r => r.is_overdue)
      .reduce((sum, r) => sum + Number(r.doc_balance), 0);

    const customersWithDebt = new Set(
      snapshotRecords.filter(r => r.doc_balance > 0).map(r => r.bina_customer_id)
    ).size;

    const overdueCustomers = new Set(
      snapshotRecords.filter(r => r.is_overdue).map(r => r.bina_customer_id)
    ).size;

    try {
      const detectUrl = `${supabaseUrl}/functions/v1/detect-and-notify-payments`;
      const detectResp = await fetch(detectUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${supabaseKey}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({}),
      });
      const detectResult = await detectResp.json();
      console.log('[bina-fetch-debt-report] detect-payments result:', JSON.stringify(detectResult.summary || detectResult));
    } catch (detectErr) {
      console.error('[bina-fetch-debt-report] detect-payments failed:', detectErr);
    }

    return new Response(
      JSON.stringify({
        success: true,
        snapshot_date: today,
        summary: {
          total_records: snapshotRecords.length,
          customers_with_debt: customersWithDebt,
          total_open_debt: Math.round(totalDebt * 100) / 100,
          overdue_customers: overdueCustomers,
          total_overdue_debt: Math.round(overdueDebt * 100) / 100,
        },
        missing_clients: {
          count: missingClients.size,
          // 20 הלקוחות הראשונים שחסרים, ממוינים לפי גודל החוב
          top_missing: Array.from(missingClients.entries())
            .map(([custId, data]) => ({
              bina_customer_id: custId,
              total_debt: Math.round(data.debtTotal * 100) / 100,
              open_invoices: data.invoiceCount,
            }))
            .sort((a, b) => b.total_debt - a.total_debt)
            .slice(0, 20),
        },
        created_clients: createdClients,
        found_names_from_tasks: foundInTasks,
        names_from_historical_orders: namesFromHistoricalOrders,
      }),
      { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    );

  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error('[debt-report] חריגה כללית:', msg);
    return new Response(
      JSON.stringify({ error: msg }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    );
  }
});

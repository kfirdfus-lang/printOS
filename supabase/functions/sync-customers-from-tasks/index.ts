// supabase/functions/sync-customers-from-tasks/index.ts
// סקריפט חד-פעמי - סורק את כל ה-tasks עם bina_cust_id וממלא את טבלת clients
//
// הלוגיקה לכל לקוח:
// 1. אם כבר קיים ב-clients עם אותו bina_customer_id → דילוג
// 2. אם קיים לקוח עם שם דומה אבל בלי קוד → עדכון bina_customer_id + הפרטים
// 3. אם לא קיים בכלל → יצירה חדשה

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

interface TaskCustomerData {
  bina_cust_id: string;
  client_name: string;
  bina_cust_address: string | null;
  bina_cust_city: string | null;
  contact: string | null;
}

interface ClientRecord {
  id: string;
  name: string;
  phone: string | null;
  address: string | null;
  bina_customer_id: string | null;
}

/** נורמליזציה של שם לקוח להשוואה (להסיר רווחים, לאיחוד) */
function normalizeName(name: string): string {
  if (!name) return '';
  return name.trim().replace(/\s+/g, ' ').toLowerCase();
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const supabaseUrl = Deno.env.get('SUPABASE_URL')!;
    const supabaseKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
    const supabase = createClient(supabaseUrl, supabaseKey);

    // ----- שלב 1: שליפת כל ה-tasks עם bina_cust_id -----
    console.log('[sync-customers] שולף משימות עם bina_cust_id...');
    
    const { data: tasksData, error: tasksErr } = await supabase
      .from('tasks')
      .select('bina_cust_id, client_name, bina_cust_address, bina_cust_city, contact')
      .not('bina_cust_id', 'is', null)
      .not('client_name', 'is', null);

    if (tasksErr) {
      throw new Error(`שגיאה בשליפת משימות: ${tasksErr.message}`);
    }

    const tasks = (tasksData || []) as TaskCustomerData[];
    console.log(`[sync-customers] נטענו ${tasks.length} משימות עם פרטי לקוח`);

    // ----- שלב 2: קיבוץ לקוחות ייחודיים לפי bina_cust_id -----
    // משתמשים ב-Map כדי לקבל את הלקוח הראשון שראינו לכל קוד
    const uniqueCustomers = new Map<string, TaskCustomerData>();
    for (const task of tasks) {
      if (!task.bina_cust_id) continue;
      const key = String(task.bina_cust_id);
      if (!uniqueCustomers.has(key)) {
        uniqueCustomers.set(key, task);
      }
    }

    console.log(`[sync-customers] ${uniqueCustomers.size} לקוחות ייחודיים`);

    // ----- שלב 3: שליפת כל הלקוחות הקיימים -----
    const { data: clientsData, error: clientsErr } = await supabase
      .from('clients')
      .select('id, name, phone, address, bina_customer_id');

    if (clientsErr) {
      throw new Error(`שגיאה בשליפת לקוחות: ${clientsErr.message}`);
    }

    const allClients = (clientsData || []) as ClientRecord[];
    
    // מפה: bina_customer_id → client (לקוחות שכבר מקושרים)
    const linkedByBinaId = new Map<string, ClientRecord>();
    // מפה: שם נורמלי → client (לקוחות בלי קישור, להצלבה לפי שם)
    const unlinkedByName = new Map<string, ClientRecord>();

    for (const c of allClients) {
      if (c.bina_customer_id) {
        linkedByBinaId.set(String(c.bina_customer_id), c);
      } else {
        const normName = normalizeName(c.name);
        if (normName) {
          unlinkedByName.set(normName, c);
        }
      }
    }

    console.log(`[sync-customers] ב-DB: ${linkedByBinaId.size} מקושרים, ${unlinkedByName.size} לא מקושרים`);

    // ----- שלב 4: עיבוד כל לקוח ייחודי -----
    const stats = {
      already_linked: 0,        // כבר היה מקושר - לא נגענו
      linked_existing: 0,        // היה קיים בלי קוד - הוספנו לו את הקוד
      created_new: 0,            // יצרנו לקוח חדש
      errors: 0,
    };

    const errorDetails: string[] = [];
    const linkedCustomers: Array<{ binaId: string; name: string; action: string }> = [];

    for (const [binaIdStr, taskData] of uniqueCustomers) {
      try {
        const customerName = (taskData.client_name || '').trim();
        if (!customerName) continue;

        // האם כבר מקושר?
        if (linkedByBinaId.has(binaIdStr)) {
          stats.already_linked++;
          continue;
        }

        // האם יש לקוח לא-מקושר עם שם דומה?
        const normName = normalizeName(customerName);
        const existingClient = unlinkedByName.get(normName);

        // בניית הכתובת (אם יש פרטים)
        const addressParts: string[] = [];
        if (taskData.bina_cust_address) addressParts.push(taskData.bina_cust_address);
        if (taskData.bina_cust_city) addressParts.push(taskData.bina_cust_city);
        const fullAddress = addressParts.join(', ') || null;

        if (existingClient) {
          // עדכון לקוח קיים - מוסיפים קוד בינה
          const updates: Record<string, unknown> = {
            bina_customer_id: binaIdStr,
          };
          
          // אם אין כתובת ל-client, נוסיף מהמשימה
          if (!existingClient.address && fullAddress) {
            updates.address = fullAddress;
          }

          const { error: updErr } = await supabase
            .from('clients')
            .update(updates)
            .eq('id', existingClient.id);

          if (updErr) {
            stats.errors++;
            errorDetails.push(`עדכון ${customerName}: ${updErr.message}`);
          } else {
            stats.linked_existing++;
            linkedCustomers.push({
              binaId: binaIdStr,
              name: customerName,
              action: 'הוספת קוד בינה',
            });
            // נסיר מהמפה כדי שלא נעדכן שוב
            unlinkedByName.delete(normName);
          }
        } else {
          // יצירת לקוח חדש
          const { error: insertErr } = await supabase
            .from('clients')
            .insert({
              name: customerName,
              address: fullAddress,
              bina_customer_id: binaIdStr,
              notes: 'נוצר אוטומטית מסנכרון בינה',
            });

          if (insertErr) {
            stats.errors++;
            errorDetails.push(`יצירת ${customerName}: ${insertErr.message}`);
          } else {
            stats.created_new++;
            linkedCustomers.push({
              binaId: binaIdStr,
              name: customerName,
              action: 'יצירה חדשה',
            });
          }
        }
      } catch (err) {
        stats.errors++;
        const msg = err instanceof Error ? err.message : String(err);
        errorDetails.push(`חריגה: ${msg}`);
      }
    }

    console.log('[sync-customers] סיכום:', stats);

    return new Response(
      JSON.stringify({
        success: true,
        summary: {
          total_tasks_scanned: tasks.length,
          unique_customers_in_tasks: uniqueCustomers.size,
          ...stats,
        },
        linked_customers: linkedCustomers.slice(0, 50), // 50 ראשונים לדוגמה
        errors: errorDetails.slice(0, 10),
      }),
      { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    );

  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error('[sync-customers] חריגה:', msg);
    return new Response(
      JSON.stringify({ success: false, error: msg }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    );
  }
});

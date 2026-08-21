// supabase/functions/update-quotes-status/index.ts
// רץ כל 10 דקות (cron) או ידנית
// 1. מוצא הצעות שהפכו להזמנות (zhe תואמת ב-tasks)
// 2. מסמן הצעות "תקועות" (4+ ימים בלי סגירה)
// 3. מסמן הצעות "פגות תוקף" (30+ ימים)

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';
import { rejectDisallowedInternalOrigin } from '../_shared/cors.ts';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

interface Quote {
  id: string;
  bina_doc_id: number | null;
  bina_cust_id: number | null;
  bina_cust_name: string | null;
  total_amount: number | null;
  auto_status: string;
  created_at: string;
}

interface Task {
  id: string;
  bina_order_id: number | null;
  bina_cust_id: number | null;
  client_name: string | null;
  total_amount: number | null;
  total_inc_vat: number | null;
  bina_order_date: string | null;
  created_at: string;
}

/** האם 2 סכומים דומים (תוך 5% הפרש) */
function similarAmount(a: number | null | undefined, b: number | null | undefined): boolean {
  if (!a || !b) return false;
  const diff = Math.abs(a - b);
  const max = Math.max(a, b);
  return (diff / max) <= 0.05;
}

Deno.serve(async (req) => {
  const originBlock = rejectDisallowedInternalOrigin(req);
  if (originBlock) return originBlock;

  if (req.method === 'OPTIONS') {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const supabaseUrl = Deno.env.get('SUPABASE_URL')!;
    const supabaseKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
    const supabase = createClient(supabaseUrl, supabaseKey);

    const now = new Date();
    const stats = {
      checked: 0,
      closed: 0,
      stuck: 0,
      expired: 0,
      errors: 0,
    };

    // ----- שלב 1: שליפת הצעות פעילות -----
    const { data: activeQuotes, error: quotesErr } = await supabase
      .from('quotes')
      .select('id, bina_doc_id, bina_cust_id, bina_cust_name, total_amount, auto_status, created_at')
      .eq('is_archive', false)
      .in('auto_status', ['ממתינה', 'תקועה']);

    if (quotesErr) {
      throw new Error(`שגיאה בשליפת הצעות: ${quotesErr.message}`);
    }

    const quotes = (activeQuotes || []) as Quote[];
    stats.checked = quotes.length;
    console.log(`[update-quotes] בודק ${quotes.length} הצעות פעילות`);

    if (quotes.length === 0) {
      return new Response(JSON.stringify({
        success: true,
        message: 'אין הצעות פעילות לבדיקה',
        stats,
      }), { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    // ----- שלב 2: שליפת tasks רלוונטיות לבדיקת סגירה -----
    // רק לקוחות שיש להם הצעות פעילות
    const custIds = Array.from(new Set(
      quotes.map(q => q.bina_cust_id).filter(id => id !== null)
    ));

    let candidateTasks: Task[] = [];
    if (custIds.length > 0) {
      const { data: tasksData, error: tasksErr } = await supabase
        .from('tasks')
        .select('id, bina_order_id, bina_cust_id, client_name, total_amount, total_inc_vat, bina_order_date, created_at')
        .in('bina_cust_id', custIds)
        .not('total_amount', 'is', null)
        .gte('created_at', new Date(now.getTime() - 35 * 86400000).toISOString());

      if (tasksErr) {
        console.error('[update-quotes] שגיאה בשליפת tasks:', tasksErr);
      } else {
        candidateTasks = (tasksData || []) as Task[];
      }
    }

    console.log(`[update-quotes] נטענו ${candidateTasks.length} tasks פוטנציאליות`);

    // ----- שלב 3: עיבוד כל הצעה -----
    const updates: Array<{ id: string; data: Record<string, unknown> }> = [];

    for (const quote of quotes) {
      const quoteCreatedAt = new Date(quote.created_at);
      const ageMs = now.getTime() - quoteCreatedAt.getTime();
      const ageDays = ageMs / 86400000;

      // בדיקה: האם פג תוקף (30+ ימים)
      if (ageDays >= 30) {
        updates.push({
          id: quote.id,
          data: {
            auto_status: 'פגת תוקף',
            last_status_check: now.toISOString(),
          }
        });
        stats.expired++;
        continue;
      }

      // בדיקה: האם הצעה נסגרה (יש task תואם)
      let matchingTask: Task | null = null;

      if (quote.bina_cust_id && quote.total_amount) {
        // חיפוש task שעונה ל:
        // - אותו לקוח
        // - נוצר אחרי ההצעה
        // - בתוך 30 יום מההצעה
        // - סכום דומה (5%)
        for (const task of candidateTasks) {
          if (task.bina_cust_id !== quote.bina_cust_id) continue;
          
          const taskDate = task.bina_order_date 
            ? new Date(task.bina_order_date) 
            : new Date(task.created_at);
          
          // task צריך להיות אחרי ההצעה
          if (taskDate < quoteCreatedAt) continue;
          
          // בתוך 30 יום מההצעה
          const daysDiff = (taskDate.getTime() - quoteCreatedAt.getTime()) / 86400000;
          if (daysDiff > 30) continue;
          
          // סכום דומה (משווים גם לפני וגם אחרי מעמ)
          const taskAmount = task.total_amount;
          const taskAmountIncVat = task.total_inc_vat;
          
          if (similarAmount(quote.total_amount, taskAmount) || 
              similarAmount(quote.total_amount, taskAmountIncVat) ||
              similarAmount(quote.total_amount * 1.18, taskAmountIncVat)) {
            matchingTask = task;
            break;
          }
        }
      }

      if (matchingTask) {
        updates.push({
          id: quote.id,
          data: {
            auto_status: 'נסגרה',
            closed_at: matchingTask.bina_order_date || matchingTask.created_at,
            closed_by_task_id: matchingTask.id,
            last_status_check: now.toISOString(),
          }
        });
        stats.closed++;
        continue;
      }

      // בדיקה: האם תקועה (4+ ימים בלי סגירה)
      if (ageDays >= 4 && quote.auto_status !== 'תקועה') {
        updates.push({
          id: quote.id,
          data: {
            auto_status: 'תקועה',
            last_status_check: now.toISOString(),
          }
        });
        stats.stuck++;
        continue;
      }

      // אחרת - רק עדכון last_status_check
      updates.push({
        id: quote.id,
        data: { last_status_check: now.toISOString() }
      });
    }

    // ----- שלב 4: ביצוע עדכונים -----
    for (const upd of updates) {
      const { error: updErr } = await supabase
        .from('quotes')
        .update(upd.data)
        .eq('id', upd.id);

      if (updErr) {
        stats.errors++;
        console.error(`[update-quotes] שגיאה בעדכון ${upd.id}:`, updErr);
      }
    }

    console.log('[update-quotes] סיכום:', stats);

    return new Response(JSON.stringify({
      success: true,
      stats,
      timestamp: now.toISOString(),
    }), { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });

  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error('[update-quotes] חריגה:', msg);
    return new Response(JSON.stringify({ success: false, error: msg }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
  }
});

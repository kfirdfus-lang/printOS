import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const supabaseUrl = Deno.env.get('SUPABASE_URL')!;
    const supabaseKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
    const supabase = createClient(supabaseUrl, supabaseKey);

    // הקלט: { customers: [{bina_id: '3122', name: 'שם'}, ...] }
    const body = await req.json();
    const customers = body.customers || [];

    if (!Array.isArray(customers) || customers.length === 0) {
      return new Response(JSON.stringify({ error: 'אין רשימת לקוחות לעדכון' }),
        { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    console.log(`[import-csv] קיבל ${customers.length} לקוחות`);

    let updated = 0;
    let created = 0;
    let skipped = 0;
    const errors = [];

    for (const c of customers) {
      const binaId = String(c.bina_id || '').trim();
      const name = String(c.name || '').trim();
      
      if (!binaId || !name) {
        skipped++;
        continue;
      }

      // בדיקה אם כבר קיים
      const { data: existing } = await supabase
        .from('clients')
        .select('id, name')
        .eq('bina_customer_id', binaId)
        .maybeSingle();

      if (existing) {
        // אם השם זמני (לקוח #X) - עדכן עם השם האמיתי
        // אם יש כבר שם אמיתי - לא נדרוס
        if (existing.name && existing.name.startsWith('לקוח #')) {
          const { error: updErr } = await supabase
            .from('clients')
            .update({
              name: name,
              notes: 'שם עודכן מקובץ CSV של דוח חייבים',
            })
            .eq('id', existing.id);
          
          if (updErr) {
            errors.push(`${binaId}: ${updErr.message}`);
          } else {
            updated++;
            // עדכון debt_snapshots
            await supabase
              .from('debt_snapshots')
              .update({ customer_name: name })
              .eq('bina_customer_id', binaId);
          }
        } else {
          skipped++;
        }
      } else {
        // יצירת לקוח חדש
        const { error: insErr } = await supabase
          .from('clients')
          .insert({
            name: name,
            bina_customer_id: binaId,
            notes: 'נוצר מקובץ CSV של דוח חייבים',
          });
        
        if (insErr) {
          errors.push(`${binaId}: ${insErr.message}`);
        } else {
          created++;
          await supabase
            .from('debt_snapshots')
            .update({ customer_name: name })
            .eq('bina_customer_id', binaId);
        }
      }
    }

    return new Response(JSON.stringify({
      success: true,
      summary: {
        total: customers.length,
        updated,
        created,
        skipped,
        errors: errors.length,
      },
      sample_errors: errors.slice(0, 5),
    }), { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });

  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return new Response(JSON.stringify({ success: false, error: msg }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
  }
});

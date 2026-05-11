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

    const body = await req.json();
    const customers = body.customers || [];

    if (!Array.isArray(customers) || customers.length === 0) {
      return new Response(JSON.stringify({ error: 'אין לקוחות' }),
        { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    let updated = 0;
    let notFound = 0;
    const errors: string[] = [];

    for (const c of customers) {
      const binaId = String(c.bina_customer_id || '').trim();
      if (!binaId) continue;

      const updates: Record<string, string> = {};
      if (c.payment_terms) updates.payment_terms = String(c.payment_terms).trim();
      if (c.contact_name) updates.contact_name = String(c.contact_name).trim();
      if (c.phone) updates.phone = String(c.phone).trim();
      if (c.email) updates.email = String(c.email).trim();

      if (Object.keys(updates).length === 0) continue;

      // בדיקה אם הלקוח קיים
      const { data: existing } = await supabase
        .from('clients')
        .select('id')
        .eq('bina_customer_id', binaId)
        .maybeSingle();

      if (existing) {
        const { error: updErr } = await supabase
          .from('clients')
          .update(updates)
          .eq('id', existing.id);

        if (updErr) {
          errors.push(`${binaId}: ${updErr.message}`);
        } else {
          updated++;
        }
      } else {
        notFound++;
      }
    }

    return new Response(JSON.stringify({
      success: true,
      summary: {
        total: customers.length,
        updated,
        not_found: notFound,
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

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';
import { fetchBinaViaQuotaGuard } from '../_shared/bina-proxy-fetch.ts';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

const ORDERS_QUERY_URL = 'https://webapps.binaw.com/PostJsonDocQuery.aspx';

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const token = Deno.env.get('BINA_TOKEN');
    if (!token?.trim()) {
      return new Response(JSON.stringify({ error: 'BINA_TOKEN לא מוגדר' }),
        { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    const supabaseUrl = Deno.env.get('SUPABASE_URL')!;
    const supabaseKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
    const supabase = createClient(supabaseUrl, supabaseKey);

    const { data: tempClients, error: clientsErr } = await supabase
      .from('clients')
      .select('id, name, bina_customer_id, address')
      .like('name', 'לקוח #%')
      .not('bina_customer_id', 'is', null);

    if (clientsErr) throw new Error(`שגיאה בשליפת לקוחות: ${clientsErr.message}`);

    const tempClientList = tempClients || [];
    console.log(`[update-names] נמצאו ${tempClientList.length} לקוחות עם שמות זמניים`);

    if (tempClientList.length === 0) {
      return new Response(JSON.stringify({
        success: true,
        message: 'אין לקוחות עם שמות זמניים לעדכון',
        updated: 0,
      }), { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    const fromDate = new Date();
    fromDate.setFullYear(fromDate.getFullYear() - 5);
    const toDate = new Date();
    toDate.setMonth(toDate.getMonth() + 6);

    console.log(`[update-names] שולח לבינה: ${fromDate.toISOString().split('T')[0]} עד ${toDate.toISOString().split('T')[0]}`);

    const ordersResult = await fetchBinaViaQuotaGuard(ORDERS_QUERY_URL, {
      tokenId: token.trim(),
      docType: -15,
      fromDate: fromDate.toISOString().split('T')[0],
      toDate: toDate.toISOString().split('T')[0],
    });

    if (!ordersResult.ok) {
      return new Response(JSON.stringify({
        success: false,
        error: `בינה החזירו HTTP ${ordersResult.status}`,
      }), { status: 502, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    let ordersData;
    try {
      ordersData = JSON.parse(ordersResult.text);
    } catch {
      return new Response(JSON.stringify({
        success: false,
        error: 'תגובה לא-JSON מבינה',
      }), { status: 502, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
    }

    const orders = Array.isArray(ordersData) ? ordersData : (ordersData?.Orders || []);
    console.log(`[update-names] התקבלו ${orders.length} הזמנות`);

    const nameMap = new Map();
    for (const order of orders) {
      const cid = String(order.custId);
      if (!nameMap.has(cid) && order.custName) {
        nameMap.set(cid, {
          name: String(order.custName).trim(),
          address: String(order.custAddress || '').trim(),
          city: String(order.custCity || '').trim(),
        });
      }
    }
    console.log(`[update-names] נבנתה מפת שמות עם ${nameMap.size} ערכים`);

    let updated = 0;
    let notFound = 0;
    const updates = [];

    for (const client of tempClientList) {
      const info = nameMap.get(String(client.bina_customer_id));
      if (!info) {
        notFound++;
        continue;
      }

      const newAddress = client.address || [info.address, info.city].filter(Boolean).join(', ');

      const { error: updErr } = await supabase
        .from('clients')
        .update({
          name: info.name,
          address: newAddress,
          notes: 'שם עודכן אוטומטית מהזמנה היסטורית בבינה',
        })
        .eq('id', client.id);

      if (!updErr) {
        updated++;
        updates.push({
          bina_customer_id: client.bina_customer_id,
          old_name: client.name,
          new_name: info.name,
        });
      } else {
        console.error(`[update-names] שגיאה בעדכון ${client.bina_customer_id}:`, updErr.message);
      }
    }

    for (const upd of updates) {
      await supabase
        .from('debt_snapshots')
        .update({ customer_name: upd.new_name })
        .eq('bina_customer_id', upd.bina_customer_id);
    }

    console.log(`[update-names] סיכום: ${updated} עודכנו, ${notFound} לא נמצאו`);

    return new Response(JSON.stringify({
      success: true,
      summary: {
        total_temp_clients: tempClientList.length,
        names_found_in_bina: nameMap.size,
        updated,
        not_found: notFound,
      },
      sample_updates: updates.slice(0, 10),
    }), { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });

  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error('[update-names] חריגה:', msg);
    return new Response(JSON.stringify({ success: false, error: msg }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });
  }
});

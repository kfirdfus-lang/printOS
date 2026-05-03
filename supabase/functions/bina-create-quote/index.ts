import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { fetchBinaViaQuotaGuard } from '../_shared/bina-proxy-fetch.ts'

// ⚠️ Endpoint שונה ליצירה: PostJsonDoc.aspx (לא V2!)
const BINA_CREATE_URL = 'https://webfiles.binaw.com/post/PostJsonDoc.aspx'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    const supabaseUrl = Deno.env.get('SUPABASE_URL')!
    const supabaseServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!
    const binaToken = Deno.env.get('BINA_TOKEN')!

    if (!binaToken) {
      throw new Error('BINA_TOKEN not configured')
    }

    const supabase = createClient(supabaseUrl, supabaseServiceKey)

    const { quote_id } = await req.json()
    
    if (!quote_id) {
      throw new Error('quote_id is required')
    }

    // שליפת ההצעה
    const { data: quote, error: quoteError } = await supabase
      .from('quotes')
      .select('*')
      .eq('id', quote_id)
      .single()

    if (quoteError || !quote) {
      throw new Error(`Quote not found: ${quoteError?.message}`)
    }

    // שליפת פריטים
    const { data: items, error: itemsError } = await supabase
      .from('quote_items')
      .select('*')
      .eq('quote_id', quote_id)
      .order('line_number')

    if (itemsError) {
      throw new Error(`Failed to fetch items: ${itemsError.message}`)
    }

    if (!items || items.length === 0) {
      throw new Error('No items in quote')
    }

    // יצירת requestId חד-ערכי (timestamp + random)
    const requestId = Math.floor(Date.now() / 1000) + Math.floor(Math.random() * 1000)

    // 🎯 בניית הבקשה לבינה לפי המבנה הנכון
    const binaRequest = {
      tokenId: binaToken,
      requestId: requestId,
      docType: 14,  // הצעת מחיר
      docWithVat: 1,
      docTitle: quote.title.substring(0, 100),
      
      // פרטי לקוח — תת-אובייקט
      Cust: {
        custId: quote.bina_cust_id,
        custName: (quote.bina_cust_name || '').substring(0, 50),
        custCity: (quote.bina_cust_city || quote.bina_cust_address || 'לא צוין').substring(0, 50),
        custAddress: (quote.bina_cust_address || 'לא צוין').substring(0, 100),
        custTel: (quote.bina_cust_phone || '').substring(0, 20),
      },
      
      // פריטים — שם השדה הוא docItems!
      docItems: items.map((item: any, idx: number) => ({
        ItemId: (item.item_name || `item-${idx + 1}`).substring(0, 30),
        ItemDesc: (item.description || item.item_name || '').substring(0, 200),
        ItemQty: Number(item.quantity),
        UnitPrice: Number(item.unit_price),
        UnitCurrency: 'ILS',
        CurValue: 1,
        Discount: Number(item.discount_pct || 0),
      })),
    }

    console.log('Sending to Bina:', JSON.stringify(binaRequest, null, 2))

    const binaResponse = await fetchBinaViaQuotaGuard(BINA_CREATE_URL, binaRequest)
    const responseText = binaResponse.text
    console.log('Bina response:', responseText)

    if (!binaResponse.ok) {
      throw new Error(`Bina HTTP error: ${binaResponse.status} - ${responseText}`)
    }

    let binaData
    try {
      binaData = JSON.parse(responseText)
    } catch {
      throw new Error(`Invalid JSON from Bina: ${responseText.substring(0, 200)}`)
    }

    // התגובה יכולה להיות אובייקט או מערך עם אובייקט
    const responseObj = Array.isArray(binaData) ? binaData[0] : binaData

    if (responseObj?.ResCode !== 0) {
      // שגיאה
      await supabase
        .from('quotes')
        .update({ 
          bina_error: responseObj?.ResMsg || 'Unknown error',
          status: 'שגיאה',
        })
        .eq('id', quote_id)

      return new Response(
        JSON.stringify({
          success: false,
          error: 'Bina rejected the quote',
          binaMessage: responseObj?.ResMsg,
          binaResponse: binaData,
          requestSent: binaRequest,
        }),
        { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      )
    }

    // הצלחה
    const binaDocId = responseObj.docId

    await supabase
      .from('quotes')
      .update({ 
        bina_doc_id: binaDocId,
        bina_synced_at: new Date().toISOString(),
        bina_error: null,
        status: 'נשלחה',
      })
      .eq('id', quote_id)

    return new Response(
      JSON.stringify({
        success: true,
        binaDocId: binaDocId,
        message: 'Quote created in Bina successfully',
      }),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )

  } catch (error) {
    console.error('Error:', error)
    return new Response(
      JSON.stringify({ error: error.message }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )
  }
})
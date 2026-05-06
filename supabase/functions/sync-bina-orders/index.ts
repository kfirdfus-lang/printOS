import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { fetchBinaViaQuotaGuard } from '../_shared/bina-proxy-fetch.ts'

const BINA_API_URL = 'https://webfiles.binaw.com/post/PostJsonDocV2.aspx'
const DEFAULT_DEPT = 'חדש'

const VALID_DEPTS = [
  'פורמט רחב',
  'דיגיטלי צבעוני',
  'דיגיטלי שחור לבן',
  'אופסט',
  'עבודות חוץ',
  'משלוחים',
  'תזכורות',
  'ביגוד ומוצרי פרסום'
]

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

function extractDept(title: string): { dept: string, cleanTitle: string } {
  if (!title) return { dept: DEFAULT_DEPT, cleanTitle: '' }

  const match = title.match(/^\s*\[([^\]]+)\]\s*(.*)/)
  if (match) {
    const extractedDept = match[1].trim()
    const cleanTitle = match[2].trim()

    if (VALID_DEPTS.includes(extractedDept)) {
      return { dept: extractedDept, cleanTitle: cleanTitle || title }
    }
    return { dept: DEFAULT_DEPT, cleanTitle: title }
  }

  return { dept: DEFAULT_DEPT, cleanTitle: title }
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
      throw new Error('BINA_TOKEN secret is not configured')
    }

    const supabase = createClient(supabaseUrl, supabaseServiceKey)

    // חלון יום אספקה אחד בלבד: היום (YYYY-MM-DD). הזמנות חדשות מקבלות תאריך אספקה של היום — מצמצם נפח תשובה מבינה.
    const todayStr = new Date().toISOString().split('T')[0]
    const fromDate = todayStr
    const toDate = todayStr

    const requestBody = {
      tokenId: binaToken,
      docType: -15,
      fromDate: fromDate,
      toDate: toDate,
    }

    console.log(
      '[sync-bina-orders] מסנכרנים יום אספקה אחד בלבד (fromDate=toDate=היום, פורמט YYYY-MM-DD):',
      fromDate,
      '| שליחה לבינה:',
      JSON.stringify(requestBody),
    )

    const binaResponse = await fetchBinaViaQuotaGuard(BINA_API_URL, requestBody)
    const responseText = binaResponse.text
    console.log('Bina raw response:', responseText)

    if (!binaResponse.ok) {
      throw new Error(`Bina API HTTP error: ${binaResponse.status} - ${responseText}`)
    }

    let binaData
    try {
      binaData = JSON.parse(responseText)
    } catch {
      throw new Error(`Bina returned invalid JSON: ${responseText}`)
    }

    if (Array.isArray(binaData) && binaData[0]?.ResCode !== undefined && binaData[0]?.ResCode !== 0) {
      return new Response(
        JSON.stringify({
          error: 'Bina API returned error',
          binaResponse: binaData,
          requestSent: requestBody,
        }),
        { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      )
    }

    let orders = []
    if (binaData.Orders && Array.isArray(binaData.Orders)) {
      orders = binaData.Orders
    } else if (Array.isArray(binaData)) {
      orders = binaData
    }

    let created = 0
    let skipped = 0
    let errors = 0
    const errorDetails: string[] = []

    for (const order of orders) {
      try {
        const binaOrderId = order.orderId

        if (!binaOrderId) {
          errors++
          errorDetails.push('Order without orderId')
          continue
        }

        // כולל משימות בארכיון — כדי שלא תיווצר כפילות אחרי ארכוב (במקום מחיקה).
        const { data: existing } = await supabase
          .from('tasks')
          .select('id')
          .eq('bina_order_id', binaOrderId)
          .maybeSingle()

        if (existing) {
          skipped++
          continue
        }

        const orderTitle = order.orderTitle || ''
        const { dept, cleanTitle } = extractDept(orderTitle)

        let itemsDescription = ''
        if (order.Order?.items && Array.isArray(order.Order.items)) {
          itemsDescription = order.Order.items
            .map((item: any) => `${item.itemDesc} (${item.itemQty})`)
            .join(', ')
        }

        const taskTitle = cleanTitle
          || itemsDescription
          || `הזמנה #${binaOrderId}`

        let dueDate = null
        if (order.orderDeliveryDate) {
          const parts = order.orderDeliveryDate.split('/')
          if (parts.length === 3) {
            dueDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`
          }
        }

        const taskData: any = {
          title: taskTitle,
          dept: dept,
          status: 'חדש',
          priority: 'רגיל',
          client_name: order.custName || '',
          contact: order.orderTo || '',
          due_date: dueDate,
          notes: itemsDescription
            ? `פירוט: ${itemsDescription}\nסוכן: ${order.orderSalesMan || ''}\nכתובת: ${order.custAddress || ''}, ${order.custCity || ''}`
            : '',
          bina_order_id: binaOrderId,
          bina_cust_id: order.custId || null,
          bina_cust_address: order.custAddress || null,
          bina_cust_city: order.custCity || null,
          bina_synced_at: new Date().toISOString(),
          source: 'bina',
          created_by: 'sync-bina',
        }

        const { error: insertError } = await supabase
          .from('tasks')
          .insert(taskData)

        if (insertError) {
          errors++
          errorDetails.push(`Order ${binaOrderId}: ${insertError.message}`)
        } else {
          created++
        }
      } catch (err) {
        errors++
        errorDetails.push(`Order processing error: ${err.message}`)
      }
    }

    return new Response(
      JSON.stringify({
        success: true,
        totalFromBina: orders.length,
        created,
        skipped,
        errors,
        errorDetails: errors > 0 ? errorDetails : undefined,
        dateRange: { from: fromDate, to: toDate },
      }),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )

  } catch (error) {
    console.error('Sync error:', error)
    return new Response(
      JSON.stringify({ error: error.message }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )
  }
})
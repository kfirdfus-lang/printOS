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

interface ClientRecord {
  id: string
  name: string
  phone: string | null
  address: string | null
  bina_customer_id: string | null
}

function parseHebrewDateForDB(dateStr: unknown): string | null {
  if (!dateStr || typeof dateStr !== 'string') return null
  const parts = dateStr.trim().split('/')
  if (parts.length !== 3) return null
  const [dd, mm, yyyy] = parts
  if (!dd || !mm || !yyyy) return null
  return `${yyyy}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`
}

function num(v: unknown): number | null {
  if (v === undefined || v === null || v === '') return null
  const n = Number(v)
  return Number.isFinite(n) ? n : null
}

function normalizeName(name: string): string {
  if (!name) return ''
  return name.trim().replace(/\s+/g, ' ').toLowerCase()
}

async function loadClientMaps(supabase: ReturnType<typeof createClient>): Promise<{
  linkedByBinaId: Map<string, ClientRecord>
  unlinkedByName: Map<string, ClientRecord>
}> {
  const linkedByBinaId = new Map<string, ClientRecord>()
  const unlinkedByName = new Map<string, ClientRecord>()
  const { data, error } = await supabase
    .from('clients')
    .select('id, name, phone, address, bina_customer_id')

  if (error) {
    console.error('[sync-bina-orders] failed to preload clients:', error.message)
    return { linkedByBinaId, unlinkedByName }
  }

  for (const c of data || []) {
    const row = c as ClientRecord
    if (row.bina_customer_id) {
      linkedByBinaId.set(String(row.bina_customer_id), row)
    } else {
      const nn = normalizeName(row.name || '')
      if (nn) unlinkedByName.set(nn, row)
    }
  }
  return { linkedByBinaId, unlinkedByName }
}

/** Upsert לקוח בטבלה — כשלים רק נרשמים ללוג; לא זורק. */
async function ensureClientFromOrder(
  order: Record<string, unknown>,
  supabase: ReturnType<typeof createClient>,
  linkedByBinaId: Map<string, ClientRecord>,
  unlinkedByName: Map<string, ClientRecord>,
): Promise<void> {
  const custIdRaw = order.custId
  if (custIdRaw === undefined || custIdRaw === null || String(custIdRaw).trim() === '') return

  const binaIdStr = String(custIdRaw)
  if (linkedByBinaId.has(binaIdStr)) return

  const addressParts: string[] = []
  if (order.custAddress) addressParts.push(String(order.custAddress))
  if (order.custCity) addressParts.push(String(order.custCity))
  const fullAddress = addressParts.length > 0 ? addressParts.join(', ') : null

  const custNameTrim = String(order.custName || '').trim()
  const nameKey = normalizeName(custNameTrim)
  const existingUnlinked = nameKey ? unlinkedByName.get(nameKey) : undefined

  try {
    if (existingUnlinked) {
      const updates: Record<string, unknown> = { bina_customer_id: binaIdStr }
      if (!existingUnlinked.address && fullAddress) updates.address = fullAddress

      const { error: updErr } = await supabase
        .from('clients')
        .update(updates)
        .eq('id', existingUnlinked.id)

      if (updErr) {
        console.error(`[sync-bina-orders] client update failed (${custNameTrim || binaIdStr}):`, updErr.message)
        return
      }

      unlinkedByName.delete(nameKey)
      linkedByBinaId.set(binaIdStr, {
        ...existingUnlinked,
        bina_customer_id: binaIdStr,
        address: fullAddress ?? existingUnlinked.address,
      })
      return
    }

    const { data: inserted, error: insertErr } = await supabase
      .from('clients')
      .insert({
        name: custNameTrim || `לקוח ${binaIdStr}`,
        address: fullAddress,
        bina_customer_id: binaIdStr,
        notes: 'נוצר אוטומטית מסנכרון בינה',
      })
      .select('id, name, phone, address, bina_customer_id')
      .single()

    if (insertErr) {
      console.error(`[sync-bina-orders] client insert failed (bina ${binaIdStr}):`, insertErr.message)
      return
    }

    linkedByBinaId.set(binaIdStr, inserted as ClientRecord)
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e)
    console.error('[sync-bina-orders] ensureClientFromOrder:', msg)
  }
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

    const { linkedByBinaId, unlinkedByName } = await loadClientMaps(supabase)

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

        const o = order as Record<string, unknown>
        const taskData: Record<string, unknown> = {
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
          total_amount: num(o.orderTotalAfterDiscount) ?? num(o.orderTotal),
          total_inc_vat: num(o.orderTotalIncVat),
          discount_amount: num(o.orderDiscount),
          sales_agent: o.orderSalesMan ? String(o.orderSalesMan) : null,
          bina_order_date: parseHebrewDateForDB(o.orderDate),
        }

        await ensureClientFromOrder(order, supabase, linkedByBinaId, unlinkedByName)

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
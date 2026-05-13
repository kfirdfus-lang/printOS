import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { fetchBinaViaQuotaGuard } from '../_shared/bina-proxy-fetch.ts'

const BINA_API_URL = 'https://webfiles.binaw.com/post/PostJsonDocV2.aspx'
const DEFAULT_DEPT = 'חדש'

/** קוד ב-itemId בבינה (1–7) → שם מחלקה ב-PrintOS */
const DEPARTMENT_CODES: Record<string, string> = {
  '1': 'פורמט רחב',
  '2': 'ביגוד ומוצרי פרסום',
  '3': 'דיגיטלי צבעוני',
  '4': 'דיגיטלי שחור לבן',
  '5': 'אופסט',
  '6': 'עבודות חוץ',
  '7': 'מתקני תצוגה ומוצרים נלווים',
}

const VALID_DEPTS = [
  'פורמט רחב',
  'דיגיטלי צבעוני',
  'דיגיטלי שחור לבן',
  'אופסט',
  'עבודות חוץ',
  'משלוחים',
  'תזכורות',
  'ביגוד ומוצרי פרסום',
  'מתקני תצוגה ומוצרים נלווים',
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

async function upsertTaskItemsFromBinaOrder(
  supabase: ReturnType<typeof createClient>,
  taskId: string,
  binaOrderId: number,
  order: Record<string, unknown>,
): Promise<boolean> {
  const orderNested = order.Order as { items?: unknown[] } | undefined
  if (!orderNested?.items || !Array.isArray(orderNested.items)) return false

  const rows: Record<string, unknown>[] = []
  let fallbackLine = 0
  for (const raw of orderNested.items) {
    fallbackLine += 1
    const item = raw as Record<string, unknown>
    const itemCode = String(item.itemId ?? '').trim()
    if (!itemCode || !DEPARTMENT_CODES[itemCode]) continue

    const ln = Number(item.itemLineNumber)
    const line_number = Number.isFinite(ln) ? ln : fallbackLine

    rows.push({
      task_id: taskId,
      bina_order_id: binaOrderId,
      line_number,
      bina_item_code: itemCode,
      department: DEPARTMENT_CODES[itemCode],
      description: (String(item.itemDesc ?? '').trim() || '—'),
      quantity: Number(item.itemQty) || 0,
      price: Number(item.itemPrice) || 0,
      total: Number(item.itemTotal) || 0,
      status: 'בעבודה',
    })
  }

  if (rows.length === 0) return false

  const { error: itemsErr } = await supabase.from('task_items').upsert(rows, {
    onConflict: 'bina_order_id,line_number',
  })

  if (itemsErr) {
    console.error('[sync-bina-orders] task_items upsert:', itemsErr.message)
    return false
  }

  const { error: flagErr } = await supabase.from('tasks').update({ has_items: true }).eq('id', taskId)
  if (flagErr) {
    console.error('[sync-bina-orders] has_items update:', flagErr.message)
  }
  return true
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

    // טווח תאריכי אספקה: 7 ימים אחורה עד 30 ימים קדימה (YYYY-MM-DD) — כדי לתפוס הזמנות שנכנסו בערב למועד מחר וכו׳.
    const fromDate = new Date()
    fromDate.setDate(fromDate.getDate() - 7)
    const toDate = new Date()
    toDate.setDate(toDate.getDate() + 30)

    const fromStr = fromDate.toISOString().split('T')[0]
    const toStr = toDate.toISOString().split('T')[0]

    const requestBody = {
      tokenId: binaToken,
      docType: -15,
      fromDate: fromStr,
      toDate: toStr,
    }

    console.log(
      '[sync-bina-orders] מסנכרנים טווח אספקה (from=-7d, to=+30d, YYYY-MM-DD):',
      fromStr,
      '→',
      toStr,
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

    for (let orderIdx = 0; orderIdx < orders.length; orderIdx++) {
      const order = orders[orderIdx]
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

        const { data: insertedTask, error: insertError } = await supabase
          .from('tasks')
          .insert(taskData)
          .select('id')
          .single()

        if (insertError) {
          errors++
          errorDetails.push(`Order ${binaOrderId}: ${insertError.message}`)
        } else {
          created++
          const taskId = insertedTask?.id as string | undefined
          if (taskId) {
            await upsertTaskItemsFromBinaOrder(supabase, taskId, Number(binaOrderId), order as Record<string, unknown>)
          }
        }
      } catch (err) {
        errors++
        const msg = err instanceof Error ? err.message : String(err)
        errorDetails.push(`Order processing error: ${msg}`)
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
        dateRange: { from: fromStr, to: toStr },
      }),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )

  } catch (error) {
    console.error('Sync error:', error)
    const msg = error instanceof Error ? error.message : String(error)
    return new Response(
      JSON.stringify({ error: msg }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )
  }
})
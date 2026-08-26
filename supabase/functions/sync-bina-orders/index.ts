import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { fetchBinaViaQuotaGuard } from '../_shared/bina-proxy-fetch.ts'
import { rejectDisallowedInternalOrigin } from '../_shared/cors.ts'

const BINA_API_URL = 'https://webfiles.binaw.com/post/PostJsonDocV2.aspx'
const DEFAULT_DEPT = 'חדש'

/** קוד ב-itemId בבינה (2–8; 1=כללי בבינה — לא ממופה) → שם מחלקה ב-PrintOS */
const DEPARTMENT_CODES: Record<string, string> = {
  '2': 'ביגוד ומוצרי פרסום',
  '3': 'דיגיטלי צבעוני',
  '4': 'דיגיטלי שחור לבן',
  '5': 'אופסט',
  '6': 'עבודות חוץ',
  '7': 'מתקני תצוגה ומוצרים נלווים',
  '8': 'פורמט רחב',
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

/** orderStatus / orderState מבינה — שמירה כ-text (ריק → null) */
function binaOrderStatusFields(order: Record<string, unknown>): {
  bina_order_status: string | null
  bina_order_state: string | null
} {
  const statusRaw = order.orderStatus
  const stateRaw = order.orderState
  const status = statusRaw != null && String(statusRaw).trim() !== '' ? String(statusRaw).trim() : null
  const state = stateRaw != null && String(stateRaw).trim() !== '' ? String(stateRaw).trim() : null
  return { bina_order_status: status, bina_order_state: state }
}

/** invNumber מבינה → bina_invoice_number (null אם חסר / לא מספרי) */
function binaInvoiceNumber(order: Record<string, unknown>): number | null {
  return num(order.invNumber)
}

function normalizeBinaSalesAgentServer(val: unknown): string | null {
  const v = String(val ?? '').trim()
  if (!v) return null
  if (v === 'כפיר' || /^כפיר\b/.test(v)) return 'כפיר צמח'
  if (/^נטלי\b/.test(v)) return 'נטלי'
  if (/^ברק\b/.test(v)) return 'ברק'
  return v
}

function deptFromOrderItems(order: Record<string, unknown>): string | null {
  const nested = order.Order as { items?: unknown[] } | undefined
  if (!nested?.items || !Array.isArray(nested.items)) return null
  for (const raw of nested.items) {
    const item = raw as Record<string, unknown>
    const code = String(item.itemId ?? '').trim()
    if (DEPARTMENT_CODES[code]) return DEPARTMENT_CODES[code]
  }
  return null
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

/** סנכרון task_items ממבנה בינה: מחיקת שורות שלא קיימות יותר, upsert, ועדכון has_items — לא נוגע בשדות tasks אחרים. */
async function syncTaskItemsForBinaOrder(
  supabase: ReturnType<typeof createClient>,
  taskId: string,
  binaOrderId: number,
  order: Record<string, unknown>,
): Promise<boolean> {
  const orderNested = order.Order as { items?: unknown[] } | undefined
  const rawItems = orderNested?.items && Array.isArray(orderNested.items) ? orderNested.items : []

  /** בינה לא החזירה מערך פריטים — לא מוחקים ולא מעדכנים (לא לדרוס מצב קיים). */
  if (rawItems.length === 0) return true

  const { data: existingRows, error: existingErr } = await supabase
    .from('task_items')
    .select('id, line_number, status')
    .eq('bina_order_id', binaOrderId)

  if (existingErr) {
    console.error('[sync-bina-orders] task_items preload:', existingErr.message)
    return false
  }

  const statusByLine = new Map<number, string>()
  for (const r of existingRows || []) {
    const row = r as { line_number: number; status: string }
    statusByLine.set(Number(row.line_number), String(row.status || ''))
  }

  const lineNumbersFromBina: number[] = []
  let fb = 0
  for (const raw of rawItems) {
    fb += 1
    const item = raw as Record<string, unknown>
    const ln = Number(item.itemLineNumber)
    lineNumbersFromBina.push(Number.isFinite(ln) ? ln : fb)
  }
  const uniqueSet = new Set(lineNumbersFromBina)

  if (uniqueSet.size > 0) {
    const idsToRemove = (existingRows || [])
      .filter((row) => {
        const r = row as { id: string; line_number: number }
        return !uniqueSet.has(Number(r.line_number))
      })
      .map((row) => (row as { id: string }).id)

    if (idsToRemove.length > 0) {
      const { error: delErr } = await supabase.from('task_items').delete().in('id', idsToRemove)
      if (delErr) {
        console.error('[sync-bina-orders] task_items delete stale lines:', delErr.message)
        return false
      }
    }
  }

  const rows: Record<string, unknown>[] = []
  let fallbackLine = 0
  for (const raw of rawItems) {
    fallbackLine += 1
    const item = raw as Record<string, unknown>
    const itemCode = String(item.itemId ?? '').trim()
    if (!itemCode || !DEPARTMENT_CODES[itemCode]) continue

    const ln = Number(item.itemLineNumber)
    const line_number = Number.isFinite(ln) ? ln : fallbackLine
    const prev = statusByLine.get(line_number)
    const status = prev === 'מוכן' ? 'מוכן' : 'בעבודה'

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
      status,
    })
  }

  if (rows.length > 0) {
    const { error: itemsErr } = await supabase.from('task_items').upsert(rows, {
      onConflict: 'bina_order_id,line_number',
    })

    if (itemsErr) {
      console.error('[sync-bina-orders] task_items upsert:', itemsErr.message)
      return false
    }
  }

  const { count, error: cntErr } = await supabase
    .from('task_items')
    .select('id', { count: 'exact', head: true })
    .eq('task_id', taskId)

  if (cntErr) {
    console.error('[sync-bina-orders] task_items count:', cntErr.message)
    return false
  }

  const hasItems = (count ?? 0) > 0
  const { error: flagErr } = await supabase.from('tasks').update({ has_items: hasItems }).eq('id', taskId)
  if (flagErr) {
    console.error('[sync-bina-orders] has_items update:', flagErr.message)
    return false
  }

  return true
}

Deno.serve(async (req) => {
  const originBlock = rejectDisallowedInternalOrigin(req)
  if (originBlock) return originBlock

  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    const url = new URL(req.url)
    const debugFirstOrderInResponse = url.searchParams.get('debug') === '1'

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

    const binaOrders = orders

    let created = 0
    let skipped = 0
    let taskItemsSynced = 0
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

        const orderHasItemsPayload = Boolean(
          order.Order &&
            Array.isArray((order.Order as { items?: unknown }).items) &&
            (order.Order as { items: unknown[] }).items.length > 0,
        )

        // כולל משימות בארכיון — כדי שלא תיווצר כפילות אחרי ארכוב (במקום מחיקה).
        const { data: existing } = await supabase
          .from('tasks')
          .select('id, dept, has_items, total_amount, total_inc_vat, discount_amount, sales_agent, source, due_date, bina_order_date')
          .eq('bina_order_id', binaOrderId)
          .maybeSingle()

        if (existing?.id) {
          skipped++
          await ensureClientFromOrder(order, supabase, linkedByBinaId, unlinkedByName)

          const o = order as Record<string, unknown>
          const invNum = binaInvoiceNumber(o)
          await supabase
            .from('tasks')
            .update({
              ...binaOrderStatusFields(o),
              ...(invNum != null ? { bina_invoice_number: invNum } : {}),
            })
            .eq('id', existing.id)

          const deptFromItemsName = deptFromOrderItems(o)

          let dueDate = null
          if (order.orderDeliveryDate) {
            const parts = order.orderDeliveryDate.split('/')
            if (parts.length === 3) {
              dueDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`
            }
          }

          const totalAmount = num(o.orderTotalAfterDiscount) ?? num(o.orderTotal)
          const totalIncVat = num(o.orderTotalIncVat)
          const discountAmount = num(o.orderDiscount)
          const normalizedSalesAgent = normalizeBinaSalesAgentServer(o.orderSalesMan)

          // לא לדרוס task_items שיש להם כבר department.
          const { data: existingItems } = await supabase
            .from('task_items')
            .select('id, department')
            .eq('task_id', existing.id)

          const hasDeptItems =
            (existingItems || []).some((it) => String((it as any).department || '').trim() !== '')

          if (!hasDeptItems) {
            const ok = await syncTaskItemsForBinaOrder(
              supabase,
              existing.id as string,
              Number(binaOrderId),
              order as Record<string, unknown>,
            )
            if (ok && orderHasItemsPayload) taskItemsSynced++
          } else {
            // אם has_items לא מסומן, נרשום true בלבד (בלי לשנות פריטים).
            if (!existing.has_items) {
              await supabase.from('tasks').update({ has_items: true }).eq('id', existing.id)
            }
          }

          // להשלים שדות ריקים בלבד.
          const patch: Record<string, unknown> = {}
          const existingDept = existing.dept != null ? String(existing.dept).trim() : ''
          if (!existingDept && deptFromItemsName) patch.dept = deptFromItemsName
          if (existing.total_amount == null && totalAmount != null) patch.total_amount = totalAmount
          if (existing.total_inc_vat == null && totalIncVat != null) patch.total_inc_vat = totalIncVat
          if (existing.discount_amount == null && discountAmount != null) patch.discount_amount = discountAmount
          if (!existing.source || String(existing.source).trim() === '') patch.source = 'bina'
          if ((!existing.sales_agent || String(existing.sales_agent).trim() === '') && normalizedSalesAgent)
            patch.sales_agent = normalizedSalesAgent
          const parsedBinaOrderDate = o.orderDate ? parseHebrewDateForDB(o.orderDate) : null
          if ((!existing.bina_order_date || String(existing.bina_order_date).trim() === '') && parsedBinaOrderDate)
            patch.bina_order_date = parsedBinaOrderDate
          if ((!existing.due_date || String(existing.due_date).trim() === '') && dueDate) patch.due_date = dueDate

          if (Object.keys(patch).length > 0) {
            await supabase.from('tasks').update(patch).eq('id', existing.id)
          }

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
          sales_agent: normalizeBinaSalesAgentServer(o.orderSalesMan),
          bina_order_date: parseHebrewDateForDB(o.orderDate),
          bina_invoice_number: binaInvoiceNumber(o),
          ...binaOrderStatusFields(o),
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
            const ok = await syncTaskItemsForBinaOrder(
              supabase,
              taskId,
              Number(binaOrderId),
              order as Record<string, unknown>,
            )
            if (ok && orderHasItemsPayload) taskItemsSynced++
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
        taskItemsSynced,
        errors,
        errorDetails: errors > 0 ? errorDetails : undefined,
        dateRange: { from: fromStr, to: toStr },
        ...(debugFirstOrderInResponse && binaOrders.length > 0
          ? { _debug_first_bina_order: binaOrders[0] }
          : {}),
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
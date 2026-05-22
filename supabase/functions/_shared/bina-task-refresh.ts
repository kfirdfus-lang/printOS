import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { fetchBinaViaQuotaGuard } from './bina-proxy-fetch.ts'

const BINA_API_URL = 'https://webfiles.binaw.com/post/PostJsonDocV2.aspx'
const DEFAULT_DEPT = 'חדש'

export const DEPARTMENT_CODES: Record<string, string> = {
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

export type BinaChange =
  | { field: 'title'; old: string; new: string }
  | { field: 'total_price'; old: number | null; new: number | null }
  | { field: 'item_added'; description: string; quantity: number }
  | { field: 'item_removed'; description: string; quantity: number }
  | { field: 'quantity_changed'; description: string; old: number; new: number }

export function num(v: unknown): number | null {
  if (v === undefined || v === null || v === '') return null
  const n = Number(v)
  return Number.isFinite(n) ? n : null
}

function extractDept(title: string): { dept: string; cleanTitle: string } {
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

export function titleFromBinaOrder(order: Record<string, unknown>): string {
  const binaOrderId = order.orderId
  const orderTitle = String(order.orderTitle || '')
  const { cleanTitle } = extractDept(orderTitle)
  const nested = order.Order as { items?: unknown[] } | undefined
  let itemsDescription = ''
  if (nested?.items && Array.isArray(nested.items)) {
    itemsDescription = nested.items
      .map((item: Record<string, unknown>) => `${item.itemDesc} (${item.itemQty})`)
      .join(', ')
  }
  return cleanTitle || itemsDescription || `הזמנה #${binaOrderId}`
}

export function compareItemsFromBina(
  oldItems: { description: string; quantity: number }[],
  order: Record<string, unknown>,
): { changed: boolean; changes: BinaChange[] } {
  const nested = order.Order as { items?: unknown[] } | undefined
  const rawItems = nested?.items && Array.isArray(nested.items) ? nested.items : []
  const newItems: { description: string; quantity: number }[] = []
  for (const raw of rawItems) {
    const item = raw as Record<string, unknown>
    const itemCode = String(item.itemId ?? '').trim()
    if (!itemCode || !DEPARTMENT_CODES[itemCode]) continue
    newItems.push({
      description: (String(item.itemDesc ?? '').trim() || '—'),
      quantity: Number(item.itemQty) || 0,
    })
  }

  const changes: BinaChange[] = []
  const oldByDesc: Record<string, { description: string; quantity: number }> = {}
  const newByDesc: Record<string, { description: string; quantity: number }> = {}
  oldItems.forEach((i) => { oldByDesc[i.description] = i })
  newItems.forEach((i) => { newByDesc[i.description] = i })

  newItems.forEach((newItem) => {
    if (!oldByDesc[newItem.description]) {
      changes.push({
        field: 'item_added',
        description: newItem.description,
        quantity: newItem.quantity,
      })
    } else if (oldByDesc[newItem.description].quantity !== newItem.quantity) {
      changes.push({
        field: 'quantity_changed',
        description: newItem.description,
        old: oldByDesc[newItem.description].quantity,
        new: newItem.quantity,
      })
    }
  })

  oldItems.forEach((oldItem) => {
    if (!newByDesc[oldItem.description]) {
      changes.push({
        field: 'item_removed',
        description: oldItem.description,
        quantity: oldItem.quantity,
      })
    }
  })

  const changed = changes.length > 0
  return { changed, changes }
}

export async function fetchBinaOrders(fromDate: string, toDate: string): Promise<Record<string, unknown>[]> {
  const binaToken = Deno.env.get('BINA_TOKEN')
  if (!binaToken) throw new Error('BINA_TOKEN secret is not configured')

  const requestBody = {
    tokenId: binaToken,
    docType: -15,
    fromDate,
    toDate,
  }

  const binaResponse = await fetchBinaViaQuotaGuard(BINA_API_URL, requestBody)
  if (!binaResponse.ok) {
    throw new Error(`Bina API HTTP error: ${binaResponse.status} - ${binaResponse.text}`)
  }

  let binaData: unknown
  try {
    binaData = JSON.parse(binaResponse.text)
  } catch {
    throw new Error(`Bina returned invalid JSON: ${binaResponse.text}`)
  }

  if (Array.isArray(binaData) && (binaData as { ResCode?: number }[])[0]?.ResCode !== undefined &&
    (binaData as { ResCode?: number }[])[0]?.ResCode !== 0) {
    throw new Error('Bina API returned error')
  }

  const bd = binaData as { Orders?: unknown[] }
  if (bd.Orders && Array.isArray(bd.Orders)) return bd.Orders as Record<string, unknown>[]
  if (Array.isArray(binaData)) return binaData as Record<string, unknown>[]
  return []
}

export function findBinaOrder(
  orders: Record<string, unknown>[],
  binaOrderId: string | number,
): Record<string, unknown> | null {
  const idStr = String(binaOrderId)
  return orders.find((o) => String(o.orderId) === idStr) ?? null
}

export function dateRangeForTask(task: Record<string, unknown>): { fromDate: string; toDate: string } {
  const anchor = task.bina_order_date
    ? new Date(String(task.bina_order_date))
    : task.created_at
      ? new Date(String(task.created_at))
      : new Date()

  const from = new Date(anchor)
  from.setDate(from.getDate() - 90)
  const to = new Date(anchor)
  to.setDate(to.getDate() + 60)

  return {
    fromDate: from.toISOString().split('T')[0],
    toDate: to.toISOString().split('T')[0],
  }
}

/** סנכרון task_items מבינה — לא נוגע בשדות tasks מלבד has_items (נגזר מפריטים). */
export async function syncTaskItemsForBinaOrder(
  supabase: ReturnType<typeof createClient>,
  taskId: string,
  binaOrderId: number,
  order: Record<string, unknown>,
): Promise<boolean> {
  const orderNested = order.Order as { items?: unknown[] } | undefined
  const rawItems = orderNested?.items && Array.isArray(orderNested.items) ? orderNested.items : []

  if (rawItems.length === 0) return true

  const { data: existingRows, error: existingErr } = await supabase
    .from('task_items')
    .select('id, line_number, status')
    .eq('bina_order_id', binaOrderId)

  if (existingErr) {
    console.error('[bina-task-refresh] task_items preload:', existingErr.message)
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
        console.error('[bina-task-refresh] task_items delete:', delErr.message)
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
      console.error('[bina-task-refresh] task_items upsert:', itemsErr.message)
      return false
    }
  }

  const { count, error: cntErr } = await supabase
    .from('task_items')
    .select('id', { count: 'exact', head: true })
    .eq('task_id', taskId)

  if (cntErr) {
    console.error('[bina-task-refresh] task_items count:', cntErr.message)
    return false
  }

  const hasItems = (count ?? 0) > 0
  const { error: flagErr } = await supabase.from('tasks').update({ has_items: hasItems }).eq('id', taskId)
  if (flagErr) {
    console.error('[bina-task-refresh] has_items:', flagErr.message)
    return false
  }

  return true
}

export async function refreshOneTaskFromBina(
  supabase: ReturnType<typeof createClient>,
  taskId: string,
): Promise<{
  success: boolean
  changes?: BinaChange[]
  message?: string
  error?: string
}> {
  const { data: task, error: taskErr } = await supabase
    .from('tasks')
    .select('id, title, total_amount, bina_order_id, bina_order_date, created_at')
    .eq('id', taskId)
    .maybeSingle()

  if (taskErr || !task) {
    return { success: false, error: taskErr?.message || 'משימה לא נמצאה' }
  }

  if (!task.bina_order_id) {
    return { success: false, error: 'אין מספר הזמנת בינה' }
  }

  const { fromDate, toDate } = dateRangeForTask(task as Record<string, unknown>)
  const orders = await fetchBinaOrders(fromDate, toDate)
  const order = findBinaOrder(orders, task.bina_order_id as string | number)

  if (!order) {
    return { success: false, error: `הזמנה #${task.bina_order_id} לא נמצאה בבינה (${fromDate}–${toDate})` }
  }

  const { data: existingItems } = await supabase
    .from('task_items')
    .select('description, quantity')
    .eq('task_id', taskId)
    .neq('status', 'ארכיון')

  const oldItems = (existingItems || []).map((r) => ({
    description: String((r as { description: string }).description || '—'),
    quantity: Number((r as { quantity: number }).quantity) || 0,
  }))

  const changes: BinaChange[] = []
  const updates: Record<string, unknown> = {}

  const newTitle = titleFromBinaOrder(order)
  if (newTitle && newTitle !== task.title) {
    changes.push({ field: 'title', old: String(task.title || ''), new: newTitle })
    updates.title = newTitle
  }

  const newTotal = num(order.orderTotalAfterDiscount) ?? num(order.orderTotal)
  const oldTotal = task.total_amount != null ? Number(task.total_amount) : null
  if (newTotal != null && newTotal !== oldTotal) {
    changes.push({ field: 'total_price', old: oldTotal, new: newTotal })
    updates.total_amount = newTotal
  }

  const itemDiff = compareItemsFromBina(oldItems, order)
  if (itemDiff.changed) {
    changes.push(...itemDiff.changes)
  }

  if (changes.length === 0) {
    return { success: true, changes: [], message: 'אין שינויים' }
  }

  updates.last_bina_sync = new Date().toISOString()

  const { error: updErr } = await supabase.from('tasks').update(updates).eq('id', taskId)
  if (updErr) {
    return { success: false, error: updErr.message }
  }

  if (itemDiff.changed) {
    const ok = await syncTaskItemsForBinaOrder(
      supabase,
      taskId,
      Number(task.bina_order_id),
      order,
    )
    if (!ok) {
      return { success: false, error: 'עדכון פריטים נכשל' }
    }
  }

  return { success: true, changes }
}

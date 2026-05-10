import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { fetchBinaViaQuotaGuard } from '../_shared/bina-proxy-fetch.ts'

const BINA_API_URL = 'https://webfiles.binaw.com/post/PostJsonDocV2.aspx'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
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

    const daysBack = 365
    const toDate = new Date().toISOString().split('T')[0]
    const from = new Date()
    from.setDate(from.getDate() - daysBack)
    const fromDate = from.toISOString().split('T')[0]

    const requestBody = {
      tokenId: binaToken,
      docType: -15,
      fromDate,
      toDate,
    }

    console.log('[backfill-task-prices] fetching Bina orders', fromDate, '→', toDate)

    const binaResponse = await fetchBinaViaQuotaGuard(BINA_API_URL, requestBody)
    const responseText = binaResponse.text

    if (!binaResponse.ok) {
      throw new Error(`Bina API HTTP error: ${binaResponse.status} - ${responseText}`)
    }

    let binaData: unknown
    try {
      binaData = JSON.parse(responseText)
    } catch {
      throw new Error(`Bina returned invalid JSON: ${responseText}`)
    }

    let orders: Record<string, unknown>[] = []
    const bd = binaData as { Orders?: unknown[] }
    if (bd.Orders && Array.isArray(bd.Orders)) {
      orders = bd.Orders as Record<string, unknown>[]
    } else if (Array.isArray(binaData)) {
      orders = binaData as Record<string, unknown>[]
    }

    const byOrderId = new Map<string, Record<string, unknown>>()
    for (const order of orders) {
      const id = order.orderId as number | string | undefined
      if (id !== undefined && id !== null) {
        byOrderId.set(String(id), order)
      }
    }

    const { data: tasksNeeding, error: selErr } = await supabase
      .from('tasks')
      .select('id, bina_order_id')
      .not('bina_order_id', 'is', null)
      .is('total_amount', null)

    if (selErr) {
      throw new Error(selErr.message)
    }

    let updated = 0
    let missingFromBina = 0
    const errors: string[] = []

    for (const row of tasksNeeding || []) {
      const taskId = row.id as string
      const binaOrderId = row.bina_order_id as number | string
      const order = byOrderId.get(String(binaOrderId))
      if (!order) {
        missingFromBina++
        continue
      }

      const patch = {
        total_amount: num(order.orderTotalAfterDiscount) ?? num(order.orderTotal),
        total_inc_vat: num(order.orderTotalIncVat),
        discount_amount: num(order.orderDiscount),
        sales_agent: order.orderSalesMan ? String(order.orderSalesMan) : null,
        bina_order_date: parseHebrewDateForDB(order.orderDate),
      }

      const { error: updErr } = await supabase.from('tasks').update(patch).eq('id', taskId)

      if (updErr) {
        errors.push(`${taskId}: ${updErr.message}`)
      } else {
        updated++
      }
    }

    return new Response(
      JSON.stringify({
        success: true,
        fromDate,
        toDate,
        ordersFromBina: orders.length,
        tasksChecked: (tasksNeeding || []).length,
        updated,
        missingFromBina,
        errors: errors.length ? errors : undefined,
      }),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } },
    )
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error)
    console.error('[backfill-task-prices]', msg)
    return new Response(JSON.stringify({ error: msg }), {
      status: 500,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    })
  }
})

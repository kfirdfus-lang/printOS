import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { refreshOneTaskFromBina } from '../_shared/bina-task-refresh.ts'

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
    const supabase = createClient(supabaseUrl, supabaseServiceKey)

    const { data: tasks, error } = await supabase
      .from('tasks')
      .select('id')
      .not('bina_order_id', 'is', null)
      .is('completed_at', null)
      .is('archived_at', null)

    if (error) throw error

    let updated = 0
    let unchanged = 0
    let failed = 0

    for (const row of tasks || []) {
      const taskId = row.id as string
      try {
        const r = await refreshOneTaskFromBina(supabase, taskId)
        if (!r.success) {
          failed++
          continue
        }
        if (r.changes && r.changes.length > 0) updated++
        else unchanged++
      } catch {
        failed++
      }
    }

    return new Response(
      JSON.stringify({
        success: true,
        total: (tasks || []).length,
        updated,
        unchanged,
        failed,
        timestamp: new Date().toISOString(),
      }),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } },
    )
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e)
    return new Response(JSON.stringify({ error: msg }), {
      status: 500,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    })
  }
})

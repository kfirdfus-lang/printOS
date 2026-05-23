import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

const STORAGE_BUCKET = 'printos-files'

interface MockupRequest {
  product_id: string
  product_name_en: string
  product_ai_description: string
  color: string
  views: string[]
  print_locations: Array<{
    name: string
    view: string
    width_cm: number
    height_cm: number
    file_path?: string
    file_url?: string
  }>
  brief?: string
  request_id: string
}

function buildPrompt(req: MockupRequest): string {
  const colorMap: Record<string, string> = {
    'שחור': 'black',
    'לבן': 'white',
    'כחול': 'blue',
    'כחול נייבי': 'navy blue',
    'אדום': 'red',
    'ירוק': 'green',
    'אפור': 'gray',
    'צהוב': 'yellow',
    'כתום': 'orange',
    'ורוד': 'pink',
    'תכלת': 'light blue',
    "טבעי בז'": 'natural beige',
    'צבע מלא': 'full color',
  }

  const colorEn = colorMap[req.color] || req.color

  const locMap: Record<string, string> = {
    'חזית מרכז': 'front center',
    'חזית שמאל (לוגו)': 'front left chest (logo)',
    'חזית שמאל': 'front left',
    'גב גדול': 'back large print',
    'גב': 'back',
    'שרוול שמאל': 'left sleeve',
    'שרוול ימין': 'right sleeve',
    'כיס קדמי': 'front pocket',
    "קפוצ'ון מאחור": 'hood back print',
    'מלא': 'full surface',
    'מלא חזית': 'full front',
    'מלא גב': 'full back',
  }

  const printsDescription = req.print_locations
    .map((loc) => {
      const locEn = locMap[loc.name] || loc.name
      return `${locEn} (${loc.width_cm}×${loc.height_cm}cm)`
    })
    .join(', ')

  let prompt = `Professional product mockup photography of a ${colorEn} ${req.product_ai_description}, `
  prompt += `with branded prints at: ${printsDescription}. `
  prompt += `Studio lighting, white background, high quality commercial photography, `
  prompt += `realistic textures, clear and sharp details, professional product showcase, `
  prompt += `4K resolution.`

  if (req.brief) {
    prompt += ` Additional context: ${req.brief}`
  }

  return prompt
}

async function generateMockupAI(prompt: string): Promise<string[]> {
  const openaiKey = Deno.env.get('OPENAI_API_KEY')
  if (!openaiKey) {
    throw new Error('OPENAI_API_KEY not configured in Supabase Secrets')
  }

  const response = await fetch('https://api.openai.com/v1/images/generations', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${openaiKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      model: 'gpt-image-1',
      prompt,
      n: 1,
      size: '1024x1024',
      quality: 'high',
      output_format: 'png',
    }),
  })

  if (!response.ok) {
    const error = await response.text()
    throw new Error(`OpenAI error: ${error}`)
  }

  const data = await response.json()
  const items = data?.data || []
  const b64List = items.map((item: { b64_json?: string }) => item.b64_json).filter(Boolean)
  if (b64List.length === 0) {
    throw new Error('OpenAI returned no image data')
  }
  return b64List as string[]
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  let requestId: string | undefined

  try {
    const body: MockupRequest = await req.json()
    requestId = body.request_id

    if (!requestId) {
      return new Response(JSON.stringify({ error: 'request_id required' }), {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      })
    }

    const supabaseUrl = Deno.env.get('SUPABASE_URL')!
    const supabaseServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!
    const supabase = createClient(supabaseUrl, supabaseServiceKey)

    const { data: existing } = await supabase
      .from('design_requests')
      .select('processing_log')
      .eq('id', requestId)
      .single()

    await supabase
      .from('design_requests')
      .update({ status: 'processing', started_at: new Date().toISOString(), error_message: null })
      .eq('id', requestId)

    const prompt = buildPrompt(body)
    console.log('Generating mockup with prompt:', prompt)

    const images = await generateMockupAI(prompt)

    const timestamp = Date.now()
    const outputFiles: Array<{
      path: string
      name: string
      size: number
      type: string
      uploaded_at: string
    }> = []

    for (let i = 0; i < images.length; i++) {
      const imgB64 = images[i]
      const imgBytes = Uint8Array.from(atob(imgB64), (c) => c.charCodeAt(0))

      const fileName = `mockup_${timestamp}_${i + 1}.png`
      const filePath = `design_requests/${requestId}/output/${fileName}`

      const { error: upError } = await supabase.storage
        .from(STORAGE_BUCKET)
        .upload(filePath, imgBytes, { contentType: 'image/png', upsert: true })

      if (upError) {
        console.error('Upload error:', upError)
        continue
      }

      outputFiles.push({
        path: filePath,
        name: fileName,
        size: imgBytes.length,
        type: 'image/png',
        uploaded_at: new Date().toISOString(),
      })
    }

    if (outputFiles.length === 0) {
      throw new Error('No images generated successfully')
    }

    const existingLog = Array.isArray(existing?.processing_log) ? existing.processing_log : []

    await supabase
      .from('design_requests')
      .update({
        status: 'completed',
        completed_at: new Date().toISOString(),
        output_files: outputFiles,
        processing_log: [
          ...existingLog,
          {
            timestamp: new Date().toISOString(),
            message: `Generated ${outputFiles.length} mockup(s) via GPT-Image-1`,
            prompt,
          },
        ],
        error_message: null,
      })
      .eq('id', requestId)

    return new Response(
      JSON.stringify({ success: true, output_files: outputFiles, prompt }),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } },
    )
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e)
    console.error('Mockup error:', msg)

    if (requestId) {
      try {
        const supabaseUrl = Deno.env.get('SUPABASE_URL')!
        const supabaseServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!
        const supabase = createClient(supabaseUrl, supabaseServiceKey)
        await supabase
          .from('design_requests')
          .update({ status: 'failed', error_message: msg })
          .eq('id', requestId)
      } catch (updErr) {
        console.error('Failed to mark request failed:', updErr)
      }
    }

    return new Response(JSON.stringify({ error: msg }), {
      status: 500,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    })
  }
})

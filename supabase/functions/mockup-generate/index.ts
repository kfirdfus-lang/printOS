import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

const STORAGE_BUCKET = 'printos-files'

interface PrintLocation {
  name: string
  view: string
  width_cm: number
  height_cm: number
  file_path: string
}

interface MockupRequest {
  request_id: string
  product_id: string
  product_name_en: string
  product_ai_description: string
  color: string
  views: string[]
  print_locations: PrintLocation[]
  brief?: string
}

function colorToEn(color: string): string {
  const map: Record<string, string> = {
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
    "טבעי בז'": 'natural beige canvas',
    'צבע מלא': 'full color',
  }
  return map[color] || color
}

function locationToEn(locName: string): string {
  const map: Record<string, string> = {
    'חזית מרכז': 'front center',
    'חזית שמאל (לוגו)': 'front left chest as small logo',
    'גב גדול': 'back, large',
    'גב': 'back',
    'שרוול שמאל': 'left sleeve',
    'שרוול ימין': 'right sleeve',
    'כיס קדמי': 'front pocket',
    'חזית': 'front',
    'צד אחד': 'one side',
    'שני צדדים': 'wrapped around both sides',
    'כריכה': 'cover',
    'מלא': 'full coverage',
    'מלא חזית': 'full front',
    'מלא גב': 'full back',
    'גוף העט': 'pen body',
    'חזית שמאל': 'front left side',
  }
  return map[locName] || locName
}

async function detectTransparency(imageBytes: Uint8Array): Promise<boolean> {
  const isPng =
    imageBytes[0] === 0x89 &&
    imageBytes[1] === 0x50 &&
    imageBytes[2] === 0x4e &&
    imageBytes[3] === 0x47

  if (!isPng) {
    console.log('Not a PNG - has background')
    return false
  }

  if (imageBytes.length < 26) return false

  const colorType = imageBytes[25]
  const hasAlphaChannel = colorType === 4 || colorType === 6

  if (!hasAlphaChannel) {
    console.log('PNG without alpha channel - has background')
    return false
  }

  console.log('PNG with alpha channel - likely transparent')
  return true
}

async function removeBackground(imageBytes: Uint8Array): Promise<Uint8Array> {
  const apiKey = Deno.env.get('REMOVE_BG_API_KEY')
  if (!apiKey) {
    throw new Error('REMOVE_BG_API_KEY not configured in Supabase Secrets')
  }

  console.log('Calling Remove.bg API...')

  const formData = new FormData()
  formData.append('image_file', new Blob([imageBytes]), 'logo.png')
  formData.append('size', 'auto')
  formData.append('format', 'png')

  const response = await fetch('https://api.remove.bg/v1.0/removebg', {
    method: 'POST',
    headers: { 'X-Api-Key': apiKey },
    body: formData,
  })

  if (!response.ok) {
    const errorText = await response.text()
    console.error('Remove.bg error:', errorText)
    throw new Error(`Remove.bg failed (${response.status}): ${errorText}`)
  }

  const resultBuffer = await response.arrayBuffer()
  console.log(`Remove.bg success: ${resultBuffer.byteLength} bytes returned`)

  return new Uint8Array(resultBuffer)
}

function buildPrompt(req: MockupRequest): string {
  const colorEn = colorToEn(req.color)

  const printsDescription = req.print_locations
    .map((loc) => {
      const locEn = locationToEn(loc.name)
      return `${locEn} (size approximately ${loc.width_cm}×${loc.height_cm} cm)`
    })
    .join(', ')

  let prompt = `Place the provided transparent logo exactly as shown onto a `
  prompt += `${colorEn} ${req.product_ai_description}. `
  prompt += `Print location: ${printsDescription}. `
  prompt += `Preserve the logo's exact design, colors, shape, and text. `
  prompt += `Apply realistic fabric/material texture to the logo so it looks printed on the product. `
  prompt += `Professional product mockup photography, studio lighting, clean white background, `
  prompt += `realistic shadows, high quality commercial photography, sharp details, 4K resolution.`

  if (req.brief) {
    prompt += ` Additional context: ${req.brief}`
  }

  return prompt
}

async function generateMockupAI(logoBytes: Uint8Array, prompt: string): Promise<string> {
  const openaiKey = Deno.env.get('OPENAI_API_KEY')
  if (!openaiKey) {
    throw new Error('OPENAI_API_KEY not configured')
  }

  console.log('Calling OpenAI Images Edits API...')
  console.log('Prompt:', prompt)
  console.log('Logo size:', logoBytes.length, 'bytes')

  const formData = new FormData()
  formData.append('image', new Blob([logoBytes], { type: 'image/png' }), 'logo.png')
  formData.append('model', 'gpt-image-1')
  formData.append('prompt', prompt)
  formData.append('n', '1')
  formData.append('size', '1024x1024')
  formData.append('quality', 'high')

  const response = await fetch('https://api.openai.com/v1/images/edits', {
    method: 'POST',
    headers: { Authorization: `Bearer ${openaiKey}` },
    body: formData,
  })

  if (!response.ok) {
    const errorText = await response.text()
    console.error('OpenAI error:', errorText)
    throw new Error(`OpenAI API error (${response.status}): ${errorText}`)
  }

  const data = await response.json()

  if (!data.data?.[0]?.b64_json) {
    throw new Error('No image in OpenAI response')
  }

  return data.data[0].b64_json as string
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  const log: Array<Record<string, unknown>> = []
  let requestId: string | undefined

  try {
    const supabaseUrl = Deno.env.get('SUPABASE_URL')!
    const supabaseServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!
    const supabase = createClient(supabaseUrl, supabaseServiceKey)

    const body: MockupRequest = await req.json()
    requestId = body.request_id

    if (!requestId) {
      return new Response(JSON.stringify({ error: 'request_id required' }), {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      })
    }

    await supabase
      .from('design_requests')
      .update({ status: 'processing', started_at: new Date().toISOString(), error_message: null })
      .eq('id', requestId)

    log.push({ step: 'start', timestamp: new Date().toISOString() })

    if (!body.print_locations?.length) {
      throw new Error('No print locations provided')
    }

    const firstLocation = body.print_locations[0]
    if (!firstLocation.file_path) {
      throw new Error('No logo file path provided')
    }

    log.push({ step: 'download_logo', file_path: firstLocation.file_path })

    const { data: fileData, error: dlError } = await supabase.storage
      .from(STORAGE_BUCKET)
      .download(firstLocation.file_path)

    if (dlError || !fileData) {
      throw new Error('Failed to download logo: ' + (dlError?.message || 'unknown'))
    }

    let logoBytes = new Uint8Array(await fileData.arrayBuffer())
    log.push({ step: 'logo_downloaded', size: logoBytes.length })

    const isTransparent = await detectTransparency(logoBytes)
    log.push({ step: 'transparency_check', is_transparent: isTransparent })

    if (!isTransparent) {
      log.push({ step: 'removing_background', service: 'remove.bg' })
      try {
        logoBytes = await removeBackground(logoBytes)
        log.push({ step: 'background_removed', new_size: logoBytes.length })
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e)
        console.error('Background removal failed, continuing with original:', msg)
        log.push({
          step: 'background_removal_failed',
          error: msg,
          continued_with_original: true,
        })
      }
    } else {
      log.push({ step: 'skipping_background_removal', reason: 'already_transparent' })
    }

    const prompt = buildPrompt(body)
    log.push({ step: 'prompt_built', prompt })

    const imageB64 = await generateMockupAI(logoBytes, prompt)
    log.push({ step: 'ai_generation_complete' })

    const timestamp = Date.now()
    const fileName = `mockup_${timestamp}.png`
    const filePath = `design_requests/${requestId}/output/${fileName}`

    const imgBytes = Uint8Array.from(atob(imageB64), (c) => c.charCodeAt(0))

    const { error: upError } = await supabase.storage
      .from(STORAGE_BUCKET)
      .upload(filePath, imgBytes, { contentType: 'image/png', upsert: true })

    if (upError) {
      throw new Error('Failed to upload mockup: ' + upError.message)
    }

    log.push({ step: 'mockup_uploaded', file_path: filePath, size: imgBytes.length })

    const outputFile = {
      path: filePath,
      name: fileName,
      size: imgBytes.length,
      type: 'image/png',
      uploaded_at: new Date().toISOString(),
    }

    const outputFiles: typeof outputFile[] = [outputFile]

    if (!isTransparent) {
      const logoFileName = `logo_clean_${timestamp}.png`
      const logoFilePath = `design_requests/${requestId}/output/${logoFileName}`

      const { error: logoUpError } = await supabase.storage
        .from(STORAGE_BUCKET)
        .upload(logoFilePath, logoBytes, { contentType: 'image/png', upsert: true })

      if (!logoUpError) {
        outputFiles.push({
          path: logoFilePath,
          name: logoFileName,
          size: logoBytes.length,
          type: 'image/png',
          uploaded_at: new Date().toISOString(),
        })
        log.push({ step: 'logo_clean_saved', file_path: logoFilePath })
      }
    }

    await supabase
      .from('design_requests')
      .update({
        status: 'completed',
        completed_at: new Date().toISOString(),
        output_files: outputFiles,
        processing_log: log,
        error_message: null,
      })
      .eq('id', requestId)

    return new Response(
      JSON.stringify({ success: true, output_files: outputFiles, processing_log: log }),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } },
    )
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e)
    console.error('Mockup error:', msg)
    log.push({ step: 'error', message: msg })

    if (requestId) {
      try {
        const supabaseUrl = Deno.env.get('SUPABASE_URL')!
        const supabaseServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!
        const supabase = createClient(supabaseUrl, supabaseServiceKey)
        await supabase
          .from('design_requests')
          .update({
            status: 'failed',
            error_message: msg,
            processing_log: log,
          })
          .eq('id', requestId)
      } catch (updErr) {
        console.error('Failed to mark request failed:', updErr)
      }
    }

    return new Response(JSON.stringify({ error: msg, log }), {
      status: 500,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    })
  }
})

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { Image, decode } from 'https://deno.land/x/imagescript@1.2.15/mod.ts'

const PDFSHIFT_API_URL = 'https://api.pdfshift.io/v3/convert/pdf'

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
  position?: string
  rotation?: number
}

interface MockupRequest {
  request_id: string
  product_id: string
  product_name_en: string
  product_name_he?: string
  product_ai_description: string
  color: string
  views: string[]
  print_locations: PrintLocation[]
  brief?: string
  generate_pdf?: boolean
  client_name?: string
  project_name?: string
  precision_mode?: boolean
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

function positionToEn(pos: string): string {
  const map: Record<string, string> = {
    'top-left': 'top-left corner',
    'top-center': 'top center',
    'top-right': 'top-right corner',
    'middle-left': 'middle left',
    'middle-center': 'exact center',
    'middle-right': 'middle right',
    'bottom-left': 'bottom-left corner',
    'bottom-center': 'bottom center',
    'bottom-right': 'bottom-right corner',
  }
  return map[pos] || 'center'
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
      const posEn = positionToEn(loc.position || 'middle-center')
      const rotation = loc.rotation ? ` rotated ${loc.rotation} degrees` : ''
      return `at ${locEn} positioned ${posEn} (size EXACTLY ${loc.width_cm}×${loc.height_cm} cm, no bigger${rotation})`
    })
    .join(', and ')

  let prompt = `CRITICAL INSTRUCTIONS:\n\n`
  prompt += `Take the EXACT logo from the input image and place it onto a `
  prompt += `${colorEn} ${req.product_ai_description}.\n\n`
  prompt += `STRICT RULES:\n`
  prompt += `1. Use the EXACT logo as provided - DO NOT redraw, restyle, or modify it.\n`
  prompt += `2. Preserve every detail: exact colors, exact letters, exact shapes - 100% identical.\n`
  prompt += `3. The logo must be SMALL and proportional (approximately ${req.print_locations[0]?.width_cm || 20}cm wide).\n`
  prompt += `4. DO NOT cover the entire product surface with the logo.\n`
  prompt += `5. Position the logo ${printsDescription}.\n`
  prompt += `6. Logo appears as a printed graphic on the fabric/material with realistic shadow.\n\n`
  prompt += `OUTPUT STYLE:\n`
  prompt += `- Professional product mockup photography\n`
  prompt += `- Clean white studio background\n`
  prompt += `- Realistic shadows, soft lighting, fabric/material texture\n`
  prompt += `- High quality 4K commercial photography\n`
  prompt += `- The logo looks like a real print, NOT oversized or banner-style\n`

  if (req.brief) {
    prompt += `\nADDITIONAL CONTEXT: ${req.brief}`
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

function isSceneDescription(brief?: string): boolean {
  if (!brief || brief.trim().length < 10) return false
  const sceneKeywords = [
    'איש', 'אישה', 'אדם', 'ילד', 'משרד', 'בית', 'רחוב', 'חנות', 'מסעדה',
    'קפה', 'יושב', 'עומד', 'הולך', 'מחזיק', 'לובש', 'שותה', 'אוכל',
    'חיצוני', 'פנימי', 'רקע', 'אווירה', 'סביבה', 'תפאורה',
    'person', 'man', 'woman', 'office', 'cafe', 'street', 'shop',
    'holding', 'wearing', 'sitting', 'standing', 'background', 'scene',
  ]
  const lower = brief.toLowerCase()
  return sceneKeywords.some((kw) => lower.includes(kw))
}

function buildScenePrompt(req: MockupRequest): string {
  const colorEn = colorToEn(req.color)
  let prompt = `Take the provided product mockup and place it in a realistic scene: `
  prompt += `${req.brief}. `
  prompt += `Keep the product (${colorEn} ${req.product_ai_description}) and the logo on it `
  prompt += `EXACTLY as shown in the original. `
  prompt += `The logo design, colors, and placement must remain identical. `
  prompt += `Create a natural, professional scene around the product. `
  prompt += `High quality commercial photography, realistic lighting and shadows, 4K resolution.`
  return prompt
}

async function compositeLogoOnMockup(
  mockupBytes: Uint8Array,
  logoBytes: Uint8Array,
  location: PrintLocation,
): Promise<Uint8Array> {
  console.log('Compositing logo precisely...')

  const mockup = (await decode(mockupBytes)) as Image
  const logo = (await decode(logoBytes)) as Image

  const mockupW = mockup.width
  const mockupH = mockup.height

  const PRODUCT_AREA_RATIO = 0.6
  const PRODUCT_TYPICAL_WIDTH_CM = 40

  const pixelsPerCm = (mockupW * PRODUCT_AREA_RATIO) / PRODUCT_TYPICAL_WIDTH_CM

  const logoTargetW = Math.round(location.width_cm * pixelsPerCm)
  const logoTargetH = Math.round(location.height_cm * pixelsPerCm)

  const maxW = Math.min(logoTargetW, Math.round(mockupW * 0.6))
  const maxH = Math.min(logoTargetH, Math.round(mockupH * 0.6))

  logo.resize(maxW, maxH)

  const productAreaX = mockupW * (1 - PRODUCT_AREA_RATIO) / 2
  const productAreaY = mockupH * (1 - PRODUCT_AREA_RATIO) / 2
  const productW = mockupW * PRODUCT_AREA_RATIO
  const productH = mockupH * PRODUCT_AREA_RATIO

  let x = 0
  let y = 0
  const pos = location.position || 'middle-center'

  if (pos.includes('left')) {
    x = productAreaX + productW * 0.1
  } else if (pos.includes('right')) {
    x = productAreaX + productW * 0.9 - maxW
  } else {
    x = productAreaX + productW / 2 - maxW / 2
  }

  if (pos.includes('top')) {
    y = productAreaY + productH * 0.1
  } else if (pos.includes('bottom')) {
    y = productAreaY + productH * 0.9 - maxH
  } else {
    y = productAreaY + productH / 2 - maxH / 2
  }

  mockup.composite(logo, Math.round(x), Math.round(y))

  const result = await mockup.encode()
  return new Uint8Array(result)
}

async function buildProposalPDF(
  mockupUrl: string,
  originalLogoUrl: string,
  natalieLogoUrl: string | null,
  data: {
    client_name: string
    project_name: string
    product_name_he: string
    color: string
    print_locations: PrintLocation[]
    brief: string
  },
): Promise<Uint8Array> {
  const pdfshiftKey = Deno.env.get('PDFSHIFT_API_KEY')
  if (!pdfshiftKey) {
    throw new Error('PDFSHIFT_API_KEY not configured')
  }

  const dateStr = new Date().toLocaleDateString('he-IL')

  const escapeHtml = (text: string): string => {
    if (!text) return ''
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;')
  }

  const locationsHtml = data.print_locations
    .map(
      (loc, i) =>
        `<li><strong>${i + 1}. ${escapeHtml(loc.name)}:</strong> ${loc.width_cm} × ${loc.height_cm} ס"מ</li>`,
    )
    .join('')

  const html = `<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<title>הצעת הדמיה - ${escapeHtml(data.client_name)}</title>
<link href="https://fonts.googleapis.com/css2?family=Heebo:wght@400;500;700;900&display=swap" rel="stylesheet">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Heebo', 'Arial', sans-serif;
    color: #1f2937;
    background: white;
    direction: rtl;
    padding: 30px 40px;
  }
  .header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 3px solid #0d9488;
    padding-bottom: 20px;
    margin-bottom: 25px;
  }
  .header-left { text-align: left; }
  .header-right { text-align: right; }
  .natalie-logo { max-height: 70px; max-width: 150px; }
  .company-name { font-size: 16px; font-weight: 700; color: #0d9488; margin-bottom: 4px; }
  .company-details { font-size: 11px; color: #6b7280; line-height: 1.5; }
  .title { font-size: 24px; font-weight: 900; color: #1f2937; margin-bottom: 20px; text-align: right; }
  .client-info {
    background: #f0fdfa;
    border-right: 4px solid #0d9488;
    padding: 12px 16px;
    margin-bottom: 25px;
    border-radius: 6px;
  }
  .client-info div { margin-bottom: 6px; font-size: 13px; }
  .client-info .label { font-weight: 700; color: #0d9488; display: inline-block; width: 80px; }
  .mockup-container { text-align: center; margin: 30px 0; }
  .mockup-image {
    max-width: 500px;
    max-height: 500px;
    border: 3px solid #0d9488;
    border-radius: 8px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
  }
  .details-section {
    display: flex;
    gap: 30px;
    margin: 25px 0;
    padding: 20px;
    background: #fafafa;
    border-radius: 8px;
  }
  .original-logo-block { flex: 0 0 140px; text-align: center; }
  .original-logo-block .block-title {
    font-size: 12px;
    font-weight: 700;
    color: #0d9488;
    margin-bottom: 10px;
  }
  .original-logo-block img {
    max-width: 130px;
    max-height: 130px;
    border: 1px solid #e5e7eb;
    border-radius: 4px;
    padding: 8px;
    background: white;
  }
  .print-details { flex: 1; }
  .print-details h3 {
    font-size: 15px;
    font-weight: 700;
    color: #0d9488;
    margin-bottom: 12px;
    border-bottom: 1px solid #e5e7eb;
    padding-bottom: 6px;
  }
  .print-details ul { list-style: none; }
  .print-details li { padding: 5px 0; font-size: 12px; color: #374151; }
  .print-details li strong { color: #1f2937; }
  .approval-section { border-top: 2px solid #e5e7eb; padding-top: 25px; margin-top: 30px; }
  .approval-section h3 { font-size: 16px; font-weight: 700; color: #1f2937; margin-bottom: 18px; }
  .approval-row { display: flex; gap: 30px; margin-bottom: 15px; font-size: 12px; }
  .approval-field { flex: 1; }
  .approval-field .field-label { font-weight: 600; color: #6b7280; margin-bottom: 5px; display: block; }
  .approval-field .field-line { border-bottom: 1px solid #9ca3af; height: 24px; }
  .footer {
    margin-top: 40px;
    padding-top: 15px;
    border-top: 1px solid #e5e7eb;
    text-align: center;
    font-size: 10px;
    color: #6b7280;
  }
</style>
</head>
<body>
  <div class="header">
    <div class="header-left">
      ${
        natalieLogoUrl
          ? `<img src="${natalieLogoUrl}" class="natalie-logo" alt="Natalie">`
          : `<div class="company-name" style="font-size:22px">נטלי</div>`
      }
    </div>
    <div class="header-right">
      <div class="company-name">נטלי פתרונות הדפסה בע"מ</div>
      <div class="company-details">
        שד' הר ציון 104, תל אביב<br>
        03-6815703 | natalie-print.com
      </div>
    </div>
  </div>
  <div class="title">📋 הצעת הדמיה למוצר</div>
  <div class="client-info">
    <div><span class="label">לקוח:</span> ${escapeHtml(data.client_name) || '—'}</div>
    ${data.project_name ? `<div><span class="label">פרויקט:</span> ${escapeHtml(data.project_name)}</div>` : ''}
    <div><span class="label">תאריך:</span> ${dateStr}</div>
  </div>
  <div class="mockup-container">
    <img src="${mockupUrl}" class="mockup-image" alt="Mockup">
  </div>
  <div class="details-section">
    <div class="original-logo-block">
      <div class="block-title">לוגו מקורי</div>
      <img src="${originalLogoUrl}" alt="Logo">
    </div>
    <div class="print-details">
      <h3>📐 פרטי הדפס</h3>
      <ul>
        <li><strong>מוצר:</strong> ${escapeHtml(data.product_name_he)}</li>
        <li><strong>צבע:</strong> ${escapeHtml(data.color)}</li>
        ${locationsHtml}
      </ul>
    </div>
  </div>
  <div class="approval-section">
    <h3>✍️ אישור הזמנה</h3>
    <div class="approval-row">
      <div class="approval-field">
        <span class="field-label">שם מאשר:</span>
        <div class="field-line"></div>
      </div>
      <div class="approval-field">
        <span class="field-label">חתימה:</span>
        <div class="field-line"></div>
      </div>
      <div class="approval-field">
        <span class="field-label">תאריך:</span>
        <div class="field-line"></div>
      </div>
    </div>
  </div>
  <div class="footer">
    נטלי פתרונות הדפסה בע"מ | שד' הר ציון 104 תל אביב | 03-6815703 | natalie-print.com
  </div>
</body>
</html>`

  console.log('Calling PDFShift API with URLs (HTML size:', html.length, 'bytes)')

  const response = await fetch(PDFSHIFT_API_URL, {
    method: 'POST',
    headers: {
      'X-API-Key': pdfshiftKey,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      source: html,
      landscape: false,
      format: 'A4',
      margin: '15mm',
      use_print: true,
      sandbox: false,
    }),
  })

  if (!response.ok) {
    const errorText = await response.text()
    console.error('PDFShift error:', errorText)
    throw new Error(`PDFShift failed (${response.status}): ${errorText}`)
  }

  const pdfBytes = new Uint8Array(await response.arrayBuffer())
  console.log(`PDF created: ${pdfBytes.length} bytes`)

  return pdfBytes
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

    let imageB64 = await generateMockupAI(logoBytes, prompt)
    log.push({ step: 'ai_generation_complete' })

    let imgBytes = Uint8Array.from(atob(imageB64), (c) => c.charCodeAt(0))

    if (isSceneDescription(body.brief) && !body.precision_mode) {
      const scenePrompt = buildScenePrompt(body)
      log.push({ step: 'scene_prompt_built', prompt: scenePrompt })
      imageB64 = await generateMockupAI(imgBytes, scenePrompt)
      imgBytes = Uint8Array.from(atob(imageB64), (c) => c.charCodeAt(0))
      log.push({ step: 'scene_generation_complete' })
    }

    if (body.precision_mode && body.print_locations.length > 0) {
      try {
        log.push({ step: 'precision_mode_active' })

        const cleanPrompt =
          `Professional product mockup photography of a clean ${colorToEn(body.color)} ${body.product_ai_description}. ` +
          `Studio lighting, white background, NO logo, NO print, NO graphic on the product. Plain product only, high quality 4K.`

        log.push({ step: 'generating_clean_product' })

        const cleanB64 = await generateMockupAI(logoBytes, cleanPrompt)
        const cleanBytes = Uint8Array.from(atob(cleanB64), (c) => c.charCodeAt(0))

        imgBytes = await compositeLogoOnMockup(cleanBytes, logoBytes, body.print_locations[0])
        log.push({ step: 'precision_compositing_complete' })
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e)
        console.error('Compositing failed:', e)
        log.push({ step: 'compositing_failed', error: msg })
      }
    }

    const timestamp = Date.now()
    const fileName = `mockup_${timestamp}.png`
    const filePath = `design_requests/${requestId}/output/${fileName}`

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

    if (body.generate_pdf) {
      try {
        log.push({ step: 'generating_proposal_pdf_with_urls' })

        const { data: mockupUrlData } = supabase.storage.from(STORAGE_BUCKET).getPublicUrl(filePath)
        const logoPathForPdf =
          outputFiles.find((f) => f.name.startsWith('logo_clean_'))?.path ?? firstLocation.file_path
        const { data: logoUrlData } = supabase.storage.from(STORAGE_BUCKET).getPublicUrl(logoPathForPdf)

        const mockupUrl = mockupUrlData.publicUrl
        const logoUrl = logoUrlData.publicUrl

        log.push({ step: 'urls_created', mockup_url: mockupUrl, logo_url: logoUrl })

        let natalieLogoUrl: string | null = null
        try {
          const { data: natalieCheck } = await supabase.storage
            .from(STORAGE_BUCKET)
            .download('branding/natalie_logo.png')

          if (natalieCheck) {
            const { data: natalieUrlData } = supabase.storage
              .from(STORAGE_BUCKET)
              .getPublicUrl('branding/natalie_logo.png')
            natalieLogoUrl = natalieUrlData.publicUrl
            log.push({ step: 'natalie_logo_url_created' })
          }
        } catch {
          log.push({ step: 'natalie_logo_not_found' })
        }

        const pdfBytes = await buildProposalPDF(mockupUrl, logoUrl, natalieLogoUrl, {
          client_name: body.client_name || '',
          project_name: body.project_name || '',
          product_name_he: body.product_name_he || body.product_name_en,
          color: body.color,
          print_locations: body.print_locations,
          brief: body.brief || '',
        })

        const pdfFileName = `proposal_${timestamp}.pdf`
        const pdfFilePath = `design_requests/${requestId}/output/${pdfFileName}`

        const { error: pdfUpError } = await supabase.storage
          .from(STORAGE_BUCKET)
          .upload(pdfFilePath, pdfBytes, { contentType: 'application/pdf', upsert: true })

        if (!pdfUpError) {
          outputFiles.push({
            path: pdfFilePath,
            name: pdfFileName,
            size: pdfBytes.length,
            type: 'application/pdf',
            uploaded_at: new Date().toISOString(),
          })
          log.push({ step: 'pdf_uploaded', file_path: pdfFilePath, size: pdfBytes.length })
        } else {
          log.push({ step: 'pdf_upload_failed', error: pdfUpError.message })
        }
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e)
        console.error('PDF generation failed:', e)
        log.push({ step: 'pdf_failed', error: msg })
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

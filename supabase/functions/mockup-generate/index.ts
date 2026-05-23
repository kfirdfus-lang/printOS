import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { PDFDocument, rgb, StandardFonts } from 'https://esm.sh/pdf-lib@1.17.1?target=deno'
import fontkit from 'https://esm.sh/@pdf-lib/fontkit@1.1.1?target=deno'

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
      return `${locEn} at ${posEn} (size ${loc.width_cm}×${loc.height_cm} cm${rotation})`
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

type PdfTextOpts = { size?: number; font?: unknown; color?: ReturnType<typeof rgb> }

async function buildProposalPDF(
  mockupBytes: Uint8Array,
  originalLogoBytes: Uint8Array,
  natalieLogoBytes: Uint8Array | null,
  heeboRegularBytes: Uint8Array | null,
  heeboBoldBytes: Uint8Array | null,
  data: {
    client_name: string
    project_name: string
    product_name_he: string
    color: string
    print_locations: PrintLocation[]
    brief: string
  },
): Promise<Uint8Array> {
  const pdfDoc = await PDFDocument.create()
  pdfDoc.registerFontkit(fontkit)
  const page = pdfDoc.addPage([595, 842])

  let font
  let fontBold
  let hasHebrew = false

  if (heeboRegularBytes && heeboBoldBytes) {
    try {
      font = await pdfDoc.embedFont(heeboRegularBytes, { subset: true })
      fontBold = await pdfDoc.embedFont(heeboBoldBytes, { subset: true })
      hasHebrew = true
      console.log('Heebo Hebrew font loaded')
    } catch (e) {
      console.error('Failed to load Heebo, falling back:', e)
      font = await pdfDoc.embedFont(StandardFonts.Helvetica)
      fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold)
    }
  } else {
    console.log('Heebo not available - using Helvetica')
    font = await pdfDoc.embedFont(StandardFonts.Helvetica)
    fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold)
  }

  const drawTextRTL = (text: string, x: number, y: number, opts: PdfTextOpts = {}) => {
    const size = opts.size || 11
    const fontToUse = (opts.font || font) as typeof font
    const color = opts.color || rgb(0.2, 0.2, 0.2)
    if (hasHebrew) {
      const textWidth = fontToUse.widthOfTextAtSize(text, size)
      page.drawText(text, { x: x - textWidth, y, size, font: fontToUse, color })
    } else {
      page.drawText(text, { x, y, size, font: fontToUse, color })
    }
  }

  const drawTextLTR = (text: string, x: number, y: number, opts: PdfTextOpts = {}) => {
    page.drawText(text, {
      x,
      y,
      size: opts.size || 11,
      font: (opts.font || font) as typeof font,
      color: opts.color || rgb(0.2, 0.2, 0.2),
    })
  }

  let y = 800

  if (natalieLogoBytes) {
    try {
      let natalieImg
      try {
        natalieImg = await pdfDoc.embedPng(natalieLogoBytes)
      } catch {
        natalieImg = await pdfDoc.embedJpg(natalieLogoBytes)
      }
      const maxLogoH = 60
      const ratio = maxLogoH / natalieImg.height
      const logoW = natalieImg.width * ratio
      const logoH = maxLogoH
      page.drawImage(natalieImg, { x: 30, y: y - logoH, width: logoW, height: logoH })
      if (hasHebrew) {
        drawTextRTL('נטלי פתרונות הדפסה בע"מ', 565, y - 15, {
          size: 12,
          font: fontBold,
          color: rgb(0.05, 0.58, 0.53),
        })
        drawTextRTL('שד\' הר ציון 104, תל אביב', 565, y - 32, { size: 9 })
        drawTextRTL('03-6815703 | natalie-print.com', 565, y - 47, { size: 9 })
      } else {
        drawTextLTR('Natalie Print Solutions Ltd.', 410, y - 15, {
          size: 11,
          font: fontBold,
          color: rgb(0.05, 0.58, 0.53),
        })
        drawTextLTR('Har Tzion 104, Tel Aviv', 410, y - 32, { size: 9 })
        drawTextLTR('03-6815703 | natalie-print.com', 410, y - 47, { size: 9 })
      }
      y -= 80
    } catch (e) {
      console.error('Failed to embed Natalie logo:', e)
    }
  } else {
    if (hasHebrew) {
      drawTextRTL('נטלי פתרונות הדפסה בע"מ', 565, y - 20, {
        size: 18,
        font: fontBold,
        color: rgb(0.05, 0.58, 0.53),
      })
    } else {
      drawTextLTR('Natalie Print Solutions', 30, y - 20, {
        size: 18,
        font: fontBold,
        color: rgb(0.05, 0.58, 0.53),
      })
    }
    y -= 50
  }

  page.drawLine({
    start: { x: 30, y },
    end: { x: 565, y },
    thickness: 1.5,
    color: rgb(0.05, 0.58, 0.53),
  })
  y -= 30

  if (hasHebrew) {
    drawTextRTL('הצעת הדמיה למוצר', 565, y, { size: 18, font: fontBold, color: rgb(0.1, 0.1, 0.1) })
  } else {
    drawTextLTR('PRODUCT MOCKUP PROPOSAL', 30, y, { size: 16, font: fontBold })
  }
  y -= 30

  if (hasHebrew) {
    drawTextRTL(`לקוח: ${data.client_name || '—'}`, 565, y, { size: 12, font: fontBold })
    y -= 20
    if (data.project_name) {
      drawTextRTL(`פרויקט: ${data.project_name}`, 565, y, { size: 11 })
      y -= 18
    }
    drawTextRTL(`תאריך: ${new Date().toLocaleDateString('he-IL')}`, 565, y, { size: 11 })
  } else {
    drawTextLTR(`Client: ${data.client_name || '—'}`, 30, y, { size: 12, font: fontBold })
    y -= 18
    if (data.project_name) {
      drawTextLTR(`Project: ${data.project_name}`, 30, y, { size: 11 })
      y -= 18
    }
    drawTextLTR(`Date: ${new Date().toLocaleDateString('en-GB')}`, 30, y, { size: 11 })
  }
  y -= 35

  try {
    const mockupImg = await pdfDoc.embedPng(mockupBytes)
    const maxWidth = 380
    const maxHeight = 320
    const ratio = Math.min(maxWidth / mockupImg.width, maxHeight / mockupImg.height)
    const w = mockupImg.width * ratio
    const h = mockupImg.height * ratio
    const xCenter = (595 - w) / 2
    page.drawRectangle({
      x: xCenter - 3,
      y: y - h - 3,
      width: w + 6,
      height: h + 6,
      borderColor: rgb(0.05, 0.58, 0.53),
      borderWidth: 2,
    })
    page.drawImage(mockupImg, { x: xCenter, y: y - h, width: w, height: h })
    y -= h + 25
  } catch (e) {
    console.error('Failed to embed mockup:', e)
    y -= 200
  }

  const detailsStartY = y
  try {
    let origLogo
    try {
      origLogo = await pdfDoc.embedPng(originalLogoBytes)
    } catch {
      origLogo = await pdfDoc.embedJpg(originalLogoBytes)
    }
    const logoMaxSize = 90
    const ratio = Math.min(logoMaxSize / origLogo.width, logoMaxSize / origLogo.height)
    const w = origLogo.width * ratio
    const h = origLogo.height * ratio
    if (hasHebrew) {
      drawTextLTR('לוגו מקורי:', 30, y, { size: 10, font: fontBold })
    } else {
      drawTextLTR('Original Logo:', 30, y, { size: 10, font: fontBold })
    }
    page.drawImage(origLogo, { x: 30, y: y - 12 - h, width: w, height: h })
  } catch (e) {
    console.error('Failed to embed original logo:', e)
  }

  let detailsY = detailsStartY
  if (hasHebrew) {
    drawTextRTL('פרטי הדפס:', 565, detailsY, { size: 13, font: fontBold, color: rgb(0.05, 0.58, 0.53) })
    detailsY -= 22
    drawTextRTL(`מוצר: ${data.product_name_he}`, 565, detailsY, { size: 10 })
    detailsY -= 16
    drawTextRTL(`צבע: ${data.color}`, 565, detailsY, { size: 10 })
    detailsY -= 16
    data.print_locations.forEach((loc, i) => {
      drawTextRTL(`${i + 1}. ${loc.name}: ${loc.width_cm} × ${loc.height_cm} ס"מ`, 565, detailsY, {
        size: 10,
      })
      detailsY -= 16
    })
  } else {
    drawTextLTR('PRINT DETAILS:', 320, detailsY, { size: 12, font: fontBold, color: rgb(0.05, 0.58, 0.53) })
    detailsY -= 18
    drawTextLTR(`Product: ${data.product_name_he}`, 320, detailsY, { size: 10 })
    detailsY -= 14
    drawTextLTR(`Color: ${data.color}`, 320, detailsY, { size: 10 })
    detailsY -= 14
    data.print_locations.forEach((loc, i) => {
      drawTextLTR(`${i + 1}. ${loc.name}: ${loc.width_cm} x ${loc.height_cm} cm`, 320, detailsY, {
        size: 10,
      })
      detailsY -= 14
    })
  }
  y = Math.min(y - 110, detailsY - 10)

  page.drawLine({ start: { x: 30, y }, end: { x: 565, y }, thickness: 0.5, color: rgb(0.7, 0.7, 0.7) })
  y -= 25

  if (hasHebrew) {
    drawTextRTL('אישור הזמנה:', 565, y, { size: 13, font: fontBold, color: rgb(0.1, 0.1, 0.1) })
    y -= 25
    for (const label of ['שם מאשר:', 'חתימה:', 'תאריך:']) {
      drawTextRTL(`${label}  ____________________________`, 565, y, { size: 10 })
      y -= 22
    }
  } else {
    drawTextLTR('APPROVAL:', 30, y, { size: 12, font: fontBold })
    y -= 22
    for (const label of ['Name', 'Signature', 'Date']) {
      drawTextLTR(`${label}: ____________________________`, 30, y, { size: 10 })
      y -= 20
    }
  }

  y = 40
  page.drawLine({
    start: { x: 30, y: y + 15 },
    end: { x: 565, y: y + 15 },
    thickness: 0.5,
    color: rgb(0.7, 0.7, 0.7),
  })
  if (hasHebrew) {
    drawTextRTL(
      'נטלי פתרונות הדפסה בע"מ | שד\' הר ציון 104 תל אביב | 03-6815703 | natalie-print.com',
      565,
      y,
      { size: 8, color: rgb(0.5, 0.5, 0.5) },
    )
  } else {
    drawTextLTR(
      'Natalie Print Solutions Ltd. | Har Tzion 104 Tel Aviv | 03-6815703 | natalie-print.com',
      30,
      y,
      { size: 8, color: rgb(0.5, 0.5, 0.5) },
    )
  }

  return await pdfDoc.save()
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

    if (isSceneDescription(body.brief)) {
      const scenePrompt = buildScenePrompt(body)
      log.push({ step: 'scene_prompt_built', prompt: scenePrompt })
      imageB64 = await generateMockupAI(imgBytes, scenePrompt)
      imgBytes = Uint8Array.from(atob(imageB64), (c) => c.charCodeAt(0))
      log.push({ step: 'scene_generation_complete' })
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
      log.push({ step: 'pdf_generation_start' })

      let heeboRegular: Uint8Array | null = null
      let heeboBold: Uint8Array | null = null

      try {
        const { data: regularData } = await supabase.storage
          .from(STORAGE_BUCKET)
          .download('fonts/Heebo-Regular.ttf')
        if (regularData) {
          heeboRegular = new Uint8Array(await regularData.arrayBuffer())
        }
      } catch {
        console.log('Heebo Regular not found')
      }

      try {
        const { data: boldData } = await supabase.storage
          .from(STORAGE_BUCKET)
          .download('fonts/Heebo-Bold.ttf')
        if (boldData) {
          heeboBold = new Uint8Array(await boldData.arrayBuffer())
        }
      } catch {
        console.log('Heebo Bold not found')
      }

      if (heeboRegular && heeboBold) {
        log.push({ step: 'heebo_fonts_loaded' })
      } else {
        log.push({ step: 'heebo_fonts_missing', fallback: 'helvetica_english' })
      }

      let natalieLogoBytes: Uint8Array | null = null
      try {
        const { data: natalieData } = await supabase.storage
          .from(STORAGE_BUCKET)
          .download('branding/natalie_logo.png')
        if (natalieData) {
          natalieLogoBytes = new Uint8Array(await natalieData.arrayBuffer())
          log.push({ step: 'natalie_logo_loaded' })
        }
      } catch {
        log.push({ step: 'natalie_logo_missing' })
      }

      const pdfBytes = await buildProposalPDF(
        imgBytes,
        logoBytes,
        natalieLogoBytes,
        heeboRegular,
        heeboBold,
        {
          client_name: body.client_name || '',
          project_name: body.project_name || '',
          product_name_he: body.product_name_he || body.product_name_en,
          color: body.color,
          print_locations: body.print_locations,
          brief: body.brief || '',
        },
      )

      const pdfFileName = `proposal_${timestamp}.pdf`
      const pdfFilePath = `design_requests/${requestId}/output/${pdfFileName}`

      const { error: pdfUpError } = await supabase.storage
        .from(STORAGE_BUCKET)
        .upload(pdfFilePath, pdfBytes, { contentType: 'application/pdf', upsert: true })

      if (pdfUpError) {
        throw new Error('Failed to upload PDF: ' + pdfUpError.message)
      }

      outputFiles.push({
        path: pdfFilePath,
        name: pdfFileName,
        size: pdfBytes.length,
        type: 'application/pdf',
        uploaded_at: new Date().toISOString(),
      })
      log.push({ step: 'pdf_uploaded', file_path: pdfFilePath, size: pdfBytes.length })
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

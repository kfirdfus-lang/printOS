import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { Image, decode } from 'https://deno.land/x/imagescript@1.2.17/mod.ts'

const PDFSHIFT_API_URL = 'https://api.pdfshift.io/v3/convert/pdf'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

const STORAGE_BUCKET = 'printos-files'

interface CanvasPosition {
  pixel_x: number
  pixel_y: number
  pixel_width: number
  pixel_height: number
  stage_width: number
  stage_height: number
  rotation: number
  logo_width_cm: number
  logo_height_cm: number
}

interface PrintLocationRow {
  name: string
  width_cm: number
  height_cm: number
  [key: string]: unknown
}

interface FinalizeRequest {
  request_id: string
  position: CanvasPosition
  generate_pdf?: boolean
  client_name?: string
  project_name?: string
}

async function compressImageToJpeg(
  bytes: Uint8Array,
  quality = 75,
  maxWidth = 800,
): Promise<Uint8Array> {
  try {
    const img = (await decode(bytes)) as Image
    if (img.width > maxWidth) {
      const ratio = maxWidth / img.width
      img.resize(maxWidth, Math.round(img.height * ratio))
    }
    const jpegBytes = await img.encodeJPEG(quality)
    return new Uint8Array(jpegBytes)
  } catch (e) {
    console.error('Compression failed:', e)
    return bytes
  }
}

async function compositeLogoOnProduct(
  productBytes: Uint8Array,
  logoBytes: Uint8Array,
  position: CanvasPosition,
): Promise<Uint8Array> {
  console.log('Starting compositing...')

  const product = (await decode(productBytes)) as Image
  const logo = (await decode(logoBytes)) as Image

  const scaleX = product.width / position.stage_width
  const scaleY = product.height / position.stage_height

  const logoFinalWidth = Math.round(position.pixel_width * scaleX)
  const logoFinalHeight = Math.round(position.pixel_height * scaleY)
  const logoFinalX = Math.round(position.pixel_x * scaleX)
  const logoFinalY = Math.round(position.pixel_y * scaleY)

  console.log(
    `Logo scaled to: ${logoFinalWidth}×${logoFinalHeight} at (${logoFinalX}, ${logoFinalY}), rotation=${position.rotation}`,
  )

  logo.resize(logoFinalWidth, logoFinalHeight)

  let logoToPlace = logo
  if (position.rotation && position.rotation !== 0) {
    logoToPlace = logo.rotate(position.rotation)
  }

  product.composite(logoToPlace, logoFinalX, logoFinalY)

  const result = await product.encode()
  return new Uint8Array(result)
}

async function buildProposalPDF(
  mockupBytes: Uint8Array,
  originalLogoBytes: Uint8Array,
  natalieLogoBytes: Uint8Array | null,
  data: {
    client_name: string
    project_name: string
    product_name_he: string
    color: string
    print_locations: PrintLocationRow[]
  },
): Promise<Uint8Array> {
  const pdfshiftKey = Deno.env.get('PDFSHIFT_API_KEY')
  if (!pdfshiftKey) throw new Error('PDFSHIFT_API_KEY not configured')

  const compressedMockup = await compressImageToJpeg(mockupBytes, 75, 800)
  const compressedLogo = await compressImageToJpeg(originalLogoBytes, 80, 400)
  const compressedNatalie = natalieLogoBytes
    ? await compressImageToJpeg(natalieLogoBytes, 80, 300)
    : null

  const toBase64 = (bytes: Uint8Array) => {
    let binary = ''
    for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i])
    return btoa(binary)
  }

  const mockupBase64 = toBase64(compressedMockup)
  const logoBase64 = toBase64(compressedLogo)
  const natalieBase64 = compressedNatalie ? toBase64(compressedNatalie) : null

  const dateStr = new Date().toLocaleDateString('he-IL')
  const escapeHtml = (t: string) =>
    (t || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

  const html = `<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Heebo:wght@400;500;700;900&display=swap" rel="stylesheet">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Heebo', 'Arial', sans-serif; color: #1f2937; direction: rtl; padding: 30px 40px; }
  .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 3px solid #0d9488; padding-bottom: 20px; margin-bottom: 25px; }
  .header-left { text-align: left; }
  .header-right { text-align: right; }
  .natalie-logo { max-height: 70px; max-width: 150px; }
  .company-name { font-size: 16px; font-weight: 700; color: #0d9488; margin-bottom: 4px; }
  .company-details { font-size: 11px; color: #6b7280; line-height: 1.5; }
  .title { font-size: 24px; font-weight: 900; margin-bottom: 20px; text-align: right; }
  .client-info { background: #f0fdfa; border-right: 4px solid #0d9488; padding: 12px 16px; margin-bottom: 25px; border-radius: 6px; }
  .client-info div { margin-bottom: 6px; font-size: 13px; }
  .client-info .label { font-weight: 700; color: #0d9488; display: inline-block; width: 80px; }
  .mockup-container { text-align: center; margin: 30px 0; }
  .mockup-image { max-width: 500px; max-height: 500px; border: 3px solid #0d9488; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
  .details-section { display: flex; gap: 30px; margin: 25px 0; padding: 20px; background: #fafafa; border-radius: 8px; }
  .original-logo-block { flex: 0 0 140px; text-align: center; }
  .original-logo-block .block-title { font-size: 12px; font-weight: 700; color: #0d9488; margin-bottom: 10px; }
  .original-logo-block img { max-width: 130px; max-height: 130px; border: 1px solid #e5e7eb; border-radius: 4px; padding: 8px; background: white; }
  .print-details { flex: 1; }
  .print-details h3 { font-size: 15px; font-weight: 700; color: #0d9488; margin-bottom: 12px; border-bottom: 1px solid #e5e7eb; padding-bottom: 6px; }
  .print-details ul { list-style: none; }
  .print-details li { padding: 5px 0; font-size: 12px; }
  .approval-section { border-top: 2px solid #e5e7eb; padding-top: 25px; margin-top: 30px; }
  .approval-section h3 { font-size: 16px; font-weight: 700; margin-bottom: 18px; }
  .approval-row { display: flex; gap: 30px; margin-bottom: 15px; font-size: 12px; }
  .approval-field { flex: 1; }
  .approval-field .field-label { font-weight: 600; color: #6b7280; margin-bottom: 5px; display: block; }
  .approval-field .field-line { border-bottom: 1px solid #9ca3af; height: 24px; }
  .footer { margin-top: 40px; padding-top: 15px; border-top: 1px solid #e5e7eb; text-align: center; font-size: 10px; color: #6b7280; }
</style>
</head>
<body>
  <div class="header">
    <div class="header-left">
      ${natalieBase64 ? `<img src="data:image/jpeg;base64,${natalieBase64}" class="natalie-logo">` : `<div class="company-name" style="font-size:22px">נטלי</div>`}
    </div>
    <div class="header-right">
      <div class="company-name">נטלי פתרונות הדפסה בע"מ</div>
      <div class="company-details">שד' הר ציון 104, תל אביב<br>03-6815703 | natalie-print.com</div>
    </div>
  </div>
  <div class="title">📋 הצעת הדמיה למוצר</div>
  <div class="client-info">
    <div><span class="label">לקוח:</span> ${escapeHtml(data.client_name) || '—'}</div>
    ${data.project_name ? `<div><span class="label">פרויקט:</span> ${escapeHtml(data.project_name)}</div>` : ''}
    <div><span class="label">תאריך:</span> ${dateStr}</div>
  </div>
  <div class="mockup-container">
    <img src="data:image/jpeg;base64,${mockupBase64}" class="mockup-image" alt="Mockup">
  </div>
  <div class="details-section">
    <div class="original-logo-block">
      <div class="block-title">לוגו מקורי</div>
      <img src="data:image/jpeg;base64,${logoBase64}" alt="Logo">
    </div>
    <div class="print-details">
      <h3>📐 פרטי הדפס</h3>
      <ul>
        <li><strong>מוצר:</strong> ${escapeHtml(data.product_name_he)}</li>
        <li><strong>צבע:</strong> ${escapeHtml(data.color)}</li>
        ${data.print_locations.map((loc, i) => `<li><strong>${i + 1}. ${escapeHtml(loc.name)}:</strong> ${loc.width_cm} × ${loc.height_cm} ס"מ</li>`).join('')}
      </ul>
    </div>
  </div>
  <div class="approval-section">
    <h3>✍️ אישור הזמנה</h3>
    <div class="approval-row">
      <div class="approval-field"><span class="field-label">שם מאשר:</span><div class="field-line"></div></div>
      <div class="approval-field"><span class="field-label">חתימה:</span><div class="field-line"></div></div>
      <div class="approval-field"><span class="field-label">תאריך:</span><div class="field-line"></div></div>
    </div>
  </div>
  <div class="footer">נטלי פתרונות הדפסה בע"מ | שד' הר ציון 104 תל אביב | 03-6815703 | natalie-print.com</div>
</body>
</html>`

  const response = await fetch(PDFSHIFT_API_URL, {
    method: 'POST',
    headers: { 'X-API-Key': pdfshiftKey, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      source: html,
      landscape: false,
      format: 'A4',
      margin: '15mm',
      use_print: true,
    }),
  })

  if (!response.ok) {
    const errorText = await response.text()
    throw new Error(`PDFShift failed (${response.status}): ${errorText}`)
  }

  return new Uint8Array(await response.arrayBuffer())
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  const log: Record<string, unknown>[] = []
  let requestId: string | undefined

  try {
    const body: FinalizeRequest = await req.json()
    requestId = body.request_id
    const { position } = body

    if (!requestId || !position) {
      throw new Error('request_id and position required')
    }

    log.push({ step: 'finalize_start', position })

    const supabaseUrl = Deno.env.get('SUPABASE_URL')!
    const supabaseServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!
    const supabase = createClient(supabaseUrl, supabaseServiceKey)

    await supabase.from('design_requests').update({ status: 'processing' }).eq('id', requestId)

    const { data: request, error: reqError } = await supabase
      .from('design_requests')
      .select('*')
      .eq('id', requestId)
      .single()

    if (reqError || !request) throw new Error('Request not found')

    const outputFiles = (request.output_files || []) as { name: string; path: string }[]
    const productFile = outputFiles.find(
      (f) => f.name.startsWith('mockup_') && !f.name.startsWith('mockup_final_'),
    )
    const cleanLogoFile = outputFiles.find((f) => f.name.startsWith('logo_clean_'))

    if (!productFile) throw new Error('No product mockup found')

    log.push({ step: 'loading_images' })

    const { data: productData, error: productDlError } = await supabase.storage
      .from(STORAGE_BUCKET)
      .download(productFile.path)

    if (productDlError || !productData) {
      throw new Error('Failed to download product: ' + (productDlError?.message || 'unknown'))
    }

    const productBytes = new Uint8Array(await productData.arrayBuffer())

    let logoBytes: Uint8Array
    let originalLogoBytes: Uint8Array

    if (cleanLogoFile) {
      const { data: logoData, error: logoDlError } = await supabase.storage
        .from(STORAGE_BUCKET)
        .download(cleanLogoFile.path)
      if (logoDlError || !logoData) {
        throw new Error('Failed to download clean logo: ' + (logoDlError?.message || 'unknown'))
      }
      logoBytes = new Uint8Array(await logoData.arrayBuffer())
      originalLogoBytes = logoBytes
    } else {
      const inputFiles = (request.input_files || []) as { path: string }[]
      const inputLogo = inputFiles[0]
      if (!inputLogo?.path) throw new Error('No input logo')

      const { data: logoData, error: logoDlError } = await supabase.storage
        .from(STORAGE_BUCKET)
        .download(inputLogo.path)
      if (logoDlError || !logoData) {
        throw new Error('Failed to download input logo: ' + (logoDlError?.message || 'unknown'))
      }
      logoBytes = new Uint8Array(await logoData.arrayBuffer())
      originalLogoBytes = logoBytes
    }

    log.push({ step: 'compositing_start' })

    const finalBytes = await compositeLogoOnProduct(productBytes, logoBytes, position)

    log.push({ step: 'compositing_complete', size: finalBytes.length })

    const timestamp = Date.now()
    const finalFileName = `mockup_final_${timestamp}.png`
    const finalFilePath = `design_requests/${requestId}/output/${finalFileName}`

    const { error: uploadError } = await supabase.storage
      .from(STORAGE_BUCKET)
      .upload(finalFilePath, finalBytes, { contentType: 'image/png', upsert: true })

    if (uploadError) throw new Error('Failed to upload final: ' + uploadError.message)

    log.push({ step: 'final_uploaded' })

    const finalOutput = {
      path: finalFilePath,
      name: finalFileName,
      size: finalBytes.length,
      type: 'image/png',
      uploaded_at: new Date().toISOString(),
    }

    let pdfOutput: typeof finalOutput | null = null

    if (body.generate_pdf) {
      try {
        log.push({ step: 'generating_pdf' })

        let natalieLogoBytes: Uint8Array | null = null
        try {
          const { data: natalieData } = await supabase.storage
            .from(STORAGE_BUCKET)
            .download('branding/natalie_logo.png')
          if (natalieData) {
            natalieLogoBytes = new Uint8Array(await natalieData.arrayBuffer())
          }
        } catch {
          log.push({ step: 'natalie_logo_not_found' })
        }

        const params = (request.parameters || {}) as Record<string, unknown>
        const locations = ((params.locations || []) as PrintLocationRow[]).map((loc, i) =>
          i === 0
            ? {
                ...loc,
                width_cm: position.logo_width_cm ?? loc.width_cm,
                height_cm: position.logo_height_cm ?? loc.height_cm,
              }
            : loc,
        )

        const pdfBytes = await buildProposalPDF(finalBytes, originalLogoBytes, natalieLogoBytes, {
          client_name: body.client_name || '',
          project_name: body.project_name || '',
          product_name_he: (params.product_name as string) || 'מוצר',
          color: (params.color as string) || '',
          print_locations: locations,
        })

        const pdfFileName = `proposal_final_${timestamp}.pdf`
        const pdfFilePath = `design_requests/${requestId}/output/${pdfFileName}`

        const { error: pdfUpError } = await supabase.storage
          .from(STORAGE_BUCKET)
          .upload(pdfFilePath, pdfBytes, { contentType: 'application/pdf', upsert: true })

        if (!pdfUpError) {
          pdfOutput = {
            path: pdfFilePath,
            name: pdfFileName,
            size: pdfBytes.length,
            type: 'application/pdf',
            uploaded_at: new Date().toISOString(),
          }
          log.push({ step: 'pdf_uploaded' })
        } else {
          log.push({ step: 'pdf_upload_failed', error: pdfUpError.message })
        }
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e)
        console.error('PDF failed:', e)
        log.push({ step: 'pdf_failed', error: msg })
      }
    }

    const allOutputs = [...outputFiles, finalOutput]
    if (pdfOutput) allOutputs.push(pdfOutput)

    const prevLog = Array.isArray(request.processing_log) ? request.processing_log : []

    await supabase
      .from('design_requests')
      .update({
        status: 'completed',
        completed_at: new Date().toISOString(),
        output_files: allOutputs,
        processing_log: [...prevLog, ...log],
        error_message: null,
      })
      .eq('id', requestId)

    log.push({ step: 'finalize_complete' })

    return new Response(
      JSON.stringify({
        success: true,
        final_output: finalOutput,
        pdf_output: pdfOutput,
        log,
      }),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } },
    )
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e)
    console.error('Finalize error:', msg)
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

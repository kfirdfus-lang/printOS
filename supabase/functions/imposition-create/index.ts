import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { PDFDocument, rgb } from 'https://esm.sh/pdf-lib@1.17.1?target=deno'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

const MM_TO_PT = 72 / 25.4
const MM_TO_CM = 0.1
const REG_MARK_MARGIN_MM = 15
const REG_MARK_SIZE_MM = 3
const STORAGE_BUCKET = 'printos-files'

interface ImpositionParams {
  shape: 'circle' | 'square' | 'rectangle'
  product_width_mm: number
  product_height_mm: number
  quantity: number
  plate_width_mm: number
  plate_height_mm: number
  bleed_mm: number
  gutter_mm: number
}

interface LayoutPosition {
  x: number
  y: number
  col: number
  row: number
}

function calculateLayout(p: ImpositionParams) {
  const safeMargin = REG_MARK_MARGIN_MM + REG_MARK_SIZE_MM + 5
  const usableWidth = p.plate_width_mm - safeMargin * 2
  const usableHeight = p.plate_height_mm - safeMargin * 2

  const itemWithBleed_w = p.product_width_mm + p.bleed_mm * 2
  const itemWithBleed_h = p.product_height_mm + p.bleed_mm * 2

  const cellWidth = itemWithBleed_w + p.gutter_mm
  const cellHeight = itemWithBleed_h + p.gutter_mm

  const cols = Math.max(0, Math.floor((usableWidth + p.gutter_mm) / cellWidth))
  const rows = Math.max(0, Math.floor((usableHeight + p.gutter_mm) / cellHeight))
  const perPlate = cols * rows
  const totalPlates = perPlate > 0 ? Math.ceil(p.quantity / perPlate) : 0
  const lastPlateCount = perPlate > 0 ? p.quantity - (totalPlates - 1) * perPlate : 0

  const totalContentWidth = cols * cellWidth - p.gutter_mm
  const totalContentHeight = rows * cellHeight - p.gutter_mm
  const startX = (p.plate_width_mm - totalContentWidth) / 2
  const startY = (p.plate_height_mm - totalContentHeight) / 2

  const positions: LayoutPosition[] = []
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      positions.push({
        x: startX + col * cellWidth,
        y: startY + (rows - 1 - row) * cellHeight,
        col,
        row,
      })
    }
  }

  return {
    cols,
    rows,
    perPlate,
    totalPlates,
    lastPlateCount,
    cellWidth,
    cellHeight,
    startX,
    startY,
    positions,
    itemWithBleed_w,
    itemWithBleed_h,
  }
}

function getRegistrationMarks(plateWidth: number, plateHeight: number) {
  const m = REG_MARK_MARGIN_MM
  return [
    { x: m, y: m, label: 'bottom-left' },
    { x: plateWidth - m, y: m, label: 'bottom-right' },
    { x: plateWidth / 2, y: m, label: 'bottom-center' },
    { x: m, y: plateHeight - m, label: 'top-left' },
    { x: plateWidth - m, y: plateHeight - m, label: 'top-right' },
  ]
}

async function buildPlatePDF(
  designBytes: Uint8Array,
  params: ImpositionParams,
  layout: ReturnType<typeof calculateLayout>,
  itemsOnThisPlate: number,
): Promise<Uint8Array> {
  const platePdf = await PDFDocument.create()
  const [embeddedPage] = await platePdf.embedPdf(designBytes, [0])

  const plateWidthPt = params.plate_width_mm * MM_TO_PT
  const plateHeightPt = params.plate_height_mm * MM_TO_PT
  const page = platePdf.addPage([plateWidthPt, plateHeightPt])

  const positionsToUse = layout.positions.slice(0, itemsOnThisPlate)

  for (const pos of positionsToUse) {
    const xPt = pos.x * MM_TO_PT
    const yPt = pos.y * MM_TO_PT
    const widthPt = params.product_width_mm * MM_TO_PT
    const heightPt = params.product_height_mm * MM_TO_PT
    const bleedPt = params.bleed_mm * MM_TO_PT

    page.drawPage(embeddedPage, {
      x: xPt - bleedPt,
      y: yPt - bleedPt,
      width: widthPt + bleedPt * 2,
      height: heightPt + bleedPt * 2,
    })
  }

  const marks = getRegistrationMarks(params.plate_width_mm, params.plate_height_mm)
  const markDiameterPt = REG_MARK_SIZE_MM * MM_TO_PT

  for (const mark of marks) {
    page.drawCircle({
      x: mark.x * MM_TO_PT,
      y: mark.y * MM_TO_PT,
      size: markDiameterPt,
      color: rgb(0, 0, 0),
      borderColor: rgb(0, 0, 0),
      borderWidth: 0,
    })
  }

  const labelText = `Plate ${new Date().toISOString().substring(0, 10)}`
  page.drawText(labelText, {
    x: plateWidthPt - 60 * MM_TO_PT,
    y: 8 * MM_TO_PT,
    size: 8,
    color: rgb(0.5, 0.5, 0.5),
  })

  return await platePdf.save()
}

function buildDXF(
  params: ImpositionParams,
  layout: ReturnType<typeof calculateLayout>,
  itemsOnThisPlate: number,
): string {
  const lines: string[] = []

  lines.push('0', 'SECTION', '2', 'HEADER')
  lines.push('9', '$INSUNITS', '70', '4')
  lines.push('0', 'ENDSEC')

  lines.push('0', 'SECTION', '2', 'TABLES')
  lines.push('0', 'TABLE', '2', 'LAYER', '70', '2')
  lines.push('0', 'LAYER', '2', 'CUT', '70', '0', '62', '1', '6', 'CONTINUOUS')
  lines.push('0', 'LAYER', '2', 'REG', '70', '0', '62', '7', '6', 'CONTINUOUS')
  lines.push('0', 'ENDTAB', '0', 'ENDSEC')

  lines.push('0', 'SECTION', '2', 'ENTITIES')

  const marks = getRegistrationMarks(params.plate_width_mm, params.plate_height_mm)
  for (const mark of marks) {
    const cx = mark.x * MM_TO_CM
    const cy = mark.y * MM_TO_CM
    const radius = (REG_MARK_SIZE_MM / 2) * MM_TO_CM
    lines.push('0', 'CIRCLE', '8', 'REG', '10', cx.toFixed(4), '20', cy.toFixed(4), '30', '0.0', '40', radius.toFixed(4))
  }

  const positionsToUse = layout.positions.slice(0, itemsOnThisPlate)

  for (const pos of positionsToUse) {
    const centerX_mm = pos.x + params.product_width_mm / 2
    const centerY_mm = pos.y + params.product_height_mm / 2
    const cx = centerX_mm * MM_TO_CM
    const cy = centerY_mm * MM_TO_CM

    if (params.shape === 'circle') {
      const radius = (params.product_width_mm / 2) * MM_TO_CM
      lines.push('0', 'CIRCLE', '8', 'CUT', '10', cx.toFixed(4), '20', cy.toFixed(4), '30', '0.0', '40', radius.toFixed(4))
    } else {
      const halfW = (params.product_width_mm / 2) * MM_TO_CM
      const halfH = (params.product_height_mm / 2) * MM_TO_CM
      const corners = [
        { x: cx - halfW, y: cy - halfH },
        { x: cx + halfW, y: cy - halfH },
        { x: cx + halfW, y: cy + halfH },
        { x: cx - halfW, y: cy + halfH },
      ]
      for (let i = 0; i < 4; i++) {
        const p1 = corners[i]
        const p2 = corners[(i + 1) % 4]
        lines.push(
          '0', 'LINE', '8', 'CUT',
          '10', p1.x.toFixed(4), '20', p1.y.toFixed(4), '30', '0.0',
          '11', p2.x.toFixed(4), '21', p2.y.toFixed(4), '31', '0.0',
        )
      }
    }
  }

  lines.push('0', 'ENDSEC', '0', 'EOF')
  return lines.join('\n')
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  let requestId: string | undefined

  try {
    const body = await req.json()
    requestId = body.request_id as string | undefined

    if (!requestId) {
      return new Response(JSON.stringify({ error: 'request_id required' }), {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      })
    }

    const supabaseUrl = Deno.env.get('SUPABASE_URL')!
    const supabaseServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!
    const supabase = createClient(supabaseUrl, supabaseServiceKey)

    const { data: request, error: reqError } = await supabase
      .from('design_requests')
      .select('*')
      .eq('id', requestId)
      .single()

    if (reqError || !request) {
      throw new Error('Request not found')
    }

    if (request.request_type !== 'imposition') {
      throw new Error('Wrong request type')
    }

    await supabase
      .from('design_requests')
      .update({ status: 'processing', started_at: new Date().toISOString() })
      .eq('id', requestId)

    const params: ImpositionParams = {
      shape: request.parameters?.shape || 'circle',
      product_width_mm: Number(request.parameters?.product_width_mm) || 50,
      product_height_mm: Number(request.parameters?.product_height_mm) || 50,
      quantity: Number(request.parameters?.quantity) || 1,
      plate_width_mm: Number(request.parameters?.plate_width_mm) || 594,
      plate_height_mm: Number(request.parameters?.plate_height_mm) || 420,
      bleed_mm: Number(request.parameters?.bleed_mm) || 3,
      gutter_mm: Number(request.parameters?.gutter_mm) || 5,
    }

    const inputFiles = request.input_files as { path?: string }[] | null
    const inputFile = inputFiles?.[0]
    if (!inputFile?.path) {
      throw new Error('No input file')
    }

    const { data: fileData, error: dlError } = await supabase.storage
      .from(STORAGE_BUCKET)
      .download(inputFile.path)

    if (dlError || !fileData) {
      throw new Error('Failed to download input file: ' + (dlError?.message || 'unknown'))
    }

    const designBytes = new Uint8Array(await fileData.arrayBuffer())
    const layout = calculateLayout(params)

    if (layout.perPlate === 0) {
      throw new Error('No items fit on plate with current dimensions')
    }

    const log = {
      timestamp: new Date().toISOString(),
      message: `Layout: ${layout.cols}×${layout.rows} = ${layout.perPlate} per plate, ${layout.totalPlates} plates total`,
      layout,
    }

    const itemsOnThisPlate = Math.min(layout.perPlate, params.quantity)
    const pdfBytes = await buildPlatePDF(designBytes, params, layout, itemsOnThisPlate)
    const dxfText = buildDXF(params, layout, itemsOnThisPlate)
    const dxfBytes = new TextEncoder().encode(dxfText)

    const timestamp = Date.now()
    const folder = `design_requests/${requestId}/output`
    const pdfPath = `${folder}/plate_${timestamp}.pdf`
    const dxfPath = `${folder}/cut_${timestamp}.dxf`

    const { error: pdfUpError } = await supabase.storage
      .from(STORAGE_BUCKET)
      .upload(pdfPath, pdfBytes, { contentType: 'application/pdf', upsert: true })

    if (pdfUpError) throw new Error('PDF upload failed: ' + pdfUpError.message)

    const { error: dxfUpError } = await supabase.storage
      .from(STORAGE_BUCKET)
      .upload(dxfPath, dxfBytes, { contentType: 'application/dxf', upsert: true })

    if (dxfUpError) throw new Error('DXF upload failed: ' + dxfUpError.message)

    const outputFiles = [
      {
        path: pdfPath,
        name: `plate_${timestamp}.pdf`,
        size: pdfBytes.length,
        type: 'application/pdf',
        uploaded_at: new Date().toISOString(),
      },
      {
        path: dxfPath,
        name: `cut_${timestamp}.dxf`,
        size: dxfBytes.length,
        type: 'application/dxf',
        uploaded_at: new Date().toISOString(),
      },
    ]

    const existingLog = Array.isArray(request.processing_log) ? request.processing_log : []

    await supabase
      .from('design_requests')
      .update({
        status: 'completed',
        completed_at: new Date().toISOString(),
        output_files: outputFiles,
        processing_log: [
          ...existingLog,
          log,
          {
            timestamp: new Date().toISOString(),
            message: `Generated PDF (${pdfBytes.length} bytes) + DXF (${dxfBytes.length} bytes)`,
          },
        ],
        error_message: null,
      })
      .eq('id', requestId)

    return new Response(
      JSON.stringify({ success: true, layout, output_files: outputFiles }),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } },
    )
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e)
    console.error('Imposition error:', msg)

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

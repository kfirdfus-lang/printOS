const ExcelJS = require('exceljs');

const SUPABASE_URL = "https://pvwcpukfhyrmdpxgfwrk.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InB2d2NwdWtmaHlybWRweGdmd3JrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY4NzIzNzMsImV4cCI6MjA5MjQ0ODM3M30.69GpTOQdwtAgouGN0MIHqQc6bImxAbfQ72xL2S4KaoM";

module.exports = async (req, res) => {
  try {
    // Fetch debt data from Supabase
    const response = await fetch(`${SUPABASE_URL}/rest/v1/debt?order=days_overdue.desc`, {
      headers: {
        apikey: SUPABASE_ANON,
        Authorization: `Bearer ${SUPABASE_ANON}`,
      }
    });
    const debt = await response.json();

    const overdue = debt.filter(d => d.status === 'באיחור').sort((a,b) => b.days_overdue - a.days_overdue);
    const current = debt.filter(d => d.status !== 'באיחור').sort((a,b) => b.amount - a.amount);
    const all = [...overdue, ...current];
    const total = all.reduce((s,d) => s + Number(d.amount), 0);
    const totalOv = overdue.reduce((s,d) => s + Number(d.amount), 0);
    const totalCu = current.reduce((s,d) => s + Number(d.amount), 0);
    const now = new Date().toLocaleDateString('he-IL');

    const wb = new ExcelJS.Workbook();
    wb.creator = 'PrintOS';

    // ── Helper styles ──
    const hdrFill = (hex) => ({ type:'pattern', pattern:'solid', fgColor:{argb:'FF'+hex} });
    const hdrFont = (hex='FFFFFF', sz=10) => ({ name:'Arial', bold:true, size:sz, color:{argb:'FF'+hex} });
    const datFont = (hex='1E293B', bold=false, sz=10) => ({ name:'Arial', bold, size:sz, color:{argb:'FF'+hex} });
    const datFill = (hex) => ({ type:'pattern', pattern:'solid', fgColor:{argb:'FF'+hex} });
    const border = (hex='E2E8F0') => ({
      top:{style:'thin',color:{argb:'FF'+hex}},
      bottom:{style:'thin',color:{argb:'FF'+hex}},
      left:{style:'thin',color:{argb:'FF'+hex}},
      right:{style:'thin',color:{argb:'FF'+hex}},
    });
    const medBorder = (pos, hex) => ({ [pos]:{style:'medium',color:{argb:'FF'+hex}} });

    // ════════════════════════════
    // SHEET 1: סיכום
    // ════════════════════════════
    const ws1 = wb.addWorksheet('סיכום', { views:[{rightToLeft:true}] });
    ws1.properties.tabColor = {argb:'FF1E3A5F'};

    ws1.columns = [
      {key:'a', width:3},
      {key:'b', width:34},
      {key:'c', width:18},
      {key:'d', width:16},
      {key:'e', width:14},
      {key:'f', width:16},
      {key:'g', width:3},
    ];

    // Title
    ws1.mergeCells('B1:F1'); ws1.getRow(1).height = 14;
    ws1.mergeCells('B2:F2'); ws1.getRow(2).height = 36;
    ws1.mergeCells('B3:F3'); ws1.getRow(3).height = 18;
    ws1.getRow(4).height = 8;

    for(let c=1;c<=7;c++){
      ['1','2','3'].forEach(r => {
        ws1.getCell(r, c).fill = hdrFill('0F172A');
      });
    }

    const t1 = ws1.getCell('B2');
    t1.value = 'דוח גבייה — בית דפוס';
    t1.font = hdrFont('F8FAFC', 20);
    t1.fill = hdrFill('0F172A');
    t1.alignment = {horizontal:'right', vertical:'middle', readingOrder:'rtl'};

    const t2 = ws1.getCell('B3');
    t2.value = `נוצר: ${now}  |  ${all.length} לקוחות חייבים`;
    t2.font = datFont('64748B', false, 10);
    t2.fill = hdrFill('0F172A');
    t2.alignment = {horizontal:'right', vertical:'middle'};

    // KPI Cards
    ws1.getRow(5).height = 50; ws1.getRow(6).height = 36; ws1.getRow(7).height = 8;
    const kpis = [
      ['B','C','סה״כ חוב פתוח', `₪${total.toLocaleString()}`, '1E40AF','DBEAFE','2563EB'],
      ['D','E','באיחור 🚨', `₪${totalOv.toLocaleString()}`, '991B1B','FEE2E2','DC2626'],
      ['F','G','שוטף ✅', `₪${totalCu.toLocaleString()}`, '14532D','DCFCE7','16A34A'],
    ];

    for(const [c1,c2,title,value,dark,light,mid] of kpis){
      ws1.mergeCells(`${c1}5:${c2}5`);
      ws1.mergeCells(`${c1}6:${c2}6`);
      const tc = ws1.getCell(`${c1}5`);
      tc.value = title;
      tc.font = hdrFont(dark, 11);
      tc.fill = hdrFill(light);
      tc.alignment = {horizontal:'center', vertical:'bottom'};
      tc.border = { top:{style:'medium',color:{argb:'FF'+mid}}, left:{style:'thin',color:{argb:'FFE2E8F0'}}, right:{style:'thin',color:{argb:'FFE2E8F0'}} };

      const vc = ws1.getCell(`${c1}6`);
      vc.value = value;
      vc.font = hdrFont(dark, 18);
      vc.fill = hdrFill(light);
      vc.alignment = {horizontal:'center', vertical:'middle'};
      vc.border = { bottom:{style:'medium',color:{argb:'FF'+mid}}, left:{style:'thin',color:{argb:'FFE2E8F0'}}, right:{style:'thin',color:{argb:'FFE2E8F0'}} };
    }

    // Section title
    ws1.getRow(8).height = 30;
    ws1.mergeCells('B8:F8');
    const st = ws1.getCell('B8');
    st.value = 'פירוט חייבים';
    st.font = datFont('1E293B', true, 13);
    st.alignment = {horizontal:'right', vertical:'middle'};
    st.border = { bottom:{style:'medium',color:{argb:'FF2563EB'}} };

    // Table header
    ws1.getRow(9).height = 28;
    const headers = ['לקוח','סכום חוב (₪)','תאריך פירעון','ימי איחור','דחיפות'];
    ['B','C','D','E','F'].forEach((col, i) => {
      const cell = ws1.getCell(`${col}9`);
      cell.value = headers[i];
      cell.font = hdrFont('FFFFFF', 10);
      cell.fill = hdrFill('1E3A5F');
      cell.alignment = {horizontal:'center', vertical:'middle'};
      cell.border = border('1E3A5F');
    });

    // Data rows
    all.forEach((d, i) => {
      const r = i + 10;
      ws1.getRow(r).height = 22;
      const isOv = d.status === 'באיחור';
      const days = d.days_overdue;
      const bg = isOv ? (i%2===0?'FEF2F2':'FFF7F7') : (i%2===0?'F8FAFC':'FFFFFF');
      const urg = isOv ? (days>90?'🔴 דחוף מאוד':days>30?'🟠 דחוף':'🟡 עקוב') : '🟢 שוטף';
      const urgColor = isOv ? (days>90?'991B1B':days>30?'C2410C':'854D0E') : '14532D';

      const bCell = ws1.getCell(`B${r}`);
      bCell.value = d.client_name; bCell.font = datFont('1E293B',true); bCell.fill = datFill(bg);
      bCell.alignment = {horizontal:'right',vertical:'middle'}; bCell.border = border();

      const cCell = ws1.getCell(`C${r}`);
      cCell.value = Number(d.amount); cCell.numFmt = '#,##0';
      cCell.font = datFont(isOv?'DC2626':'16A34A',true); cCell.fill = datFill(bg);
      cCell.alignment = {horizontal:'center',vertical:'middle'}; cCell.border = border();

      const dCell = ws1.getCell(`D${r}`);
      dCell.value = d.due_date; dCell.font = datFont('475569'); dCell.fill = datFill(bg);
      dCell.alignment = {horizontal:'center',vertical:'middle'}; dCell.border = border();

      const eCell = ws1.getCell(`E${r}`);
      eCell.value = isOv ? days : '—'; eCell.font = datFont(isOv?'DC2626':'475569',isOv);
      eCell.fill = datFill(bg); eCell.alignment = {horizontal:'center',vertical:'middle'}; eCell.border = border();

      const fCell = ws1.getCell(`F${r}`);
      fCell.value = urg; fCell.font = datFont(urgColor); fCell.fill = datFill(bg);
      fCell.alignment = {horizontal:'center',vertical:'middle'}; fCell.border = border();
    });

    // Total
    const tr = all.length + 10;
    ws1.getRow(tr).height = 26;
    ['B','C','D','E','F'].forEach((col, i) => {
      const cell = ws1.getCell(`${col}${tr}`);
      cell.value = i===0 ? 'סה״כ' : i===1 ? total : '';
      if(i===1) cell.numFmt = '#,##0';
      cell.font = hdrFont('FFFFFF', 11);
      cell.fill = hdrFill('1E3A5F');
      cell.alignment = {horizontal:'center',vertical:'middle'};
      cell.border = border('1E3A5F');
    });

    // ════════════════════════════
    // SHEET 2: חייבים באיחור
    // ════════════════════════════
    const ws2 = wb.addWorksheet('חייבים באיחור', { views:[{rightToLeft:true}] });
    ws2.properties.tabColor = {argb:'FFDC2626'};
    ws2.columns = [{key:'a',width:3},{key:'b',width:34},{key:'c',width:16},{key:'d',width:16},{key:'e',width:14},{key:'f',width:22},{key:'g',width:3}];

    for(let c=1;c<=7;c++) ['1','2','3'].forEach(r => { ws2.getCell(r,c).fill = hdrFill('450A0A'); });
    ws2.getRow(1).height=14; ws2.getRow(2).height=38; ws2.getRow(3).height=14;

    ws2.mergeCells('B2:F2');
    const h2 = ws2.getCell('B2');
    h2.value = `חייבים באיחור — ${overdue.length} לקוחות | ₪${totalOv.toLocaleString()}`;
    h2.font = hdrFont('FECACA', 15); h2.fill = hdrFill('450A0A');
    h2.alignment = {horizontal:'right',vertical:'middle'};

    ws2.getRow(4).height = 28;
    ['לקוח','סכום (₪)','תאריך פירעון','ימי איחור','הערות טיפול'].forEach((h,i) => {
      const cell = ws2.getCell(4, i+2);
      cell.value = h; cell.font = hdrFont('FFFFFF',10); cell.fill = hdrFill('DC2626');
      cell.alignment = {horizontal:'center',vertical:'middle'}; cell.border = border('DC2626');
    });

    overdue.forEach((d, i) => {
      const r = i + 5;
      ws2.getRow(r).height = 24;
      const days = d.days_overdue;
      const bg = i%2===0?'FEF2F2':'FFFFFF';
      const urg = days>90?'🔴 דחוף מאוד':days>30?'🟠 דחוף':'🟡 עקוב';
      const notes = d.payment_notes ? `  |  ${d.payment_notes}` : '';

      const b = ws2.getCell(`B${r}`); b.value=d.client_name; b.font=datFont('1E293B',true); b.fill=datFill(bg); b.alignment={horizontal:'right',vertical:'middle'}; b.border=border();
      const c = ws2.getCell(`C${r}`); c.value=Number(d.amount); c.numFmt='#,##0'; c.font=datFont('DC2626',true); c.fill=datFill(bg); c.alignment={horizontal:'center',vertical:'middle'}; c.border=border();
      const dd = ws2.getCell(`D${r}`); dd.value=d.due_date; dd.font=datFont('475569'); dd.fill=datFill(bg); dd.alignment={horizontal:'center',vertical:'middle'}; dd.border=border();
      const e = ws2.getCell(`E${r}`); e.value=days; e.font=datFont('DC2626',true); e.fill=datFill(bg); e.alignment={horizontal:'center',vertical:'middle'}; e.border=border();
      const f = ws2.getCell(`F${r}`); f.value=urg+notes; f.font=datFont(days>90?'991B1B':days>30?'C2410C':'854D0E'); f.fill=datFill(bg); f.alignment={horizontal:'right',vertical:'middle',wrapText:true}; f.border=border();
    });

    const tr2 = overdue.length + 5;
    ws2.getRow(tr2).height = 26;
    ['B','C','D','E','F'].forEach((col,i) => {
      const cell = ws2.getCell(`${col}${tr2}`);
      cell.value = i===0?'סה״כ':i===1?totalOv:'';
      if(i===1) cell.numFmt='#,##0';
      cell.font=hdrFont('FFFFFF',11); cell.fill=hdrFill('DC2626');
      cell.alignment={horizontal:'center',vertical:'middle'}; cell.border=border('DC2626');
    });

    // ════════════════════════════
    // SHEET 3: שוטף
    // ════════════════════════════
    const ws3 = wb.addWorksheet('שוטף', { views:[{rightToLeft:true}] });
    ws3.properties.tabColor = {argb:'FF16A34A'};
    ws3.columns = [{key:'a',width:3},{key:'b',width:34},{key:'c',width:16},{key:'d',width:16},{key:'e',width:3}];

    for(let c=1;c<=5;c++) ['1','2','3'].forEach(r => { ws3.getCell(r,c).fill = hdrFill('052E16'); });
    ws3.getRow(1).height=14; ws3.getRow(2).height=38; ws3.getRow(3).height=14;

    ws3.mergeCells('B2:D2');
    const h3 = ws3.getCell('B2');
    h3.value = `חוב שוטף — ${current.length} לקוחות | ₪${totalCu.toLocaleString()}`;
    h3.font = hdrFont('BBF7D0',15); h3.fill = hdrFill('052E16');
    h3.alignment = {horizontal:'right',vertical:'middle'};

    ws3.getRow(4).height = 28;
    ['לקוח','סכום (₪)','תאריך פירעון'].forEach((h,i) => {
      const cell = ws3.getCell(4, i+2);
      cell.value=h; cell.font=hdrFont('FFFFFF',10); cell.fill=hdrFill('16A34A');
      cell.alignment={horizontal:'center',vertical:'middle'}; cell.border=border('16A34A');
    });

    current.forEach((d,i) => {
      const r = i+5;
      ws3.getRow(r).height = 22;
      const bg = i%2===0?'F0FDF4':'FFFFFF';
      const b=ws3.getCell(`B${r}`); b.value=d.client_name; b.font=datFont('1E293B',true); b.fill=datFill(bg); b.alignment={horizontal:'right',vertical:'middle'}; b.border=border();
      const c=ws3.getCell(`C${r}`); c.value=Number(d.amount); c.numFmt='#,##0'; c.font=datFont('16A34A',true); c.fill=datFill(bg); c.alignment={horizontal:'center',vertical:'middle'}; c.border=border();
      const dd=ws3.getCell(`D${r}`); dd.value=d.due_date; dd.font=datFont('475569'); dd.fill=datFill(bg); dd.alignment={horizontal:'center',vertical:'middle'}; dd.border=border();
    });

    const tr3 = current.length+5;
    ws3.getRow(tr3).height=26;
    ['B','C','D'].forEach((col,i) => {
      const cell=ws3.getCell(`${col}${tr3}`);
      cell.value=i===0?'סה״כ':i===1?totalCu:'';
      if(i===1) cell.numFmt='#,##0';
      cell.font=hdrFont('FFFFFF',11); cell.fill=hdrFill('16A34A');
      cell.alignment={horizontal:'center',vertical:'middle'}; cell.border=border('16A34A');
    });

    // Write and send
    const buffer = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="debt_report.xlsx"');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.send(Buffer.from(buffer));

  } catch(e) {
    console.error(e);
    res.status(500).json({error: e.message});
  }
};

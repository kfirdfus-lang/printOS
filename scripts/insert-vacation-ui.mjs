import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, "..");
const indexPath = path.join(root, "index.html");

const snippet = `
function VacationFormPage({slug}){
  const [step,setStep]=useState("identify");
  const [loading,setLoading]=useState(true);
  const [error,setError]=useState(null);
  const [employeeName,setEmployeeName]=useState("");
  const [idLast4,setIdLast4]=useState("");
  const [history,setHistory]=useState([]);
  const [startDate,setStartDate]=useState("");
  const [endDate,setEndDate]=useState("");
  const [reasonType,setReasonType]=useState("vacation");
  const [reasonNotes,setReasonNotes]=useState("");
  const [submitting,setSubmitting]=useState(false);
  const canvasRef=useRef(null);
  const padRef=useRef(null);
  useEffect(()=>{
    let cancelled=false;
    sb(\`vacation_form_config?form_slug=eq.\${encodeURIComponent(slug)}&is_active=eq.true&limit=1\`).then(r=>{
      if(cancelled)return;
      if(!r?.length){setError("הטופס לא נמצא");setLoading(false);return;}
      setLoading(false);
    }).catch(()=>{if(!cancelled){setError("שגיאה בטעינה");setLoading(false);}});
    return()=>{cancelled=true};
  },[slug]);
  const initPad=()=>{
    if(!canvasRef.current||!window.SignaturePad)return;
    const canvas=canvasRef.current;
    const rect=canvas.parentElement?.getBoundingClientRect();
    canvas.width=rect?Math.max(300,rect.width-4):550;
    canvas.height=180;
    padRef.current=new window.SignaturePad(canvas,{backgroundColor:"rgb(255,255,255)",penColor:"rgb(0,0,0)"});
  };
  useEffect(()=>{
    if(step!=="form")return;
    ensureSignaturePadLoaded(initPad);
    return()=>{padRef.current=null;};
  },[step]);
  const handleIdentify=async()=>{
    if(!employeeName.trim()){alert("יש למלא שם מלא");return;}
    if(!/^\\d{4}$/.test(idLast4)){alert("יש להזין 4 ספרות אחרונות של ת.ז");return;}
    try{
      const r=await sb(\`vacation_requests?employee_name=eq.\${encodeURIComponent(employeeName.trim())}&employee_id_last4=eq.\${idLast4}&order=submitted_at.desc&limit=10\`);
      setHistory(r||[]);
    }catch{setHistory([]);}
    setStep("form");
  };
  const handleSubmit=async()=>{
    if(!startDate||!endDate){alert("יש למלא תאריכי התחלה וסיום");return;}
    if(new Date(startDate)>new Date(endDate)){alert("תאריך התחלה חייב להיות לפני תאריך סיום");return;}
    if(reasonType==="other"&&!reasonNotes.trim()){alert('יש למלא פירוט לסיבה "אחר"');return;}
    if(!padRef.current||padRef.current.isEmpty()){alert("יש לחתום בטופס");return;}
    setSubmitting(true);
    try{
      const signature=padRef.current.toDataURL("image/png");
      const days=calcVacationDays(startDate,endDate);
      await sb("vacation_requests",{method:"POST",body:JSON.stringify({
        employee_name:employeeName.trim(),employee_id_last4:idLast4,
        start_date:startDate,end_date:endDate,total_days:days,
        reason_type:reasonType,reason_notes:reasonNotes.trim()||null,
        signature_data:signature,status:"pending",
        user_agent:typeof navigator!=="undefined"?navigator.userAgent:null
      })});
      try{
        await sendEmail({
          to:ADMIN_NOTIFY_EMAIL,
          subject:\`📅 בקשת חופש חדשה - \${employeeName.trim()} (\${VACATION_REASON_EMAIL[reasonType]})\`,
          from:"דפוס נטלי <orders@natalie-print.com>",
          html:buildVacationManagerEmail(employeeName.trim(),idLast4,startDate,endDate,days,VACATION_REASON_EMAIL[reasonType],reasonNotes.trim(),signature)
        });
      }catch(e){console.warn("vacation email failed",e);}
      setStep("submitted");
    }catch(e){alert("שגיאה בשליחה: "+(e.message||String(e)));}
    finally{setSubmitting(false);}
  };
  if(loading)return<div style={{textAlign:"center",padding:50,direction:"rtl"}}>טוען...</motion.div>;
  if(error)return<div style={{maxWidth:500,margin:"50px auto",padding:30,textAlign:"center",direction:"rtl"}}><h2 style={{color:"#991b1b"}}>{error}</h2></motion.div>;
  const dayPreview=startDate&&endDate?calcVacationDays(startDate,endDate):0;
  return<div style={{minHeight:"100vh",background:"#f8fafc",padding:"20px 10px",direction:"rtl",fontFamily:"Segoe UI,Arial,sans-serif"}}>
    <div style={{maxWidth:600,margin:"0 auto",background:"white",borderRadius:12,overflow:"hidden",boxShadow:"0 4px 6px rgba(0,0,0,0.1)"}}>
      <motion.div style={{padding:25,textAlign:"center"}}><motion.div style={{fontSize:32,fontWeight:700,color:"#0d9488",letterSpacing:2}}>NATALIE</motion.div><motion.div style={{color:"#134e4a",marginTop:4}}>פתרונות הדפסה</motion.div></motion.div>
      <motion.div style={{height:4,background:"linear-gradient(90deg,#2dd4bf,#0d9488)"}}/>
      <motion.div style={{padding:30}}>
        <h1 style={{margin:"0 0 8px",color:"#0d9488",fontSize:24}}>📅 בקשת חופש / היעדרות</h1>
        {step==="identify"&&<>
          <p style={{color:"#64748b",marginBottom:20}}>אנא הזן את פרטיך כדי להתחיל</p>
          <SignupFormField label="שם מלא *" value={employeeName} onChange={setEmployeeName} placeholder="שם פרטי ושם משפחה"/>
          <label style={{display:"block",marginBottom:4,fontWeight:600,color:"#134e4a"}}>4 ספרות אחרונות של ת.ז *</label>
          <input value={idLast4} onChange={e=>setIdLast4(e.target.value.replace(/\\D/g,"").slice(0,4))} maxLength={4} placeholder="1234" style={{width:"100%",padding:12,borderRadius:6,border:"1px solid #d1d5db",direction:"ltr",fontSize:18,textAlign:"center",letterSpacing:8,fontFamily:"monospace",marginBottom:16}}/>
          <button type="button" onClick={handleIdentify} style={{width:"100%",padding:14,background:"#0d9488",color:"white",border:"none",borderRadius:8,fontSize:16,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>המשך ➜</button>
        </>}
        {step==="form"&&<>
          <motion.div style={{background:"#f0fdfa",padding:12,borderRadius:8,marginBottom:20,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <motion.div><motion.div style={{fontWeight:600,color:"#0d9488"}}>{employeeName}</motion.div><motion.div style={{fontSize:12,color:"#64748b",direction:"ltr"}}>ת.ז: ••••{idLast4}</motion.div></motion.div>
            <button type="button" onClick={()=>{setStep("identify");padRef.current=null;}} style={{background:"none",border:"none",color:"#0d9488",cursor:"pointer",fontFamily:"inherit"}}>שינוי</button>
          </motion.div>
          {history.length>0&&<details style={{marginBottom:20,background:"#fafafa",padding:12,borderRadius:8}}><summary style={{cursor:"pointer",fontWeight:600,color:"#0d9488"}}>📋 הבקשות הקודמות שלי ({history.length})</summary>
            <motion.div style={{marginTop:10,display:"grid",gap:6}}>{history.map(h=><motion.div key={h.id} style={{padding:10,background:"white",borderRadius:6,fontSize:13,border:"1px solid #e5e7eb"}}>
              <motion.div style={{display:"flex",justifyContent:"space-between"}}><span>{new Date(h.start_date).toLocaleDateString("he-IL")} - {new Date(h.end_date).toLocaleDateString("he-IL")}</span><span style={{fontSize:11}}>{h.status==="approved"?"✅":h.status==="rejected"?"❌":"⏳"}</span></motion.div>
              <motion.div style={{color:"#6b7280",marginTop:4}}>{h.total_days} ימים · {VACATION_REASON_EMAIL[h.reason_type]}</motion.div>
            </motion.div>)}</motion.div>
          </details>}
          <motion.div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:15}}>
            <motion.div><label style={{fontWeight:600,fontSize:14}}>מתאריך *</label><input type="date" value={startDate} onChange={e=>setStartDate(e.target.value)} style={{width:"100%",padding:10,borderRadius:6,border:"1px solid #d1d5db"}}/></motion.div>
            <motion.div><label style={{fontWeight:600,fontSize:14}}>עד תאריך *</label><input type="date" value={endDate} onChange={e=>setEndDate(e.target.value)} min={startDate} style={{width:"100%",padding:10,borderRadius:6,border:"1px solid #d1d5db"}}/></motion.div>
          </motion.div>
          {dayPreview>0&&<motion.div style={{marginBottom:15,padding:8,background:"#f0fdfa",borderRadius:6,color:"#0d9488",textAlign:"center"}}>📊 סה״כ: {dayPreview} ימים</motion.div>}
          <label style={{display:"block",marginBottom:8,fontWeight:600}}>סיבת ההיעדרות *</label>
          <motion.div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:15}}>
            {[{value:"vacation",label:"🏖️ חופשה"},{value:"sick",label:"🤒 מחלה"},{value:"reserve",label:"⚔️ מילואים"},{value:"other",label:"📋 אחר"}].map(opt=>
              <button key={opt.value} type="button" onClick={()=>setReasonType(opt.value)} style={{padding:10,background:reasonType===opt.value?"#0d9488":"white",color:reasonType===opt.value?"white":"#134e4a",border:\`2px solid \${reasonType===opt.value?"#0d9488":"#d1d5db"}\`,borderRadius:6,cursor:"pointer",fontWeight:600,fontFamily:"inherit"}}>{opt.label}</button>
            )}
          </motion.div>
          <SignupFormField label={reasonType==="other"?"פירוט נוסף *":"פירוט נוסף"} value={reasonNotes} onChange={setReasonNotes} textarea placeholder="הערות..."/>
          <label style={{display:"block",marginBottom:8,fontWeight:600}}>✍️ חתימה דיגיטלית *</label>
          <motion.div style={{border:"2px solid #d1d5db",borderRadius:8,overflow:"hidden",marginBottom:8}}><canvas ref={canvasRef} style={{display:"block",width:"100%",height:180,touchAction:"none"}}/></motion.div>
          <button type="button" onClick={()=>padRef.current?.clear()} style={{marginBottom:16,padding:"6px 12px",background:"#f3f4f6",border:"1px solid #d1d5db",borderRadius:6,cursor:"pointer",fontFamily:"inherit"}}>🗑️ נקה חתימה</button>
          <button type="button" onClick={handleSubmit} disabled={submitting} style={{width:"100%",padding:16,background:"#0d9488",color:"white",border:"none",borderRadius:8,fontSize:16,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>{submitting?"⏳ שולח...":"✅ שלח בקשה"}</button>
        </>}
        {step==="submitted"&&<motion.div style={{textAlign:"center",padding:"20px 0"}}>
          <motion.div style={{fontSize:64,marginBottom:15}}>✅</motion.div>
          <h2 style={{color:"#0d9488"}}>הבקשה נשלחה!</h2>
          <p style={{color:"#64748b"}}>ההנהלה תקבל את הבקשה ותחזור אליך בהקדם.</p>
          <button type="button" onClick={()=>{setStep("identify");setEmployeeName("");setIdLast4("");setStartDate("");setEndDate("");setReasonNotes("");setReasonType("vacation");padRef.current=null;}} style={{marginTop:20,padding:12,background:"#f3f4f6",border:"1px solid #d1d5db",borderRadius:6,cursor:"pointer",fontFamily:"inherit"}}>📝 הגש בקשה נוספת</button>
        </motion.div>}
      </motion.div>
      <motion.div style={{background:"#f8fafc",padding:15,textAlign:"center",fontSize:12,color:"#64748b"}}>natalie-print.com · 03-6815703</motion.div>
    </motion.div>
  </motion.div>;
}

function VacationReviewModal({request,user,onClose,onUpdate}){
  const [reviewNotes,setReviewNotes]=useState(request.reviewer_notes||"");
  const handleApprove=async()=>{
    await sb(\`vacation_requests?id=eq.\${request.id}\`,{method:"PATCH",body:JSON.stringify({status:"approved",reviewed_at:new Date().toISOString(),reviewed_by:user?.name||"unknown",reviewer_notes:reviewNotes.trim()||null})});
    onUpdate?.();onClose();
  };
  const handleReject=async()=>{
    if(!reviewNotes.trim()){alert("יש לכתוב סיבת דחייה בהערות");return;}
    if(!confirm("לדחות את הבקשה?"))return;
    await sb(\`vacation_requests?id=eq.\${request.id}\`,{method:"PATCH",body:JSON.stringify({status:"rejected",reviewed_at:new Date().toISOString(),reviewed_by:user?.name||"unknown",reviewer_notes:reviewNotes.trim()})});
    onUpdate?.();onClose();
  };
  return<div onClick={onClose} style={SIGNUP_MODAL_OVERLAY}>
    <motion.div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:16,padding:24,width:"100%",maxWidth:600,maxHeight:"90vh",overflowY:"auto",direction:"rtl"}}>
      <h2 style={{color:"#0d9488"}}>📋 בקשת חופש — {request.employee_name}</h2>
      <SignupDetailRow label="סוג" value={VACATION_REASON_LABELS[request.reason_type]}/>
      <SignupDetailRow label="תאריכים" value={\`\${new Date(request.start_date).toLocaleDateString("he-IL")} - \${new Date(request.end_date).toLocaleDateString("he-IL")} (\${request.total_days} ימים)\`}/>
      {request.reason_notes&&<SignupDetailRow label="פירוט" value={request.reason_notes}/>}
      <img src={request.signature_data} alt="חתימה" style={{maxHeight:120,maxWidth:"100%",display:"block",margin:"12px auto"}}/>
      {request.status==="pending"?<>
        <textarea value={reviewNotes} onChange={e=>setReviewNotes(e.target.value)} placeholder="הערות מנהל..." style={{width:"100%",minHeight:60,padding:8,borderRadius:6,border:"1px solid #d1d5db",fontFamily:"inherit",margin:"12px 0"}}/>
        <motion.div style={{display:"flex",gap:10}}>
          <button type="button" onClick={handleReject} style={{flex:1,padding:12,background:"#ef4444",color:"white",border:"none",borderRadius:6,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>❌ דחה</button>
          <button type="button" onClick={handleApprove} style={{flex:2,padding:12,background:"#10b981",color:"white",border:"none",borderRadius:6,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>✅ אשר</button>
        </motion.div>
      </>:<motion.div style={{padding:12,background:request.status==="approved"?"#d1fae5":"#fee2e2",borderRadius:6,marginTop:12}}>{request.status==="approved"?"✅ אושר":"❌ נדחה"} {request.reviewer_notes&&" — "+request.reviewer_notes}</motion.div>}
      <button type="button" onClick={onClose} style={{width:"100%",padding:10,marginTop:10,background:"#f3f4f6",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit"}}>סגור</button>
    </motion.div>
  </motion.div>;
}

function VacationRequestCard({request,user,onUpdate}){
  const [showDetails,setShowDetails]=useState(false);
  const sc={pending:"#f59e0b",approved:"#10b981",rejected:"#ef4444"};
  const sl={pending:"⏳ ממתין",approved:"✅ אושר",rejected:"❌ נדחה"};
  return<>
    <motion.div onClick={()=>setShowDetails(true)} style={{background:"white",padding:15,borderRadius:8,borderRight:\`4px solid \${sc[request.status]}\`,cursor:"pointer"}}>
      <motion.div style={{fontWeight:600,fontSize:16}}>{request.employee_name} <span style={{fontSize:12,color:"#9ca3af",direction:"ltr"}}>(••••{request.employee_id_last4})</span></motion.div>
      <motion.div style={{fontSize:14,marginTop:6}}>{VACATION_REASON_LABELS[request.reason_type]} · {request.total_days} ימים</motion.div>
      <motion.div style={{fontSize:13,color:"#6b7280",direction:"ltr",textAlign:"right"}}>📅 {new Date(request.start_date).toLocaleDateString("he-IL")} - {new Date(request.end_date).toLocaleDateString("he-IL")}</motion.div>
      <motion.div style={{fontSize:12,color:sc[request.status],fontWeight:600,marginTop:6}}>{sl[request.status]}</motion.div>
    </motion.div>
    {showDetails&&<VacationReviewModal request={request} user={user} onClose={()=>setShowDetails(false)} onUpdate={onUpdate}/>}
  </>;
}

function VacationFormManager({user}){
  const [config,setConfig]=useState(null);
  const [requests,setRequests]=useState([]);
  const [filter,setFilter]=useState("pending");
  const [copied,setCopied]=useState(false);
  const loadData=async()=>{
    try{
      const cfg=await sb("vacation_form_config?id=eq.1&limit=1");
      setConfig(cfg?.[0]||null);
      const reqs=await sb("vacation_requests?order=submitted_at.desc&limit=100");
      setRequests(reqs||[]);
    }catch(e){console.warn(e);}
  };
  useEffect(()=>{loadData();const t=setInterval(loadData,30000);return()=>clearInterval(t);},[]);
  if(!config)return<div>טוען...</motion.div>;
  const formUrl=\`\${window.location.origin}\${window.location.pathname||""}\${window.location.search||""}#vacation-form/\${config.form_slug}\`;
  const filtered=requests.filter(r=>filter==="all"||r.status===filter);
  const pendingN=requests.filter(r=>r.status==="pending").length;
  return<div>
    <h1 style={{fontSize:22,fontWeight:800,margin:"0 0 20px"}}>📅 ניהול חופשות עובדים</h1>
    <motion.div style={{background:"linear-gradient(135deg,#f0fdfa,#fff)",padding:20,borderRadius:10,marginBottom:20,border:"2px solid #2dd4bf"}}>
      <h3 style={{color:"#0d9488",margin:"0 0 10px"}}>🔗 קישור קבוע לטופס חופש</h3>
      <motion.div style={{background:"white",padding:12,borderRadius:6,direction:"ltr",fontSize:13,wordBreak:"break-all",fontFamily:"monospace",marginBottom:10,border:"1px solid #e5e7eb"}}>{formUrl}</motion.div>
      <motion.div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        <button type="button" onClick={async()=>{try{await navigator.clipboard.writeText(formUrl);setCopied(true);setTimeout(()=>setCopied(false),2000);}catch{}}} style={{padding:10,background:copied?"#10b981":"#f3f4f6",border:"1px solid #d1d5db",borderRadius:6,cursor:"pointer",fontFamily:"inherit"}}>{copied?"✅ הועתק":"📋 העתק"}</button>
        <button type="button" onClick={()=>window.open(\`https://wa.me/?text=\${encodeURIComponent("קישור לבקשת חופש בדפוס נטלי:\\n"+formUrl)}\`,"_blank")} style={{padding:10,background:"#25d366",color:"white",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>📱 WhatsApp</button>
      </motion.div>
    </motion.div>
    <motion.div style={{display:"flex",gap:8,marginBottom:15,flexWrap:"wrap"}}>
      <button type="button" onClick={()=>setFilter("pending")} style={vacationTabStyle(filter==="pending")}>⏳ ממתינות ({pendingN})</button>
      <button type="button" onClick={()=>setFilter("approved")} style={vacationTabStyle(filter==="approved")}>✅ מאושרות</button>
      <button type="button" onClick={()=>setFilter("rejected")} style={vacationTabStyle(filter==="rejected")}>❌ נדחו</button>
      <button type="button" onClick={()=>setFilter("all")} style={vacationTabStyle(filter==="all")}>הכל</button>
    </motion.div>
    {filtered.length===0?<motion.div style={{textAlign:"center",padding:40,color:"#9ca3af"}}>אין בקשות</motion.div>:
    <motion.div style={{display:"grid",gap:10}}>{filtered.map(r=><VacationRequestCard key={r.id} request={r} user={user} onUpdate={loadData}/>)}</motion.div>}
  </motion.div>;
}
`;

let html = fs.readFileSync(indexPath, "utf8");
const marker = "\n// ── MAIN APP ──";
if (!html.includes(marker)) {
  console.error("marker not found");
  process.exit(1);
}
let fixed = snippet.replace(/motion\.motion\.div/g, "div");
fixed = fixed.replace(/motion\.div/g, "motion.div");
html = html.replace(marker, fixed + marker);
html = html.replace(/motion\.div/g, "div");
fs.writeFileSync(indexPath, html, "utf8");
console.log("inserted vacation UI");

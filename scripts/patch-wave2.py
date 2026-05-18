# -*- coding: utf-8 -*-
import pathlib
import re

p = pathlib.Path(__file__).resolve().parent.parent / "index.html"
c = p.read_text(encoding="utf-8")

def must_replace(old, new, label):
    global c
    if old not in c:
        raise SystemExit(f"{label}: not found")
    c = c.replace(old, new, 1)

INSERT_MARKER = "const escHtml=(s)=>String(s??\"\").replace"

COMPONENTS = r'''
function TaskInquiryButton({task,user,compact}){
  const [showModal,setShowModal]=useState(false);
  return<>
    <button type="button" onClick={e=>{e.stopPropagation();setShowModal(true);}} title="פנה לאחראי" style={{padding:compact?"3px 6px":"4px 8px",fontSize:compact?10:11,background:"#fff",color:"#7c3aed",border:"1px solid #7c3aed",borderRadius:6,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>📧 פנייה</button>
    {showModal&&<TaskInquiryModal task={task} user={user} onClose={()=>setShowModal(false)}/>}
  </>;
}

function TaskInquiryModal({task,user,onClose}){
  const [step,setStep]=useState("selectTemplate");
  const [selectedTemplate,setSelectedTemplate]=useState(null);
  const [savedRecipients,setSavedRecipients]=useState([]);
  const [selectedRecipient,setSelectedRecipient]=useState(null);
  const [showAddNew,setShowAddNew]=useState(false);
  const [newEmail,setNewEmail]=useState("");
  const [newName,setNewName]=useState("");
  const [newRole,setNewRole]=useState("");
  const [subject,setSubject]=useState("");
  const [body,setBody]=useState("");
  const [sending,setSending]=useState(false);
  const fld={width:"100%",padding:10,borderRadius:10,border:"1px solid var(--border)",fontFamily:"inherit",fontSize:13};
  const orderNum=task.bina_order_id||String(task.id||"").slice(0,8)||"---";
  const clientName=task.client_name||"לקוח";
  const taskTitle=task.title||"";
  useEffect(()=>{
    sb("task_inquiry_recipients?order=is_default.desc,usage_count.desc,name.asc").then(d=>setSavedRecipients(Array.isArray(d)?d:[])).catch(()=>setSavedRecipients([]));
  },[]);
  const templates=[
    {id:"status",icon:"🔍",label:"סטטוס",color:"#3b82f6",title:"מה הסטטוס?",buildContent:()=>({subject:`בקשת עדכון סטטוס - ${clientName} (#${orderNum})`,body:`שלום,\n\nאשמח לקבל עדכון על סטטוס המשימה:\n\nלקוח: ${clientName}\nהזמנה: #${orderNum}\nתיאור: ${taskTitle}\n\nהאם יש התקדמות? באיזה שלב אנחנו?\n\nתודה,\n${user?.name||""}`})},
    {id:"reminder",icon:"⏰",label:"תזכורת",color:"#f59e0b",title:"תזכורת לתאריך",buildContent:()=>({subject:`תזכורת - ${clientName} (#${orderNum})`,body:`שלום,\n\nתזכורת לגבי המשימה:\n\nלקוח: ${clientName}\nהזמנה: #${orderNum}\nתיאור: ${taskTitle}\n\nההזמנה צריכה להיות מוכנה עד תאריך: [הזן תאריך]\n\nאנא וודא שאנחנו בלוח הזמנים.\n\nתודה,\n${user?.name||""}`})},
    {id:"urgent",icon:"🚨",label:"דחיפות",color:"#ef4444",title:"שינוי / דחיפות",buildContent:()=>({subject:`🚨 דחוף - שינוי בהזמנה - ${clientName} (#${orderNum})`,body:`שלום,\n\n⚠️ עדכון דחוף לגבי המשימה:\n\nלקוח: ${clientName}\nהזמנה: #${orderNum}\nתיאור: ${taskTitle}\n\nיש שינוי / דחיפות:\n[פרט]\n\nהאם אפשר להתאים בלוח הזמנים?\n\nתודה,\n${user?.name||""}`})},
    {id:"update",icon:"✓",label:"עדכון כשמוכן",color:"#10b981",title:"תן עדכון כשמוכן",buildContent:()=>({subject:`בקשת עדכון בסיום - ${clientName} (#${orderNum})`,body:`שלום,\n\nאשמח שתעדכן אותי ברגע שהמשימה מוכנה:\n\nלקוח: ${clientName}\nהזמנה: #${orderNum}\nתיאור: ${taskTitle}\n\nזה חשוב לתיאום ההמשך.\n\nתודה,\n${user?.name||""}`})},
  ];
  const handleSelectTemplate=t=>{setSelectedTemplate(t);const x=t.buildContent();setSubject(x.subject);setBody(x.body);setStep("selectRecipient");};
  const handleSelectExisting=async recipient=>{setSelectedRecipient(recipient);try{await sb(`task_inquiry_recipients?id=eq.${recipient.id}`,{method:"PATCH",body:JSON.stringify({usage_count:(recipient.usage_count||0)+1,last_used_at:new Date().toISOString()})});}catch(_){}setStep("edit");};
  const handleAddNewAndContinue=async()=>{
    if(!newEmail.trim()||!newName.trim()){window.alert("יש למלא שם ומייל");return;}
    if(!newEmail.includes("@")){window.alert("כתובת מייל לא תקינה");return;}
    try{
      const result=await sb("task_inquiry_recipients",{method:"POST",body:JSON.stringify({name:newName.trim(),email:newEmail.trim().toLowerCase(),role:newRole.trim()||null,is_default:savedRecipients.length===0,usage_count:1,created_by:user?.name||"unknown"}),headers:{Prefer:"return=representation"}});
      const row=Array.isArray(result)?result[0]:result;
      if(row){setSelectedRecipient(row);const updated=await sb("task_inquiry_recipients?order=is_default.desc,usage_count.desc,name.asc");setSavedRecipients(Array.isArray(updated)?updated:[]);setShowAddNew(false);setNewEmail("");setNewName("");setNewRole("");setStep("edit");}
    }catch(e){const msg=e.message||String(e);window.alert("כתובת מייל זו כבר קיימת במאגר" if ("duplicate" in msg or "unique" in msg) else "שגיאה: "+msg);}
  };
  const handleSend=async()=>{
    if(!subject.trim()||!body.trim()){window.alert("יש למלא נושא ותוכן");return;}
    setSending(true);setStep("sending");
    try{
      const htmlBody=`<motion.div style="direction:rtl;font-family:Arial,sans-serif;font-size:14px;line-height:1.6;color:#333;">${body.replace(/\\n/g,"<br>")}</motion.div>`;
      await sendEmail({to:selectedRecipient.email,cc:[ADMIN_NOTIFY_EMAIL],subject,html:htmlBody,text:body,from:"דפוס נטלי <orders@natalie-print.com>"});
      window.alert(`✅ הפנייה נשלחה ל-${selectedRecipient.name}!`);onClose();
    }catch(e){window.alert("שגיאה בשליחה: "+(e.message||String(e)));setStep("edit");}
    finally{setSending(false);}
  };
  return<Modal open={true} onClose={onClose} title="📧 פנייה לאחראי" maxWidth={600}>
    <motion.div style={{direction:"rtl"}}>
      <div style={{background:"#f9fafb",padding:10,borderRadius:6,marginBottom:15,fontSize:13}}><strong>{clientName}</strong>{task.bina_order_id&&<> · <span style={{direction:"ltr",display:"inline-block"}}>#{task.bina_order_id}</span></>}{taskTitle&&<div style={{color:"#6b7280",marginTop:3}}>{taskTitle}</div>}</div>
      {step==="selectTemplate"&&<><p style={{margin:"0 0 15px",color:"#6b7280"}}>בחר תבנית פנייה:</p><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>{templates.map(t=><button type="button" key={t.id} onClick={()=>handleSelectTemplate(t)} style={{padding:15,background:"#fff",border:`2px solid ${t.color}`,borderRadius:8,cursor:"pointer",textAlign:"right",fontFamily:"inherit"}}><div style={{fontSize:24,marginBottom:4}}>{t.icon}</div><motion.div style={{fontWeight:600,color:t.color}}>{t.label}</motion.div><div style={{fontSize:11,color:"#6b7280",marginTop:4}}>{t.title}</div></button>)}</div><button type="button" onClick={onClose} style={{width:"100%",padding:10,marginTop:15,background:"#f3f4f6",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit"}}>ביטול</button></>}
      {step==="selectRecipient"&&selectedTemplate&&<><div style={{background:selectedTemplate.color+"20",padding:10,borderRadius:6,marginBottom:15,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{fontWeight:600,color:selectedTemplate.color}}>{selectedTemplate.icon} {selectedTemplate.title}</div><button type="button" onClick={()=>setStep("selectTemplate")} style={{background:"transparent",border:"none",color:selectedTemplate.color,cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>שנה</button></div><p style={{margin:"0 0 12px",color:"#6b7280"}}>למי לשלוח?</p>{savedRecipients.length>0&&<div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:12}}>{savedRecipients.map(r=><button type="button" key={r.id} onClick={()=>handleSelectExisting(r)} style={{padding:12,background:"#f9fafb",border:"2px solid transparent",borderRadius:8,cursor:"pointer",textAlign:"right",fontFamily:"inherit",width:"100%"}} onMouseEnter={e=>e.currentTarget.style.borderColor="#7c3aed"} onMouseLeave={e=>e.currentTarget.style.borderColor="transparent"}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8}}><div><div style={{fontWeight:600}}>{r.is_default?"⭐ ":""}{r.name}{r.role&&<span style={{color:"#6b7280",fontSize:12,marginRight:8}}> · {r.role}</span>}</div><div style={{fontSize:12,color:"#6b7280",marginTop:2,direction:"ltr",textAlign:"right"}}>{r.email}</div></div>{r.usage_count>0&&<div style={{fontSize:11,color:"#9ca3af"}}>{r.usage_count} פניות</div>}</div></button>)}</div>}{!showAddNew?<button type="button" onClick={()=>setShowAddNew(true)} style={{width:"100%",padding:10,background:"transparent",border:"2px dashed #d1d5db",borderRadius:8,cursor:"pointer",color:"#6b7280",fontSize:13,fontFamily:"inherit"}}>➕ הוסף נמען חדש</button>:<div style={{padding:15,background:"#f9fafb",borderRadius:8,border:"1px solid #d1d5db"}}><input value={newName} onChange={e=>setNewName(e.target.value)} placeholder="שם הנמען *" style={{...fld,padding:8,marginBottom:8}} autoFocus/><input type="email" value={newEmail} onChange={e=>setNewEmail(e.target.value)} placeholder="email@example.com *" style={{...fld,padding:8,marginBottom:8,direction:"ltr"}}/><input value={newRole} onChange={e=>setNewRole(e.target.value)} placeholder="תפקיד (אופציונלי)" style={{...fld,padding:8,marginBottom:10}}/><div style={{display:"flex",gap:8}}><button type="button" onClick={()=>{setShowAddNew(false);setNewEmail("");setNewName("");setNewRole("");}} style={{flex:1,padding:8,fontSize:13,cursor:"pointer",fontFamily:"inherit",border:"1px solid var(--border)",borderRadius:6,background:"#fff"}}>ביטול</button><button type="button" onClick={handleAddNewAndContinue} style={{flex:2,padding:8,background:"#7c3aed",color:"#fff",border:"none",borderRadius:6,fontSize:13,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>שמור והמשך</button></div></motion.div>}<button type="button" onClick={onClose} style={{width:"100%",padding:10,marginTop:15,background:"#f3f4f6",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit"}}>ביטול</button></>}
      {step==="edit"&&selectedTemplate&&selectedRecipient&&<><div style={{background:selectedTemplate.color+"20",padding:10,borderRadius:6,marginBottom:10}}><div style={{fontWeight:600,color:selectedTemplate.color}}>{selectedTemplate.icon} {selectedTemplate.title}</div></div><div style={{background:"#f0fdfa",padding:10,borderRadius:6,marginBottom:15,fontSize:13}}><div>📧 שולח אל: <strong>{selectedRecipient.name}</strong></div><div style={{color:"#6b7280",direction:"ltr",textAlign:"right",marginTop:2}}>{selectedRecipient.email}</div><div style={{fontSize:11,color:"#9ca3af",marginTop:4}}>ℹ️ CC לכפיר אוטומטי</div></div><div style={{marginBottom:12}}><label style={{display:"block",marginBottom:4,fontWeight:600,fontSize:14}}>נושא:</label><input value={subject} onChange={e=>setSubject(e.target.value)} style={fld}/></div><div style={{marginBottom:15}}><label style={{display:"block",marginBottom:4,fontWeight:600,fontSize:14}}>תוכן:</label><textarea value={body} onChange={e=>setBody(e.target.value)} style={{...fld,minHeight:200,resize:"vertical",direction:"rtl"}}/></div><div style={{display:"flex",gap:10}}><button type="button" onClick={()=>setStep("selectRecipient")} style={{flex:1,padding:10,background:"#f3f4f6",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit"}}>← חזור</button><button type="button" onClick={handleSend} disabled={sending} style={{flex:2,padding:12,background:sending?"#d1d5db":"#7c3aed",color:"#fff",border:"none",borderRadius:6,fontWeight:600,cursor:sending?"not-allowed":"pointer",fontFamily:"inherit"}}>📧 שלח</button></div></>}
      {step==="sending"&&<div style={{textAlign:"center",padding:30}}><div style={{fontSize:32,marginBottom:10}}>📧</div><div>שולח...</div></div>}
    </div>
  </Modal>;
}

function RecentlyCompletedCard({task,onRestore}){
  const completedDate=new Date(task.completed_at);
  const hoursAgo=Math.floor((Date.now()-completedDate.getTime())/(1000*60*60));
  let timeText;
  if(hoursAgo<1)timeText="לפני פחות משעה";
  else if(hoursAgo<24)timeText=`לפני ${hoursAgo} שעות`;
  else timeText=`לפני ${Math.floor(hoursAgo/24)} ימים`;
  const handleRestore=async()=>{
    if(!window.confirm("להחזיר את המשימה לרשימת הפעילות?"))return;
    try{await sb(`tasks?id=eq.${task.id}`,{method:"PATCH",body:JSON.stringify({completed_at:null,status:"בביצוע"})});onRestore?.();}
    catch(e){window.alert("שגיאה: "+(e.message||String(e)));}
  };
  return<div style={{background:"#fff",padding:12,borderRadius:8,border:"1px solid #e5e7eb",display:"flex",justifyContent:"space-between",alignItems:"center",gap:10}}>
    <div style={{flex:1,minWidth:0}}>
      <div style={{fontWeight:600,fontSize:14}}>✅ {task.client_name||"—"}{task.bina_order_id&&<span style={{color:"#9ca3af",marginRight:8,fontSize:12,direction:"ltr",display:"inline-block"}}>#{task.bina_order_id}</span>}</div>
      <div style={{fontSize:13,color:"#6b7280",marginTop:3}}>{task.title}</div>
      <div style={{fontSize:11,color:"#9ca3af",marginTop:3}}>הושלם {timeText}{task.sales_agent?` · ${task.sales_agent}`:""}</div>
    </div>
    <button type="button" onClick={handleRestore} title="החזר לרשימת הפעילות" style={{padding:"6px 10px",background:"#f3f4f6",border:"1px solid #d1d5db",borderRadius:6,cursor:"pointer",fontSize:12,fontFamily:"inherit",flexShrink:0}}>↩️ החזר</button>
  </div>;
}

function RecentlyCompletedTab({tasks,loading,onRestore}){
  const recent=useMemo(()=>tasks.filter(isTaskRecentlyCompleted).sort((a,b)=>new Date(b.completed_at)-new Date(a.completed_at)),[tasks]);
  if(loading)return<div style={{padding:30,textAlign:"center"}}><Spinner sz={28}/></div>;
  return<div style={{padding:4}}>
    <h2 style={{marginTop:0,fontSize:22,fontWeight:800}}>🗂️ משימות שהושלמו ב-3 ימים האחרונים</h2>
    <div style={{background:"#f0fdfa",padding:12,borderRadius:8,marginBottom:15,fontSize:13,color:"#0d9488"}}>ℹ️ כאן מופיעות משימות שהושלמו לאחרונה. אחרי 3 ימים — הן יעברו לארכיון המלא.</div>
    {recent.length===0?<div style={{textAlign:"center",padding:40,color:"#9ca3af"}}>🎉 אין משימות שהושלמו לאחרונה</div>:
    <><div style={{marginBottom:10,color:"#6b7280",fontSize:13}}>{recent.length} משימות</div><div style={{display:"grid",gap:8}}>{recent.map(t=><RecentlyCompletedCard key={t.id} task={t} onRestore={onRestore}/>)}</div></>}
  </div>;
}

'''.replace("motion.div", "div")

if "function TaskInquiryButton" not in c:
    c = c.replace(INSERT_MARKER, COMPONENTS.strip() + "\n\n" + INSERT_MARKER, 1)

must_replace("  const visTasks=boardTasks;", "  const visTasks=boardTasks.filter(isTaskActive);", "visTasks")

must_replace(
    "  const load=async(u)=>{\n    const cu=u||user;if(!cu)return;\n    setLoading(true);\n    try{",
    "  const load=async(u)=>{\n    const cu=u||user;if(!cu)return;\n    setLoading(true);\n    try{\n      await archiveStaleCompleted();",
    "load archive",
)

must_replace(
    '      const patch={status:s};\n      if(s==="הושלם")patch.completed_at=(new Date().toISOString());\n      await sb(`tasks?id=eq.${id}`,{method:"PATCH",body:JSON.stringify(patch)});\n      setTasks(p=>p.map(t=>t.id===id?{...t,status:s}:t));',
    '      const patch={status:s};\n      if(s==="הושלם")patch.completed_at=(new Date().toISOString());\n      else patch.completed_at=null;\n      await sb(`tasks?id=eq.${id}`,{method:"PATCH",body:JSON.stringify(patch)});\n      setTasks(p=>p.map(t=>t.id===id?{...t,status:s,completed_at:patch.completed_at}:t));',
    "updStatus",
)

must_replace(
    '    {id:"tasks",ic:"☑",lb:"משימות"},\n    {id:"calendar",ic:"📅",lb:"לוח שנה"},',
    '    {id:"tasks",ic:"☑",lb:"משימות"},\n    {id:"recently_completed",ic:"🗂️",lb:"הושלמו לאחרונה"},\n    {id:"calendar",ic:"📅",lb:"לוח שנה"},',
    "nav",
)

must_replace(
    '      {tab==="calendar"&&<CalendarTab tasks={tasks} taskItems={taskItems} user={user} isA={isA} onTaskClick={setDet} clients={clients} onTaskRefresh={refreshAll}/>}',
    '      {tab==="recently_completed"&&<RecentlyCompletedTab tasks={tasks} loading={loading} onRestore={refreshAll}/>}\n      {tab==="calendar"&&<CalendarTab tasks={tasks} taskItems={taskItems} user={user} isA={isA} onTaskClick={setDet} clients={clients} onTaskRefresh={refreshAll}/>}',
    "recent tab",
)

must_replace(
    "              <TaskNotificationButtons task={t} user={user} clients={clients} onUpdate={refreshAll} />\n              <TaskMoveToDeliveryBtn task={t}/>",
    "              <TaskNotificationButtons task={t} user={user} clients={clients} onUpdate={refreshAll} />\n              <TaskInquiryButton task={t} user={user} />\n              <TaskMoveToDeliveryBtn task={t}/>",
    "tasks inquiry",
)

must_replace(
    "          <TaskNotificationButtons task={det} user={user} clients={clients} onUpdate={refreshAll} />\n          <TaskMoveToDeliveryBtn task={det}/>",
    "          <TaskNotificationButtons task={det} user={user} clients={clients} onUpdate={refreshAll} />\n          <TaskInquiryButton task={det} user={user} />\n          <TaskMoveToDeliveryBtn task={det}/>",
    "det inquiry",
)

must_replace(
    "function ShipmentDetailsModal({task,drivers,onClose,onUpdate,onOpenDeliveryNote}){",
    "function ShipmentDetailsModal({task,drivers,onClose,onUpdate,onOpenDeliveryNote,user}){",
    "shipment modal sig",
)

must_replace(
    '    {showDetails&&<ShipmentDetailsModal task={task} drivers={drivers} onClose={()=>setShowDetails(false)} onUpdate={onRefresh} onOpenDeliveryNote={onOpenDeliveryNote}/>}',
    '    {showDetails&&<ShipmentDetailsModal task={task} drivers={drivers} onClose={()=>setShowDetails(false)} onUpdate={onRefresh} onOpenDeliveryNote={onOpenDeliveryNote} user={user}/>}',
    "shipment cube modal user",
)

must_replace(
    '      <motion.div style={{display:"flex",gap:8,flexWrap:"wrap"}}>\n        <button type="button" onClick={()=>setShowEdit(true)}',
    '      <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>\n        {user&&<TaskInquiryButton task={task} user={user} compact/>}\n        <button type="button" onClick={()=>setShowEdit(true)}',
    "shipment details inquiry",
)
c = c.replace("motion.div", "div")

if 'user&&<TaskInquiryButton task={task} user={user} compact/>' not in c.split("function ShipmentCube")[1].split("function SplitByDriverModal")[0]:
    must_replace(
        '        {!readonly&&!compact&&!isDelivered&&<div style={{display:"flex",gap:4,marginTop:6,paddingTop:6,borderTop:"1px solid var(--bg2)"}} onClick={e=>e.stopPropagation()}>\n          <button type="button" onClick={handleCopySingle}',
        '        {!readonly&&!compact&&!isDelivered&&<motion.div style={{display:"flex",gap:4,marginTop:6,paddingTop:6,borderTop:"1px solid var(--bg2)",flexWrap:"wrap"}} onClick={e=>e.stopPropagation()}>\n          {user&&<TaskInquiryButton task={task} user={user} compact/>}\n          <button type="button" onClick={handleCopySingle}',
        "cube inquiry",
    )
c = c.replace("motion.div", "motion.div")
c = c.replace("motion.div", "div")

cal = c.split("function CalendarTab")[1].split("// ── MAIN APP")[0]
if "TaskInquiryButton task={t}" not in cal:
    must_replace(
        '                <div style={{flexShrink:0,transform:"scale(0.78)",transformOrigin:"center right"}}><TaskNotificationButtons task={t} user={user} clients={clients} onUpdate={onTaskRefresh} compact/></div>',
        '                <div style={{flexShrink:0,transform:"scale(0.78)",transformOrigin:"center right"}}><TaskInquiryButton task={t} user={user} compact/></div>\n                <div style={{flexShrink:0,transform:"scale(0.78)",transformOrigin:"center right"}}><TaskNotificationButtons task={t} user={user} clients={clients} onUpdate={onTaskRefresh} compact/></div>',
        "calendar",
    )

must_replace(
    "          {boardTasks.slice(0,5).map(t=><div key={t.id}",
    "          {boardTasks.filter(isTaskActive).slice(0,5).map(t=><div key={t.id}",
    "dash recent",
)

# Calendar filter completed from day view
must_replace(
    "  const getTasksForDay=(dateStr)=>{\n    let filtered=tasks.filter(t=>t.due_date===dateStr);",
    "  const getTasksForDay=(dateStr)=>{\n    let filtered=tasks.filter(t=>t.due_date===dateStr&&isTaskActive(t));",
    "cal filter",
)

if re.search(r"motion\.div", c):
    raise SystemExit("motion.div remains")

p.write_text(c, encoding="utf-8")
print("wave2 patch ok")

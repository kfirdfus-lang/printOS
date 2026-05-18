import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const indexPath = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "index.html");
let html = fs.readFileSync(indexPath, "utf8");

const replacements = [
  [
    `{user&&<div onClick={e=>e.stopPropagation()} style={{marginTop:compact?4:6}}><TaskNotificationButtons`,
    `{user&&!isDelivered&&<div onClick={e=>e.stopPropagation()} style={{marginTop:compact?4:6}}><TaskNotificationButtons`,
  ],
  [
    `{!readonly&&!compact&&<div style={{display:"flex",gap:4,marginTop:6,paddingTop:6,borderTop:"1px solid var(--bg2)"}} onClick={e=>e.stopPropagation()}>`,
    `{!readonly&&!compact&&!isDelivered&&<div style={{display:"flex",gap:4,marginTop:6,paddingTop:6,borderTop:"1px solid var(--bg2)"}} onClick={e=>e.stopPropagation()}>`,
  ],
  [
    `        </div>}
      </>
    </motion.div>
    {showDetails&&<ShipmentDetailsModal task={task} drivers={drivers} onClose={()=>setShowDetails(false)} onUpdate={onRefresh} onOpenDeliveryNote={onOpenDeliveryNote}/>}`,
    `        </div>}
        {!compact&&isDelivered&&<div style={{marginTop:6,paddingTop:6,borderTop:"1px solid #d1d5db",fontSize:11,color:"#10b981",fontWeight:600,textAlign:"center"}}>
          ✅ נמסר בהצלחה
          {task.delivery_marked_at&&<div style={{fontSize:10,color:"#9ca3af",marginTop:2}}>{new Date(task.delivery_marked_at).toLocaleDateString("he-IL")}</div>}
        </div>}
      </>
    </motion.div>
    {showDetails&&<ShipmentDetailsModal task={task} drivers={drivers} onClose={()=>setShowDetails(false)} onUpdate={onRefresh} onOpenDeliveryNote={onOpenDeliveryNote}/>}`,
  ],
];

for (const [from, to] of replacements) {
  const f = from.replace(/motion\.div/g, "div");
  const t = to.replace(/motion\.div/g, "div");
  if (!html.includes(f)) {
    console.warn("skip, not found:", f.slice(0, 60));
    continue;
  }
  html = html.replace(f, t);
}

const splitModal = `
function SplitByDriverModal({date,tasks,drivers,onClose}){
  const tasksByDriver=useMemo(()=>{
    const groups={};
    tasks.forEach(t=>{
      const driverId=t.assigned_driver_id||"_unassigned";
      if(!groups[driverId])groups[driverId]=[];
      groups[driverId].push(t);
    });
    return groups;
  },[tasks]);
  const getDriverName=(driverId)=>{
    if(driverId==="_unassigned")return "⚠️ ללא נהג";
    const d=drivers.find(dr=>dr.id===driverId);
    return d?\`🚛 \${d.name}\`:"נהג לא ידוע";
  };
  const formatDriverText=(driverId,driverTasks)=>{
    const driverName=getDriverName(driverId);
    const dateStr=date.toLocaleDateString("he-IL",{weekday:"long",day:"numeric",month:"numeric"});
    let text=\`📦 *משלוחים \${driverName}*\\n📅 \${dateStr}\\nסה"כ: \${driverTasks.length} משלוחים\\n\\n━━━━━━━━━━━━━━\\n\\n\`;
    driverTasks.forEach((t,i)=>{
      text+=\`*\${i+1}. \${t.client_name||"—"}*\\n\`;
      if(t.bina_order_id)text+=\`📝 הזמנה #\${t.bina_order_id}\\n\`;
      if(t.sales_agent)text+=\`👤 סוכן: \${t.sales_agent}\\n\`;
      if(t.delivery_address_snapshot)text+=\`📍 \${t.delivery_address_snapshot}\\n\`;
      if(t.delivery_contact_snapshot||t.delivery_phone_snapshot){
        const line=\`\${t.delivery_contact_snapshot||""} \${t.delivery_phone_snapshot||""}\`.trim();
        if(line)text+=\`📞 \${line}\\n\`;
      }
      const items=Array.isArray(t.items)?t.items:[];
      if(items.length>0){
        text+=\`\\n📦 פריטים:\\n\`;
        items.forEach(item=>{text+=\`   • \${item.description||"—"} (×\${item.quantity??0})\\n\`;});
      }
      if(t.delivery_special_notes)text+=\`\\n💬 \${t.delivery_special_notes}\\n\`;
      text+=\`\\n━━━━━━━━━━━━━━\\n\\n\`;
    });
    return text;
  };
  const handleCopyDriver=async(driverId)=>{
    const text=formatDriverText(driverId,tasksByDriver[driverId]);
    await navigator.clipboard.writeText(text);
    window.alert(\`✅ הועתק - \${tasksByDriver[driverId].length} משלוחים של \${getDriverName(driverId)}\`);
  };
  const handleWhatsAppDriver=(driverId)=>{
    const text=formatDriverText(driverId,tasksByDriver[driverId]);
    window.open(\`https://wa.me/?text=\${encodeURIComponent(text)}\`,"_blank");
  };
  const driverIds=Object.keys(tasksByDriver);
  return<div onClick={onClose} style={SIGNUP_MODAL_OVERLAY}>
    <motion.div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:16,padding:24,width:"100%",maxWidth:600,maxHeight:"90vh",overflowY:"auto",direction:"rtl"}}>
      <h2 style={{margin:"0 0 16px",color:"#0d9488"}}>🚛 פיצול משלוחים לפי נהג</h2>
      <motion.div style={{background:"#f0f9ff",padding:12,borderRadius:8,marginBottom:15,fontSize:14}}>
        📅 {date.toLocaleDateString("he-IL",{weekday:"long",day:"numeric",month:"long"})}
        <br/><span style={{color:"#6b7280",fontSize:13}}>סה"כ {tasks.length} משלוחים · {driverIds.length} נהגים</span>
      </motion.div>
      {driverIds.length===0?<motion.div style={{textAlign:"center",padding:30,color:"#9ca3af"}}>אין משלוחים ביום זה</motion.div>:
      <motion.div style={{display:"flex",flexDirection:"column",gap:10}}>
        {driverIds.map(driverId=>{
          const driverTasks=tasksByDriver[driverId];
          const driver=drivers.find(d=>d.id===driverId);
          const isUnassigned=driverId==="_unassigned";
          return<motion.div key={driverId} style={{background:isUnassigned?"#fef2f2":"white",border:\`2px solid \${isUnassigned?"#fca5a5":(driver?.color||"#d1d5db")}\`,borderRadius:8,padding:12}}>
            <motion.div style={{fontWeight:600,fontSize:15,marginBottom:4}}>{getDriverName(driverId)}</motion.div>
            <motion.div style={{fontSize:12,color:"#6b7280",marginBottom:8}}>{driverTasks.length} משלוחים</motion.div>
            <motion.div style={{fontSize:12,color:"#6b7280",marginBottom:10,maxHeight:80,overflowY:"auto"}}>
              {driverTasks.map((t,i)=><motion.div key={t.id} style={{padding:"2px 0"}}>{i+1}. {t.client_name}{t.delivery_address_snapshot?\` - \${t.delivery_address_snapshot.substring(0,30)}\${t.delivery_address_snapshot.length>30?"...":""}\`:""}</motion.div>)}
            </motion.div>
            <motion.div style={{display:"flex",gap:8}}>
              <button type="button" onClick={()=>handleCopyDriver(driverId)} style={{flex:1,padding:"8px 10px",background:"#f3f4f6",border:"1px solid #d1d5db",borderRadius:6,cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>📋 העתק</button>
              <button type="button" onClick={()=>handleWhatsAppDriver(driverId)} style={{flex:1,padding:"8px 10px",background:"#25d366",color:"white",border:"none",borderRadius:6,cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>📱 WhatsApp</button>
            </motion.div>
          </motion.div>;
        })}
      </motion.div>}
      <button type="button" onClick={onClose} style={{width:"100%",padding:10,marginTop:15,background:"#f3f4f6",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit"}}>סגור</button>
    </motion.div>
  </motion.div>;
}
`.replace(/motion\.div/g, "div");

if (!html.includes("function SplitByDriverModal")) {
  html = html.replace("function DayColumn(", splitModal + "\nfunction DayColumn(");
}

html = html.replace(
  "function DayColumn({date,tasks,drivers,onDragStart,onDrop,onCopyDay,onSendWhatsApp,isToday,onOpenDeliveryNote,onRefresh,user,clients})",
  "function DayColumn({date,tasks,drivers,onDragStart,onDrop,onCopyDay,onSendWhatsApp,onSplitDay,isToday,onOpenDeliveryNote,onRefresh,user,clients})"
);

html = html.replace(
  `<button type="button" onClick={onSendWhatsApp} style={{flex:1,padding:6,fontSize:11,background:"#25d366",color:"#fff",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit"}}>📱 WA</button>
    </motion.div>}`,
  `<button type="button" onClick={onSendWhatsApp} style={{flex:1,padding:6,fontSize:11,background:"#25d366",color:"#fff",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit"}}>📱 WA</button>
      {onSplitDay&&<button type="button" onClick={onSplitDay} style={{flex:1,padding:6,fontSize:11,background:"#f59e0b",color:"#fff",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontWeight:600}} title="הפרד לפי נהג">🚛 פצל</button>}
    </motion.div>}`
).replace(/motion\.div/g, "motion.div");

html = html.replace(
  "const [deliveryNoteModal,setDeliveryNoteModal]=useState(null);",
  "const [deliveryNoteModal,setDeliveryNoteModal]=useState(null);\n  const [showSplitByDriverModal,setShowSplitByDriverModal]=useState(null);"
);

html = html.replace(
  `onSendWhatsApp={()=>sendDayWhatsApp(date)} isToday=`,
  `onSendWhatsApp={()=>sendDayWhatsApp(date)} onSplitDay={()=>setShowSplitByDriverModal({date,tasks:getTasksForDay(date)})} isToday=`
);

html = html.replace(
  "{deliveryNoteModal&&<DeliveryNoteModal task={deliveryNoteModal.task} clients={clients} user={user} onClose={()=>setDeliveryNoteModal(null)}/>}",
  "{deliveryNoteModal&&<DeliveryNoteModal task={deliveryNoteModal.task} clients={clients} user={user} onClose={()=>setDeliveryNoteModal(null)}/>}\n    {showSplitByDriverModal&&<SplitByDriverModal date={showSplitByDriverModal.date} tasks={showSplitByDriverModal.tasks} drivers={drivers} onClose={()=>setShowSplitByDriverModal(null)}/>}"
);

html = html.replace(/motion\.div/g, "div");
fs.writeFileSync(indexPath, html, "utf8");
console.log("patched shipments wave 1");

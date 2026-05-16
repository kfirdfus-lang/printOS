const fs = require("fs");
const path = require("path");
const p = path.join(__dirname, "..", "index.html");
let lines = fs.readFileSync(p, "utf8").split(/\r?\n/);
const i = lines.findIndex((l) => l.includes('step==="pickEmail"') && l.includes("availableEmails"));
if (i < 0) {
  console.error("pickEmail line not found");
  process.exit(1);
}
const pickEmail = `      {step==="pickEmail"&&<motion.div><p style={{marginBottom:15}}>בחר מייל לשליחה ללקוח <strong>{client?.name||task.client_name}</strong>:</p><motion.div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:15}}>{savedEmails.map(emailRec=><button type="button" key={emailRec.id} onClick={()=>handleEmailPick(emailRec.email)} style={{padding:12,background:"#f3f4f6",border:"2px solid transparent",borderRadius:8,cursor:"pointer",fontFamily:"inherit",display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,width:"100%"}} onMouseEnter={e=>e.currentTarget.style.borderColor="#0e7490"} onMouseLeave={e=>e.currentTarget.style.borderColor="transparent"}><motion.div style={{textAlign:"right"}}><motion.div style={{direction:"ltr",fontWeight:600,fontSize:14}}>{emailRec.is_default?"⭐ ":""}📧 {emailRec.email}</motion.div>{emailRec.label&&<motion.div style={{fontSize:12,color:"#6b7280",marginTop:2}}>{emailRec.label}</motion.div>}</motion.div>{emailRec.usage_count>0&&<motion.div style={{fontSize:11,color:"#9ca3af"}}>נשלח {emailRec.usage_count} פעמים</motion.div>}</button>)}</motion.div>{!showAddEmail?<button type="button" onClick={()=>setShowAddEmail(true)} style={{width:"100%",padding:10,background:"transparent",border:"2px dashed #d1d5db",borderRadius:8,cursor:"pointer",color:"#6b7280",fontSize:13,fontFamily:"inherit",marginBottom:10}}>➕ הוסף מייל חדש למאגר</button>:<motion.div style={{padding:15,background:"#f9fafb",borderRadius:8,border:"1px solid #d1d5db",marginBottom:10}}><input type="email" value={newEmail} onChange={e=>setNewEmail(e.target.value)} placeholder="example@company.com" style={{...fld,padding:8,direction:"ltr",marginBottom:8}} autoFocus/><input type="text" value={newLabel} onChange={e=>setNewLabel(e.target.value)} placeholder="תווית (אופציונלי)" style={{...fld,padding:8,marginBottom:8}}/><motion.div style={{display:"flex",gap:8}}><button type="button" onClick={()=>{setShowAddEmail(false);setNewEmail("");setNewLabel("");}} style={{flex:1,padding:8,fontSize:13,cursor:"pointer",fontFamily:"inherit",border:"1px solid var(--border)",borderRadius:6,background:"#fff"}}>ביטול</button><button type="button" onClick={handleAddEmail} style={{flex:2,padding:8,background:"#0e7490",color:"#fff",border:"none",borderRadius:6,fontSize:13,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>הוסף ושלח</button></motion.div></motion.div>}<button type="button" onClick={onClose} style={{width:"100%",padding:10,background:"#f3f4f6",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit"}}>ביטול</button></motion.div>}`;

const addFirst = `      {step==="addFirstEmail"&&<motion.div>
        <motion.div style={{padding:15,background:"#fef3c7",color:"#78350f",borderRadius:8,marginBottom:15,fontSize:14}}>ℹ️ אין מיילי התראה ללקוח <strong>{client?.name||task.client_name}</strong>.<br/><span style={{fontSize:13}}>הוסף כתובת מייל אחת ויותר כדי לאפשר שליחת התראות בעתיד.</span></motion.div>
        <motion.div style={{marginBottom:12}}><label style={{display:"block",marginBottom:4,fontWeight:600}}>📧 כתובת מייל:</label><input type="email" value={newEmail} onChange={e=>setNewEmail(e.target.value)} placeholder="example@company.com" style={{...fld,direction:"ltr"}} autoFocus/></motion.div>
        <motion.div style={{marginBottom:15}}><label style={{display:"block",marginBottom:4,fontWeight:600}}>🏷️ תווית (אופציונלי):</label><input type="text" value={newLabel} onChange={e=>setNewLabel(e.target.value)} placeholder="למשל: רחל מכירות" style={fld}/></motion.div>
        <motion.div style={{display:"flex",gap:10}}><button type="button" onClick={onClose} style={{flex:1,padding:10,borderRadius:8,border:"1px solid var(--border)",background:"#fff",cursor:"pointer",fontFamily:"inherit"}}>ביטול</button><button type="button" onClick={handleAddEmail} style={{flex:2,padding:12,borderRadius:8,border:"none",background:"#0e7490",color:"#fff",fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>💾 שמור והמשך</button></motion.div>
      </motion.div>}`;

const fix = (s) => s.replace(/motion\.div/g, "div");
lines[i] = fix(pickEmail);
lines.splice(i, 0, fix(addFirst));
fs.writeFileSync(p, lines.join("\n"), "utf8");
console.log("patched lines", i, "and", i + 1);

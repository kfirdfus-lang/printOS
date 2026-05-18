import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const indexPath = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "index.html");
let s = fs.readFileSync(indexPath, "utf8");

const contactLine =
  '          <div><label style={{display:"block",marginBottom:4,fontSize:12}}>👤 איש קשר</label><input value={contactName} onChange={e=>setContactName(e.target.value)} style={fld}/></motion.div>';
const contactLineOk =
  '          <div><label style={{display:"block",marginBottom:4,fontSize:12}}>👤 איש קשר</label><input value={contactName} onChange={e=>setContactName(e.target.value)} style={fld}/></div>';

const lineIdx = s.indexOf(contactLineOk);
if (lineIdx < 0) {
  console.error("contact line not found");
  process.exit(1);
}

const gridStart = s.lastIndexOf(
  '<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>',
  lineIdx
);
const endIdx = s.indexOf("      </>}", lineIdx);
if (gridStart < 0 || endIdx < 0) {
  console.error("bounds", gridStart, endIdx);
  process.exit(1);
}

const newUi = `        <div style={{marginBottom:15}}>
          <label style={{display:"block",marginBottom:6,fontWeight:600,fontSize:14}}>👤 איש קשר למשלוח:</label>
          {contactMode==="pick"&&!showAddContact&&<>
            <div style={{background:"#f0fdfa",padding:12,borderRadius:8,border:"1px solid #2dd4bf",marginBottom:8}}>
              <div style={{fontSize:12,color:"#0d9488",marginBottom:8}}>⚡ אנשי קשר שמורים ללקוח זה ({savedContacts.length})</div>
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {savedContacts.map(c=><button type="button" key={c.id} onClick={()=>handleSelectContact(c)} style={{padding:10,textAlign:"right",background:selectedContactId===c.id?"#0d9488":"#fff",color:selectedContactId===c.id?"#fff":"#374151",border:\`1px solid \${selectedContactId===c.id?"#0d9488":"#d1d5db"}\`,borderRadius:6,cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",fontFamily:"inherit"}}>
                  <motion.div>
                    <div style={{fontWeight:600}}>{c.is_default&&"⭐ "}{c.contact_name}</div>
                    <div style={{fontSize:12,color:selectedContactId===c.id?"rgba(255,255,255,0.85)":"#6b7280",marginTop:2,direction:"ltr",textAlign:"right"}}>📞 {c.contact_phone}</div>
                  </div>
                  {c.usage_count>0&&<div style={{fontSize:11,color:selectedContactId===c.id?"rgba(255,255,255,0.7)":"#9ca3af"}}>{c.usage_count} פעמים</div>}
                </button>)}
              </div>
            </div>
            <button type="button" onClick={handleStartAddNew} style={{width:"100%",padding:10,background:"transparent",border:"2px dashed #d1d5db",borderRadius:8,cursor:"pointer",color:"#6b7280",fontSize:13,fontFamily:"inherit"}}>➕ הוסף איש קשר חדש</button>
          </>}
          {(contactMode==="addNew"||showAddContact)&&<div style={{background:"#fefce8",padding:12,borderRadius:8,border:"1px solid #fde68a"}}>
            {contactMode==="addNew"&&<div style={{fontSize:12,color:"#78350f",marginBottom:8}}>ℹ️ אין אנשי קשר שמורים ללקוח זה - הוסף את הראשון</div>}
            <div style={{marginBottom:8}}><input type="text" value={contactName} onChange={e=>setContactName(e.target.value)} placeholder="שם איש הקשר *" style={{...fld,borderRadius:6}} autoFocus={contactMode==="addNew"}/></div>
            <div style={{marginBottom:contactMode==="pick"?8:0}}><input type="tel" value={phone} onChange={e=>setPhone(e.target.value)} placeholder="טלפון *" style={{...fld,borderRadius:6,direction:"ltr"}}/></div>
            {contactMode==="pick"&&showAddContact&&<div style={{display:"flex",gap:6,marginTop:8}}>
              <button type="button" onClick={()=>{
                setShowAddContact(false);
                const prevId=prevContactIdRef.current;
                if(prevId){
                  const prev=savedContacts.find(c=>c.id===prevId);
                  if(prev){setSelectedContactId(prev.id);setContactName(prev.contact_name);setPhone(prev.contact_phone);}
                }
              }} style={{flex:1,padding:8,fontSize:12,background:"#f3f4f6",border:"1px solid #d1d5db",borderRadius:6,cursor:"pointer",fontFamily:"inherit"}}>ביטול</button>
              <button type="button" onClick={handleSaveNewContact} style={{flex:2,padding:8,fontSize:12,background:"#0d9488",color:"#fff",border:"none",borderRadius:6,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>💾 שמור והשתמש</button>
            </div>}
          </div>}
        </div>
`.replace(/<\/?motion\.motion.div/g, (m) => m.replace("motion.", "")).replace(/<\/?motion\.div/g, (m) => m.replace("motion.", ""));

s = s.slice(0, gridStart) + newUi + s.slice(endIdx);
fs.writeFileSync(indexPath, s, "utf8");
console.log("OK");

const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const indexPath = path.join(root, "index.html");

let fn = fs.readFileSync(path.join(root, "scripts/email-fn-raw.js"), "utf8");
fn = fn.split("motion.div").join("motion.div"); // noop safety
fn = fn.split("motion.div").join("div");

let html = fs.readFileSync(indexPath, "utf8");
const s = html.indexOf("function buildClientNotificationEmail(");
const e = html.indexOf("async function loadClientForNotification(");
html = html.slice(0, s) + fn + "\n" + html.slice(e);

html = html.replace(
  '  const [body,setBody]=useState("");',
  '  const [emailHtml,setEmailHtml]=useState("");\n  const [emailText,setEmailText]=useState("");'
);
html = html.replace(
  "    const {subject:subj,body:bodyText}=buildClientNotificationEmail({task,itemsList:taskItems,clientData,notificationType});\n    setSubject(subj);setBody(bodyText);",
  "    const {subject:subj,html:htmlOut,text:textOut}=buildClientNotificationEmail({task,itemsList:taskItems,clientData,notificationType});\n    setSubject(subj);setEmailHtml(htmlOut);setEmailText(textOut);"
);
html = html.replace(
  "      await sendEmail({to:selectedEmail,cc:ccEmails,subject,text:body});",
  "      await sendEmail({to:selectedEmail,cc:ccEmails,subject,html:emailHtml,text:emailText});"
);
html = html.replace(/body,sent_by:user/g, "body:emailText,sent_by:user");
html = html.replace(
  "const sendEmail=async({to,subject,html,text,cc})=>{\n  const payload={to,subject,html,text};",
  "const sendEmail=async({to,subject,html,text,cc,from})=>{\n  const payload={to,subject,html,text};\n  if(from)payload.from=from;"
);

const oldPreview =
  '<motion.div style={{marginBottom:15}}><label style={{display:"block",marginBottom:4,fontWeight:600}}>תוכן:</label><textarea value={body} onChange={e=>setBody(e.target.value)} style={{...fld,minHeight:280,resize:"vertical"}}/></motion.div>';
const newPreview =
  '<motion.div style={{marginBottom:15}}><label style={{display:"block",marginBottom:4,fontWeight:600}}>תצוגה מקדימה:</label><iframe srcDoc={emailHtml} style={{width:"100%",height:400,border:"1px solid var(--border)",borderRadius:10,background:"#fff"}} title="email preview"/><details style={{marginTop:8}}><summary style={{cursor:"pointer",fontSize:12,color:"var(--text3)"}}>ערוך תוכן (טקסט בלבד)</summary><textarea value={emailText} onChange={e=>setEmailText(e.target.value)} style={{...fld,minHeight:200,resize:"vertical",marginTop:8,direction:"rtl",textAlign:"right"}}/></details></motion.div>';
const oldPreview2 = oldPreview.split("motion.div").join("div");
const newPreview2 = newPreview.split("motion.div").join("div");
if (html.includes(oldPreview2)) {
  html = html.replace(oldPreview2, newPreview2);
} else {
  console.warn("preview block not found - trying alternate");
  html = html.replace(
    /value=\{body\} onChange=\{e=>setBody/g,
    "value={emailText} onChange={e=>setEmailText"
  );
}

html = html.split("motion.div").join("div");

fs.writeFileSync(indexPath, html, "utf8");
console.log("OK patched", !html.includes("motion.div"), html.includes("emailHtml"));

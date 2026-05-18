import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const indexPath = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "index.html");
let c = fs.readFileSync(indexPath, "utf8");

function mustReplace(from, to, label) {
  if (!c.includes(from)) throw new Error(`${label}: block not found`);
  c = c.replace(from, to);
}

mustReplace(
  `      {driver&&<div style={{fontSize:10,color:driver.color,marginTop:4,fontWeight:600}}>🚛 {driver.name}</motion.div>}`,
  `      {driver&&<div style={{fontSize:10,color:isDelivered?textMuted:driver.color,marginTop:4,fontWeight:600}}>🚛 {driver.name}</div>}`,
  "driver"
);

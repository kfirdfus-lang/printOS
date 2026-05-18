# -*- coding: utf-8 -*-
import pathlib

p = pathlib.Path(__file__).resolve().parent.parent / "index.html"
c = p.read_text(encoding="utf-8")

def must_replace(old, new, label):
    global c
    if old not in c:
        raise SystemExit(f"{label}: not found\n---\n{old[:120]}")
    c = c.replace(old, new, 1)

must_replace(
    '      {!compact&&items.length>0&&<div style={{marginTop:6,paddingTop:6,borderTop:"1px dashed #e5e7eb",fontSize:10,color:"#374151"}}>\n        <div style={{fontWeight:600,marginBottom:3,color:"var(--text3)"}}>📦 פריטים למסירה:</div>',
    '      {!compact&&items.length>0&&<motion.div style={{marginTop:6,paddingTop:6,borderTop:"1px dashed #e5e7eb",fontSize:10,color:isDelivered?"#9ca3af":"#374151"}}>\n        <div style={{fontWeight:600,marginBottom:3,color:textMuted}}>📦 פריטים למסירה:</div>',
    "items",
)

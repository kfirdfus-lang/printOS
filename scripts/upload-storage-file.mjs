import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const SU = 'https://pvwcpukfhyrmdpxgfwrk.supabase.co'
const SA = process.env.SUPABASE_ANON_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InB2d2NwdWtmaHlybWRweGdmd3JrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY4NzIzNzMsImV4cCI6MjA5MjQ0ODM3M30.69GpTOQdwtAgouGN0MIHqQc6bImxAbfQ72xL2S4KaoM'
const BUCKET = 'printos-files'

const [localPath, storagePath] = process.argv.slice(2)
if (!localPath || !storagePath) {
  console.error('Usage: node scripts/upload-storage-file.mjs <local> <storage-path>')
  process.exit(1)
}

const abs = path.isAbsolute(localPath) ? localPath : path.join(process.cwd(), localPath)
const buf = fs.readFileSync(abs)
const ext = path.extname(abs).toLowerCase()
const types = { '.png': 'image/png', '.pdf': 'application/pdf', '.ttf': 'font/ttf' }
const contentType = types[ext] || 'application/octet-stream'

const res = await fetch(`${SU}/storage/v1/object/${BUCKET}/${storagePath}`, {
  method: 'POST',
  headers: {
    apikey: SA,
    Authorization: `Bearer ${SA}`,
    'Content-Type': contentType,
    'x-upsert': 'true',
  },
  body: buf,
})

const text = await res.text()
if (!res.ok) {
  console.error('Upload failed:', res.status, text)
  process.exit(1)
}
console.log('OK:', storagePath, buf.length, 'bytes')

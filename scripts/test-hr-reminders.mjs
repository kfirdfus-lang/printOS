// Manual test for the hr-reminders-check Edge Function.
// Default: dry run (no email). Pass --send to actually send the email.
const SUPABASE_URL = 'https://pvwcpukfhyrmdpxgfwrk.supabase.co'
const ANON_KEY =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InB2d2NwdWtmaHlybWRweGdmd3JrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY4NzIzNzMsImV4cCI6MjA5MjQ0ODM3M30.69GpTOQdwtAgouGN0MIHqQc6bImxAbfQ72xL2S4KaoM'

const dryRun = !process.argv.includes('--send')
const res = await fetch(`${SUPABASE_URL}/functions/v1/hr-reminders-check`, {
  method: 'POST',
  headers: {
    apikey: ANON_KEY,
    Authorization: `Bearer ${ANON_KEY}`,
    'Content-Type': 'application/json',
  },
  body: JSON.stringify({ dry_run: dryRun }),
})
console.log('status:', res.status, dryRun ? '(dry run - no email)' : '(real send)')
console.log(JSON.stringify(await res.json(), null, 2))

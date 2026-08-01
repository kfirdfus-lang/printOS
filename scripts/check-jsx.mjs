// Dev helper: extracts the in-browser Babel scripts from index.html and
// verifies they parse as JSX (same as the browser Babel would).
import { readFileSync, writeFileSync, mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { execFileSync } from 'node:child_process'

const html = readFileSync('index.html', 'utf8')
const scripts = [...html.matchAll(/<script type="text\/babel"[^>]*>([\s\S]*?)<\/script>/g)]
console.log('babel scripts found:', scripts.length)

const dir = mkdtempSync(path.join(tmpdir(), 'jsxcheck-'))
let failed = false
scripts.forEach((s, i) => {
  const file = path.join(dir, `part${i}.jsx`)
  writeFileSync(file, s[1])
  try {
    execFileSync('npx', ['--yes', 'esbuild', file, '--loader:.jsx=jsx', '--outfile=' + file + '.out.js'], {
      stdio: 'pipe',
      shell: true,
    })
    console.log(`script ${i}: OK (${s[1].length} chars)`)
  } catch (e) {
    failed = true
    console.error(`script ${i}: PARSE FAILED`)
    console.error(String(e.stderr || e.message))
  }
})
process.exit(failed ? 1 : 0)

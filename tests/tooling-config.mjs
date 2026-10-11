import assert from 'node:assert/strict'
import test from 'node:test'
import { ESLint } from 'eslint'

const eslint = new ESLint()
const messages = async (filePath, code) => {
  const [result] = await eslint.lintText(code, { filePath })
  assert.ok(result, `${filePath} must be processed`)
  assert.equal(result.messages.some(message => /ignored/i.test(message.message)), false)
  return result.messages
}

test('legacy JavaScript keeps ES3 syntax and shared browser/Node globals', async () => {
  assert.deepEqual(await messages('probe.js', 'var value = 1; console.log(value);'), [])
  assert.deepEqual(await messages('probe.js', 'console.log(JSON.stringify({safe: true}));'), [])
  assert.ok((await messages('probe.js', 'const value = 1;')).some(message => message.fatal))
})

test('JavaScript still reports undefined globals and missing semicolons', async () => {
  assert.ok((await messages('probe.js', 'missingGlobal();')).some(message => message.ruleId === 'no-undef'))
  assert.ok((await messages('probe.js', 'console.log(1)')).some(message => message.ruleId === 'semi'))
})

test('existing correctness rules still reject unreachable statements', async () => {
  assert.ok((await messages('probe.js', 'function example() { return 1; console.log(2); }'))
    .some(message => message.ruleId === 'no-unreachable' && message.severity === 2))
})

test('the CLI script extension is processed', async () => {
  assert.deepEqual(await messages('probe.njs', 'console.log(1);'), [])
  assert.deepEqual(await messages('bin/probe.njs', 'console.log(process.argv, require("fs"));'), [])
  assert.ok((await messages('probe.njs', 'missingGlobal();')).some(message => message.ruleId === 'no-undef'))
})

test('JSON accepts valid data and rejects invalid syntax', async () => {
  assert.deepEqual(await messages('probe.json', '{"safe": true}'), [])
  assert.ok((await messages('probe.json', '{ invalid }')).some(message => message.severity === 2))
})

test('HTML script linting remains active for both extensions', async () => {
  for (const extension of ['html', 'htm']) {
    assert.deepEqual(await messages(`probe.${extension}`, '<script>console.log(1);</script>'), [])
    assert.deepEqual(await messages(`probe.${extension}`, '<script>console.log(document.body, new FileReader());</script>'), [])
    assert.ok((await messages(`probe.${extension}`, '<script>missingGlobal();</script>'))
      .some(message => message.ruleId === 'no-undef'))
  }
})

test('the generated CJS bundle retains its existing exclusion', async () => {
  assert.equal(await eslint.isPathIgnored('xlsx.js'), true)
})

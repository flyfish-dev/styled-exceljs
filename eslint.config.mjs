import { readFileSync } from 'node:fs'
import { FlatCompat } from '@eslint/eslintrc'
import json from 'eslint-plugin-json'

const compat = new FlatCompat({
  baseDirectory: import.meta.dirname,
})

// The project keeps its original ES3 rule policy while using current ESLint.
// .eslintrc explicitly retains the previous recommended rules and overrides.
export default [
  { ignores: ['**/xlsx.js'] },
  ...compat.config(
    JSON.parse(readFileSync(new URL('.eslintrc', import.meta.url), 'utf8')),
  ),
  {
    files: ['**/*.{js,njs,json,html,htm}'],
    languageOptions: { sourceType: 'script', globals: { JSON: 'readonly' } },
  },
  ...compat.config({ env: { node: true } })
    .map(config => ({ ...config, files: ['bin/**/*.njs'] })),
  ...compat.config({ env: { browser: true } })
    .map(config => ({ ...config, files: ['**/*.{html,htm}'] })),
  {
    files: ['**/*.json'],
    ...json.configs.recommended,
  },
]

const assert = require('node:assert/strict');

const actual = process.versions.bun ? 'bun' : process.versions.deno ? 'deno' : 'node';
assert.equal(actual, process.env.STYLED_TEST_RUNTIME, 'The corpus must execute in the requested runtime');
console.log('Corpus runtime: ' + actual + ' ' + process.versions[actual]);

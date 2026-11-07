// MatOffice Sheets – the function registry is complete: every function has a category, its syntax and a description in English and German
// (Resources/Texts/JsSheetsFunctions.json, shown by the Insert Function dialog).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { FUNCTIONS, CATEGORIES } from '../../../src/MatOffice/wwwroot/js/sheets/engine/functions/index.js';

const file = new URL('../../../src/MatOffice/Resources/Texts/JsSheetsFunctions.json', import.meta.url);
const raw = readFileSync(file, 'utf8');
const texts = JSON.parse(raw);
const PREFIX = 'Js.Sheets.Fn.';

test('every function has a category, a valid argument count and its syntax', () => {
  assert.ok(FUNCTIONS.size >= 300, `only ${FUNCTIONS.size} functions`);
  for (const [name, spec] of FUNCTIONS) {
    assert.equal(spec.name, name);
    assert.ok(CATEGORIES.includes(spec.category), `${name}: category ${spec.category}`);
    assert.ok(Number.isInteger(spec.min) && spec.min >= 0 && spec.min <= spec.max && spec.max <= 255, `${name}: ${spec.min} … ${spec.max}`);
    assert.equal(typeof spec.args, 'string', `${name}: args`);
    assert.equal(typeof spec.fn, 'function', `${name}: fn`);
    assert.ok(/^[A-Z][A-Z0-9.]*$/.test(name), `${name}: name`);
  }
});

test('the volatile and lazy functions are marked', () => {
  for (const name of ['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'OFFSET', 'INDIRECT']) assert.equal(FUNCTIONS.get(name).volatile, true, name);
  for (const name of ['IF', 'IFS', 'IFERROR', 'IFNA', 'SWITCH', 'CHOOSE']) assert.equal(FUNCTIONS.get(name).lazy, true, name);
  assert.equal(FUNCTIONS.get('SUM').volatile, false);
});

test('every function has a description in English and German, and every description a function', () => {
  for (const name of FUNCTIONS.keys()) {
    const entry = texts[PREFIX + name];
    assert.ok(entry, `no description for ${name}`);
    for (const lang of ['en', 'de']) {
      assert.equal(typeof entry[lang], 'string', `${name}.${lang}`);
      assert.ok(entry[lang].trim().length > 10, `${name}.${lang} is empty`);
      assert.ok(!/\{\d+\}/.test(entry[lang]), `${name}.${lang} has a placeholder`);
    }
  }
  for (const key of Object.keys(texts)) {
    assert.ok(key.startsWith(PREFIX) && FUNCTIONS.has(key.slice(PREFIX.length)), `${key} is no function`);
  }
});

test('the description file is UTF-8 without BOM, one entry per line, sorted by name', () => {
  assert.notEqual(raw.charCodeAt(0), 0xfeff);
  const lines = raw.trim().split('\n').slice(1, -1);
  assert.equal(lines.length, Object.keys(texts).length);
  const keys = lines.map(l => JSON.parse('{' + l.replace(/,$/, '') + '}')).map(o => Object.keys(o)[0]);
  assert.deepEqual(keys, [...keys].sort());
});

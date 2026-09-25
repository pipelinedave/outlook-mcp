import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { findTdzSites, findTdzInSource } from './tdz-guard.mjs';

const fmt = (hits) => hits.map((h) => `${h.name}@Z${h.useLine} (deklariert Z${h.declLine})`);

describe('TDZ-Guard — kein Statement vor seiner eigenen const-Deklaration', () => {

  test('owa.js: 0 TDZ-Stellen', () => {
    // fc943cd hatte hier zwei Instanzen in getFreeBusy:
    //   rawLabels/eventElementsFound/selectorUsed vor `const eventCollection`
    //   und der zerrissene `let note = '...' +`-String (ESM-SyntaxError).
    assert.deepEqual(fmt(findTdzSites('owa.js')), []);
  });

  test('browser.js: 0 TDZ-Stellen', () => {
    assert.deepEqual(fmt(findTdzSites('browser.js')), []);
  });

  // META-GUERTS: die Fehlerklasse muss erkannt werden. Ein synthetisches
  // Modul in Live-Fehlerform MUSS als Treffer auffallen — sonst ist der
  // Wächter blind und meldet vorsichtig falsch "gruen".
  test('Der Guard erkennt die Live-Fehlerform (Meta-Guard)', () => {
    const buggy = [
      'export async function f() {',
      '  const a = later;',
      '  const later = 5;',
      '  return a;',
      '}'
    ].join('\n');
    const hits = findTdzInSource(buggy);
    assert.equal(hits.length, 1, `hits=${fmt(hits).join(',')}`);
    assert.equal(hits[0].name, 'later');
    assert.ok(hits[0].useLine < hits[0].declLine);
  });

  test('Der Guard erkennt die dreifache eventCollection-Variante', () => {
    const buggy = [
      'export async function g() {',
      '  const rawLabels = eventCollection.labels;',
      '  const eventElementsFound = eventCollection.eventElementsFound;',
      '  const selectorUsed = eventCollection.selectorUsed;',
      '  const eventCollection = await collect(page);',
      '  return [rawLabels, eventElementsFound, selectorUsed];',
      '}'
    ].join('\n');
    const hits = findTdzInSource(buggy);
    assert.equal(hits.length, 3, `hits=${fmt(hits).join(',')}`);
    assert.ok(hits.every((h) => h.name === 'eventCollection'));
  });

  test('Korrekte Reihenfolge erzeugt keinen Treffer', () => {
    const ok = [
      'export async function h() {',
      '  const eventCollection = await collect(page);',
      '  const rawLabels = eventCollection.labels;',
      '  return rawLabels;',
      '}'
    ].join('\n');
    assert.deepEqual(findTdzInSource(ok), []);
  });

  test('Verschiedene Bloecke sind kein TDZ (verschiedene Tiefen)', () => {
    const ok = [
      'export function i() {',
      '  if (x) { const inner = 1; use(inner); }',
      '  const outer = 2;',
      '  return outer;',
      '}'
    ].join('\n');
    assert.deepEqual(findTdzInSource(ok), []);
  });

  test('Verschiedene Funktionen sind kein TDZ', () => {
    const ok = [
      'export function a() { return helper(); }',
      'export function b() { const helper = () => 1; return helper; }'
    ].join('\n');
    assert.deepEqual(findTdzInSource(ok), []);
  });

  test('Lambda-Parameter schatten (muss KEINen Treffer ergeben)', () => {
    // Genau die False-Positive-Quelle, die im echten Code zuerst anschlaegt:
    //   .filter((t) => t && t.length <= 300);
    //   for (const t of texts) { ... }
    const ok = [
      'export function j(texts) {',
      '  const texts2 = [1, 2].map((n) => n).filter((t) => t && t <= 2);',
      '  for (const t of texts) { use(t); }',
      '  return texts2;',
      '}'
    ].join('\n');
    assert.deepEqual(findTdzInSource(ok), []);
  });
});

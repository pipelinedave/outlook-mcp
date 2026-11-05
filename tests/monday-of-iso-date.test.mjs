// Unit-Tests: mondayOfIsoDate — Montag der Woche, die ein ISO-Datum enthält.
// Kontext: Tool-Lücke 05.11.2026 (gbrain #10040) — get/delete/find Kalender-Events
// fanden nur Events in der AKTUELLEN Woche. Der Fix navigiert verifiziert zur
// Zielwoche; der Montag der Zielwoche ist der Anker dafür.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mondayOfIsoDate } from '../src/owa.js';

test('Mitte der Woche -> Montag derselben Woche', () => {
  // 2026-11-19 ist ein Donnerstag; Woche: Mo 16.11. – So 22.11.
  assert.equal(mondayOfIsoDate('2026-11-19'), '2026-11-16');
});

test('Montag selbst -> unverändert', () => {
  assert.equal(mondayOfIsoDate('2026-11-16'), '2026-11-16');
});

test('Sonntag -> Montag der LAUFENDEN Woche (OWA beginnt montags)', () => {
  assert.equal(mondayOfIsoDate('2026-11-22'), '2026-11-16');
});

test('Zeitkomponente wird ignoriert', () => {
  assert.equal(mondayOfIsoDate('2026-11-17T10:00:00'), '2026-11-16');
});

test('Monatsgrenze: 1. eines Monats mitten in der Woche', () => {
  // 2026-12-01 ist ein Dienstag; Woche: Mo 30.11. – So 06.12.
  assert.equal(mondayOfIsoDate('2026-12-01'), '2026-11-30');
});

test('Jahresgrenze: 1. Januar 2026 (Donnerstag) -> 29.12.2025', () => {
  assert.equal(mondayOfIsoDate('2026-01-01'), '2025-12-29');
});

test('Ungültige Eingaben -> null', () => {
  assert.equal(mondayOfIsoDate(''), null);
  assert.equal(mondayOfIsoDate(null), null);
  assert.equal(mondayOfIsoDate('19.11.2026'), null);
  assert.equal(mondayOfIsoDate('2026-13-99'), null);
});

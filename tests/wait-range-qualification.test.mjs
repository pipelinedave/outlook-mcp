import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
  waitForCalendarEvents,
  makeEventObserver,
  collectCalendarEventLabels,
  eventWaitObservationState,
  eventWaitWarning,
  reconcileEventWait,
  resolveBusySlotsReliable,
  resolveDataQuality,
  isBelievableBusySlots,
  slotExtractionWarning,
  eventWaitBudgetMs,
  EVENT_WAIT_STATES,
  EVENT_WAIT_ABORT_REASONS,
  FREEBUSY_EVENT_WAIT_POLL_MS,
  FREEBUSY_EVENT_WAIT_SETTLE_MS
} from '../src/owa.js';

// ===========================================================================
// FIX 2 (Live-Lauf 5, 26.09.2026): 'found' war DATUMS-qualifiziert, nicht
// BEREICHS-qualifiziert.
//
// Live-Beleg: eventWaitDateQualified: true bei eventWaitInRangeLabels: 0
// -> found nach 41ms von 20000ms Budget. Ursache: der als Header gelesene
// Knoten war der "Go to next week"-Button, dessen Label die ZIELWOCHE nannte
// (FIX 1). Ein datierter Header der FALSCHEN Woche ist ein GEGENBELEG — er
// sagt nicht "nichts gefunden", sondern "es ist nachweislich die falsche
// Woche". Er darf weder 'found' noch den Leer-Beweis ausloesen und bekommt
// darum einen eigenen, benannten Zustand: header-outside-requested-range.
//
// Diese Tests fahren den ECHTEN Draht: makeEventObserver (Produktion) ->
// collectCalendarEventLabels 'fast' (Produktion, echter evaluate-Callback
// gegen ein Fake-DOM) -> waitForCalendarEvents (Produktion).
// ===========================================================================

const virtualClock = (start = 1_000_000) => {
  let t = start;
  const sleeps = [];
  return { now: () => t, sleep: async (ms) => { sleeps.push(ms); t += ms; }, get slept() { return sleeps; } };
};

const node = (over = {}) => ({
  tagName: over.tag || 'DIV',
  getAttribute: (n) => (over.attrs && over.attrs[n] !== undefined ? over.attrs[n] : null),
  getBoundingClientRect: () => over.rect || { left: 0, top: 0, width: 120, height: 24, x: 0, y: 0 },
  offsetParent: null,
  innerText: over.text || '',
  className: over.className || '',
  attributes: Object.entries(over.attrs || {}).map(([name, value]) => ({ name, value })),
  children: []
});

const matches = (el, sel) => {
  const s = String(sel).trim();
  if (s === '*') return true;
  const tag = /^([a-zA-Z][\w-]*)$/.exec(s);
  if (tag) return String(el.tagName || '').toLowerCase() === tag[1].toLowerCase();
  const attrSub = /\[([\w-]+)\*=\s*"([^"]*)"(\s*i)?\]/gi;
  let m;
  const subs = [];
  attrSub.lastIndex = 0;
  while ((m = attrSub.exec(s)) !== null) subs.push({ name: m[1], value: m[2], ci: Boolean(m[3]) });
  for (const sub of subs) {
    const actual = el.getAttribute ? el.getAttribute(sub.name) : null;
    if (actual == null) return false;
    const hay = sub.ci ? String(actual).toLowerCase() : String(actual);
    const nee = sub.ci ? sub.value.toLowerCase() : sub.value;
    if (hay.indexOf(nee) < 0) return false;
  }
  if (subs.length) return true;
  const attrEq = /\[([\w-]+)\s*=\s*"([^"]*)"(\s*i)?\]/gi;
  attrEq.lastIndex = 0;
  let allEq = true;
  let anyEq = false;
  while ((m = attrEq.exec(s)) !== null) {
    anyEq = true;
    const [, name, value, ci] = m;
    const actual = el.getAttribute ? el.getAttribute(name) : null;
    if (actual == null || (ci ? String(actual).toLowerCase() !== value.toLowerCase() : String(actual) !== value)) {
      allEq = false;
      break;
    }
  }
  if (anyEq) return allEq;
  const attrHas = /\[([\w-]+)\]/g;
  attrHas.lastIndex = 0;
  while ((m = attrHas.exec(s)) !== null) {
    if (el.getAttribute && el.getAttribute(m[1]) != null) return true;
  }
  return false;
};

const mutableDoc = (nodes) => {
  let current = nodes;
  return {
    set(next) { current = next; },
    get() { return current; },
    querySelectorAll: (sel) => {
      if (sel === null) return current;
      const parts = String(sel).split(',').map((s) => s.trim());
      return current.filter((el) => parts.some((p) => matches(el, p)));
    },
    querySelector: () => null,
    body: null
  };
};

const fakePage = { evaluate: async (fn, arg) => fn(arg) };
const withFakeDocument = async (doc, fn) => {
  const prev = globalThis.document;
  globalThis.document = doc;
  try { return await fn(); } finally { globalThis.document = prev; }
};

const TARGET_WEEK_START = '2026-09-28';   // angefragt: KW 40
const TEST_RANGE = {
  rangeStartMs: new Date('2026-09-28T00:00:00').getTime(),
  rangeEndMs: new Date('2026-10-05T00:00:00').getTime()
};
// Der Header, den die View WIRKLICH zeigt (KW 39) — nach FIX 1 der next-Button.
const DISPLAYED_HEADER = { text: '21–27 September, 2026', startDate: '2026-09-21', endDate: '2026-09-27' };
// Ein echter Event-Chip der FALSCHEN Woche (KW 39) — der Live-Zustand.
const WRONG_WEEK_CHIP = { attrs: { 'aria-label': 'MKK Nachbereitung Update Prod, 08:00 to 10:00, Monday September 21 2026, Busy' } };
// Ein Event-Chip der ZIELWOCHE.
const TARGET_CHIP = { attrs: { 'aria-label': 'TeamSync, 10:00 to 11:00, Monday, September 28, 2026, Busy' } };

const runWait = (doc, opts = {}) => withFakeDocument(doc, () => waitForCalendarEvents({
  observe: makeEventObserver(opts.collect || (() =>
    collectCalendarEventLabels(fakePage, { weekStartIso: TARGET_WEEK_START, fast: true })), {
    readWeek: opts.readWeek || null,
    range: opts.range === null ? null : (opts.range || TEST_RANGE),
    defaultYear: 2026,
    // Der Beobachter braucht DIE GLEICHE Uhr wie die Wartephase: sein
    // Header-Read ist zeitlich gedrosselt (headerReadIntervalMs). Mit der
    // echten Uhr wuerde der erste Header fuer den ganzen Testlauf
    // eingefroren — der Kalender kann aber zwischen den Polls umschalten.
    now: opts.clock ? opts.clock.now : undefined
  }),
  budgetMs: opts.budgetMs === undefined ? 20000 : opts.budgetMs,
  pollMs: FREEBUSY_EVENT_WAIT_POLL_MS,
  settleMs: FREEBUSY_EVENT_WAIT_SETTLE_MS,
  targetStart: opts.targetStart === undefined ? TARGET_WEEK_START : opts.targetStart,
  now: opts.clock ? opts.clock.now : virtualClock().now,
  sleep: opts.clock ? opts.clock.sleep : virtualClock().sleep
}));

// ===========================================================================
describe('FIX 4 — Invariante: Widerspruchswarnung => busySlotsReliable === false', () => {

  test('der Live-Fall aus Lauf 2: 18 Elemente, 0 Slots, dataQuality verified', () => {
    const dataQuality = resolveDataQuality({ weekVerified: true, eventElementsFound: 18 });
    assert.equal(dataQuality, 'verified');
    assert.ok(slotExtractionWarning(18, 0), 'der Widerspruchstext entsteht …');
    assert.equal(resolveBusySlotsReliable({ dataQuality, eventElementsFound: 18, busySlotCount: 0 }), false,
      '… und busySlotsReliable muss dann false sein');
  });

  test('die Invariante gilt fuer JEDE Elementzahl', () => {
    for (const n of [1, 6, 18, 120]) {
      const dq = resolveDataQuality({ weekVerified: true, eventElementsFound: n });
      assert.equal(resolveBusySlotsReliable({ dataQuality: dq, eventElementsFound: n, busySlotCount: 0 }), false,
        `${n} Elemente / 0 Slots muss unzuverlaessig sein`);
    }
  });

  test('ohne Widerspruch bleibt es true (der Fix darf nichts ueberkorrigieren)', () => {
    const dq = resolveDataQuality({ weekVerified: true, eventElementsFound: 6 });
    assert.equal(resolveBusySlotsReliable({ dataQuality: dq, eventElementsFound: 6, busySlotCount: 3 }), true);
    // dataQuality unzuverlaessig -> unabhaengig vom Widerspruch false.
    for (const dq of ['unverified', 'unverified-empty']) {
      assert.equal(resolveBusySlotsReliable({ dataQuality: dq, eventElementsFound: 6, busySlotCount: 3 }), false);
    }
  });

  test('der Fall "0 Elemente, 0 Slots" ist KEIN Widerspruch (leere Woche)', () => {
    assert.equal(slotExtractionWarning(0, 0), null);
    assert.equal(resolveBusySlotsReliable({ dataQuality: 'unverified-empty', eventElementsFound: 0, busySlotCount: 0 }), false);
  });

  test('der Aufrufer benutzt die Invariante, nicht die alte Ein-Zeilen-Form', () => {
    // Verdrahtungswächter: die Funktion kann korrekt sein und der Aufrufer
    // trotzdem die alte Form benutzen (genau der Fehler in Lauf 2).
    const src = readFileSync(fileURLToPath(new URL('../src/owa.js', import.meta.url)), 'utf8');
    assert.ok(!/busySlotsReliable: isBelievableBusySlots\(dataQuality\)/.test(src),
      'die widerspruchsfreie Ein-Zeilen-Form darf nicht zurueckkehren');
    assert.ok(/const busySlotsReliable = resolveBusySlotsReliable\(\{/.test(src),
      'das Ergebnisfeld muss aus resolveBusySlotsReliable kommen');
    // Und die Warnung wird EINMAL berechnet und fuer beides benutzt.
    const iWarn = src.indexOf('const extractionWarning = slotExtractionWarning(eventElementsFound, busySlots.length);');
    const iReliable = src.indexOf('const busySlotsReliable = resolveBusySlotsReliable({');
    const iResult = src.indexOf('busySlotsReliable,', iReliable);
    const iUse = src.indexOf('extractionWarning,', iResult);
    assert.ok(iWarn >= 0 && iReliable > iWarn, 'die Warnung wird VOR dem Ergebnis berechnet');
    assert.ok(iResult > iReliable, 'busySlotsReliable steht im Ergebnis');
    assert.ok(iUse > iResult, 'dieselbe Warnung wird im Diagnostik-Block wiederverwendet');
    assert.equal((src.match(/slotExtractionWarning\(eventElementsFound, busySlots\.length\)/g) || []).length, 1,
      'genau eine Berechnung (zwei Stellen koennten auseinanderlaufen)');
  });

  test('der Widerspruch erzeugt zwingend Diagnostik (nicht nur eine Warnung im Text)', () => {
    const src = readFileSync(fileURLToPath(new URL('../src/owa.js', import.meta.url)), 'utf8');
    // Bei busySlotsReliable === false wird der Diagnostik-Block angehaengt —
    // im Widerspruchsfall also immer.
    assert.ok(/if \(!result\.busySlotsReliable \|\| busySlots\.length === 0 \|\| eventWait\.outcome !== 'found'\) \{/.test(src));
  });
});

// ===========================================================================
// Verdrahtungswächter für FIX 2
// ===========================================================================

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
describe('FIX 2 — Header der FALSCHEN Woche ist ein Gegenbeleg, kein Beleg', () => {

  test('KONTROLLFALL (Live): Header KW 39 + 10 Events ausserhalb -> KEIN found, Poll bis Budget', async () => {
    // Der Zustand, den Lauf 5 mit 'found' nach 41ms verlassen hat: 0
    // In-Range-Labels, ein DATIERTER Header — aber der der falschen Woche.
    const doc = mutableDoc([node(WRONG_WEEK_CHIP), node(WRONG_WEEK_CHIP), node(WRONG_WEEK_CHIP)]);
    const clock = virtualClock();
    const res = await runWait(doc, { clock, readWeek: async () => DISPLAYED_HEADER });
    assert.equal(res.found, false, 'FIX 2: ein Header der falschen Woche ist kein Beleg');
    assert.equal(res.dateQualified, false, 'BUG 4/5-Kennzahl: dateQualified muss false sein');
    assert.notEqual(res.outcome, 'found');
    // Die 41ms aus Lauf 5 sind der Kern des Befunds.
    assert.ok(res.polls > 30, `es wurde bis zum Budget gepollt (war ${res.polls} Polls)`);
    assert.ok(res.waitedMs >= 20000 - FREEBUSY_EVENT_WAIT_POLL_MS,
      `das Budget wurde ausge schoepft (war ${res.waitedMs}ms, Live-Lauf 5: 41ms)`);
    // Der Zustand ist BENANNT, nicht in 'events-outside-range' aufgeloest.
    assert.equal(res.abortReason, EVENT_WAIT_ABORT_REASONS.headerOutsideRange);
    assert.equal(res.abortReason, 'header-outside-requested-range');
    assert.notEqual(res.abortReason, EVENT_WAIT_ABORT_REASONS.eventsOutsideRange,
      'der Header traegt die staerkere Information und wird zuerst genannt');
    assert.notEqual(res.abortReason, EVENT_WAIT_ABORT_REASONS.renderedEmpty,
      'ein Header, der die falsche Woche nennt, ist per Definition kein Leer-Beweis');
    assert.equal(res.lastState, EVENT_WAIT_STATES.headerOutsideRange);
    // Der eigene Zaehler: er macht den Gegenbeleg im Ergebnis sichtbar und
    // trennt ihn von "Events ausserhalb des Zeitraums" (anderer Zaehler).
    assert.ok(res.headerOutsideRangePolls > 30, `jeder Poll sah den Header-Gegenbeleg (war ${res.headerOutsideRangePolls})`);
  });

  test('KONTROLLFALL: Header der falschen Woche, 0 Events -> auch KEIN Leer-Beweis', async () => {
    // Ohne Events wuerde der Leer-Beweis 'rendered-empty-proven' greifen
    // ("Kalender nachweislich leer"). Der Header sagt aber das Gegenteil:
    // die Ansicht zeigt KW 39. Das ist die schlimmste moegliche Falschaussage.
    const doc = mutableDoc([]);
    const clock = virtualClock();
    const res = await runWait(doc, { clock, readWeek: async () => DISPLAYED_HEADER });
    assert.equal(res.found, false);
    assert.notEqual(res.abortReason, EVENT_WAIT_ABORT_REASONS.renderedEmpty,
      '"leer" waere eine Falschaussage — der Header widerspricht ihr');
    assert.equal(res.abortReason, EVENT_WAIT_ABORT_REASONS.headerOutsideRange);
    assert.equal(res.outcome, 'budget-exhausted', 'bis zum Budget weitergepollt');
    assert.ok(res.waitedMs >= 20000 - FREEBUSY_EVENT_WAIT_POLL_MS, `Budget ausge schoepft (war ${res.waitedMs}ms)`);
  });

  test('der Zustand loest das Ruhe-Fenster NICHT aus (kein Leer-Beweis aus settleMs)', async () => {
    // Ohne diesen Zweig wuerde nach settleMs 'rendered-empty-proven' gemeldet.
    const doc = mutableDoc([]);
    const clock = virtualClock();
    const res = await runWait(doc, { clock, readWeek: async () => DISPLAYED_HEADER });
    assert.ok(res.waitedMs > FREEBUSY_EVENT_WAIT_SETTLE_MS,
      `das Ruhe-Fenster (${FREEBUSY_EVENT_WAIT_SETTLE_MS}ms) wurde nicht als Leer-Beweis gewertet (war ${res.waitedMs}ms)`);
  });

  test('Header der ZIELWOCHE qualifiziert weiterhin (der Fall darf nicht kaputtgehen)', async () => {
    const doc = mutableDoc([]);
    const clock = virtualClock();
    const res = await runWait(doc, {
      clock,
      readWeek: async () => ({ text: '28 September – 4 October, 2026', startDate: '2026-09-28', endDate: '2026-10-04' })
    });
    assert.equal(res.found, true);
    assert.equal(res.evidence, 'week-header');
    assert.equal(res.dateQualified, true);
    assert.equal(res.abortReason, null);
  });

  test('In-Range-Label qualifiziert auch bei Header der falschen Woche (starkeres Label)', async () => {
    // Der Header ist ein Gegenbeleg, aber ein Label AUS dem angefragten
    // Zeitraum ist ein direkter Beleg — er darf nicht unterdrueckt werden.
    const doc = mutableDoc([node(TARGET_CHIP)]);
    const clock = virtualClock();
    const res = await runWait(doc, { clock, readWeek: async () => DISPLAYED_HEADER });
    assert.equal(res.found, true, 'ein Label aus dem angefragten Zeitraum ist ein Beleg');
    assert.equal(res.evidence, 'in-range-label');
  });

  test('der Spinner schlaegt den Header-Gegenbeleg (noch nicht fertig ist noch nicht fertig)', async () => {
    const spinner = node({ tag: 'MS-SPINNER', className: 'ms-Spinner ms-Spinner-circle', attrs: { class: 'ms-Spinner ms-Spinner-circle' } });
    const doc = mutableDoc([node(WRONG_WEEK_CHIP), spinner]);
    const clock = virtualClock();
    const res = await runWait(doc, { clock, readWeek: async () => DISPLAYED_HEADER });
    assert.equal(res.found, false);
    assert.equal(res.lastState, EVENT_WAIT_STATES.stillLoading,
      'ein Spinner ist "laedt noch" und traegt die staerkere Aussage');
  });

  test('UMSCHALTEN: Header falsch, dann Header der Zielwoche -> found erst dann', async () => {
    let n = 0;
    const doc = mutableDoc([node(WRONG_WEEK_CHIP)]);
    const clock = virtualClock();
    const res = await runWait(doc, {
      clock,
      readWeek: async () => {
        n++;
        return n <= 4 ? DISPLAYED_HEADER : { text: '28 September – 4 October, 2026', startDate: '2026-09-28', endDate: '2026-10-04' };
      }
    });
    assert.equal(res.found, true, 'nach dem Wechsel ist der Zeitraum belegt');
    assert.equal(res.evidence, 'week-header');
    assert.ok(res.polls >= 5, `es wurde ueber die Polls hinweg weitergemacht (war ${res.polls})`);
    assert.ok(res.headerOutsideRangePolls >= 4, 'die Gegenbeleg-Polls wurden gezaehlt, nicht uebersprungen');
  });

  test('eventWaitObservationState: der neue Zustand ist eigenstaendig und disjunkt', () => {
    const q = (args) => eventWaitObservationState({ targetConfigured: true, targetStart: TARGET_WEEK_START, ...args });
    const st = q({ weekHeader: DISPLAYED_HEADER, eventElementsFound: 10, inRangeLabels: 0 });
    assert.equal(st.category, EVENT_WAIT_STATES.headerOutsideRange);
    assert.equal(st.dateQualified, false);
    assert.equal(st.loading, false);
    // Mit In-Range-Label gewinnt der Label-Beleg.
    assert.equal(q({ weekHeader: DISPLAYED_HEADER, inRangeLabels: 1 }).category, EVENT_WAIT_STATES.inRange);
    // Mit Spinner gewinnt "laedt noch".
    assert.equal(q({ weekHeader: DISPLAYED_HEADER, loading: true }).category, EVENT_WAIT_STATES.stillLoading);
    // MIT Zielwoche-Header: nie headerOutsideRange.
    const tgt = { text: 'x', startDate: '2026-09-28', endDate: '2026-10-04' };
    assert.notEqual(q({ weekHeader: tgt, eventElementsFound: 10 }).category, EVENT_WAIT_STATES.headerOutsideRange);
    // Und die uebrigen Zustaende bleiben unberuehrt.
    assert.equal(q({ eventElementsFound: 10, inRangeLabels: 0 }).category, EVENT_WAIT_STATES.outsideRange);
    assert.equal(q({ eventElementsFound: 0 }).category, EVENT_WAIT_STATES.emptyCandidate);
    // Kein datierter Header -> der Gegenbeleg existiert nicht.
    assert.equal(q({ weekHeader: { startDate: null, endDate: null }, eventElementsFound: 0 }).category, EVENT_WAIT_STATES.emptyCandidate);
    // Der Zustandswert ist eindeutig (kein Alias von eventsOutsideRange).
    assert.notEqual(EVENT_WAIT_STATES.headerOutsideRange, EVENT_WAIT_STATES.outsideRange);
  });

  test('reconcileEventWait behaelt den neuen Grund (kein Abbild auf Budget-Abbruch)', () => {
    const r = reconcileEventWait({
      wait: { polls: 40, outcome: 'budget-exhausted', abortReason: EVENT_WAIT_ABORT_REASONS.headerOutsideRange },
      eventElementsFound: 0, waitBudgetMs: 20000
    });
    assert.equal(r.abortReason, EVENT_WAIT_ABORT_REASONS.headerOutsideRange);
    assert.notEqual(r.abortReason, EVENT_WAIT_ABORT_REASONS.budgetExhausted);
    assert.notEqual(r.abortReason, EVENT_WAIT_ABORT_REASONS.renderedEmpty);
  });

  test('der Hinweistext sagt "Gegenbeleg" und nennt den Header', () => {
    const w = eventWaitWarning({
      outcome: 'budget-exhausted',
      abortReason: EVENT_WAIT_ABORT_REASONS.headerOutsideRange,
      waitedMs: 19800, budgetMs: 20000, renderedWeek: '2026-09-21–2026-09-27'
    });
    assert.ok(w && /Gegenbeleg/.test(w), w);
    assert.ok(/2026-09-21–2026-09-27/.test(w), 'der konkrete Header wird genannt');
    assert.ok(/Weder Treffer noch Leer-Beweis/.test(w));
    // Und er ist NICHT der Events-Ausserhalb-Text.
    const outside = eventWaitWarning({ outcome: 'budget-exhausted', abortReason: EVENT_WAIT_ABORT_REASONS.eventsOutsideRange });
    assert.notEqual(w, outside);
  });
});

// ===========================================================================
// FIX 4 (Live-Lauf 2): busySlotsReliable:true GLEICHZEITIG mit der eigenen
// Warnung "18 Event-Element(e) im DOM gefunden, aber 0 Busy-Slots extrahiert".
// Ein Ergebnis, das sich selbst widerspricht, ist nicht auswertbar. Die
// Invariante: sobald der Widerspruchstext entsteht, ist busySlotsReliable ein
// HARTES false.
// ===========================================================================
describe('FIX 2 — Verdrahtung: der neue Zustand ist im Ergebnis sichtbar', () => {
  const src = readFileSync(fileURLToPath(new URL('../src/owa.js', import.meta.url)), 'utf8');

  test('die Zustandsreihenfolge: Header-Gegenbeleg VOR der Event-Zahl', () => {
    const i = src.indexOf('export function eventWaitObservationState(');
    assert.ok(i > 0);
    const fn = src.slice(i, src.indexOf('export async function waitForCalendarEvents', i));
    const iContains = fn.indexOf('if (renderedWeekContains(weekHeader, targetStart))');
    const iInRange = fn.indexOf('if (inRange > 0)');
    const iLoading = fn.indexOf('if (loading)');
    const iHeader = fn.indexOf('if (weekHeader && weekHeader.startDate && weekHeader.endDate)');
    const iNodes = fn.indexOf('if (n > 0)');
    assert.ok(iContains > 0 && iInRange > iContains && iLoading > iInRange &&
      iHeader > iLoading && iNodes > iHeader,
    `Reihenfolge header-range@${iContains} < inRange@${iInRange} < loading@${iLoading} < header-ausserhalb@${iHeader} < events@${iNodes}`);
  });

  test('der Warte-Loop fuehrt den neuen Zustand (kein stilles Fallenlassen)', () => {
    const i = src.indexOf('export async function waitForCalendarEvents(');
    const fn = src.slice(i, src.indexOf('// Klassifiziert EINE Beobachtung', i + 100));
    assert.ok(/else if \(st\.category === EVENT_WAIT_STATES\.headerOutsideRange\)/.test(fn),
      'der Zustand braucht einen eigenen Zweig — sonst greift der Leer-Beweis');
    assert.ok(/headerOutsideRangePolls\+\+;/.test(fn), 'und einen eigenen Zaehler');
    assert.ok(/lastState === EVENT_WAIT_STATES\.headerOutsideRange\s*\n\s*\? EVENT_WAIT_ABORT_REASONS\.headerOutsideRange/.test(fn),
      'und einen eigenen Abbruchgrund');
  });

  test('die neue Ehrlichkeits-Kennzahl wird berichtet (sonst ist es in Runde 6 unsichtbar)', () => {
    assert.ok(src.includes('diag.eventWaitHeaderOutsideRangePolls'),
      'ohne dieses Feld laesst sich der Header-Gegenbeleg nicht von echter Leere unterscheiden');
    assert.ok(src.includes('diag.eventWaitLastState'));
  });

  test('reconcileEventWait bildet den neuen Grund NICHT auf Budget-Abbruch ab', () => {
    const i = src.indexOf('export function reconcileEventWait(');
    const fn = src.slice(i, src.indexOf('// EIN ehrlicher Abbruchgrund fuer die Diagnostik', i));
    assert.ok(/w\.abortReason === EVENT_WAIT_ABORT_REASONS\.headerOutsideRange/.test(fn),
      'der Grund muss durchgereicht werden');
  });
});

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

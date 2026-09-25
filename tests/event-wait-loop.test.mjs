import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import {
  waitForCalendarEvents,
  reconcileEventWait,
  effectiveAbortReason,
  makeEventObserver,
  isLoadIndicatorNode,
  loadIndicatorReason,
  LOAD_INDICATOR_PATTERNS,
  LOAD_INDICATOR_SELECTOR,
  FREEBUSY_EVENT_WAIT_SETTLE_MS,
  FREEBUSY_EVENT_WAIT_POLL_MS,
  EVENT_WAIT_ABORT_REASONS,
  EVENT_WAIT_OUTCOMES,
  eventWaitBudgetMs,
  collectCalendarEventLabels
} from '../src/owa.js';

// ===========================================================================
// Bug 3 — der Warte-Loop pollt nicht (Root Cause, Live-Befund 25.09.2026)
//
// Live-Kennzahlen des Smoke-Tests Runde 3 (echte OWA, KW 40):
//     waitedForEventsMs: 0   eventWaitPolls: 0   eventWaitBudgetMs: 20000
//     eventWaitOutcome: 'not-found'   Abbruch nach ~4s
// und census.structureNodes enthielt einen "ms-Spinner-circle" bei y≈487.
//
// Ursache (an der Quelle belegt): pollUntilEventGrid war die einzige
// Pollschleife und wurde per `if (hasEvents) { polls: 0, waitedMs: 0 }`
// uebersprungen. hasEvents kam aus der schnellen Vorab-Probe und war STALE —
// die Vorabmeldung entschied, OB ueberhaupt gepollt wird, und verhinderte
// damit jede Widerlegung. Zusaetzlich hat resolveEventWaitOutcome polls und
// waitBudgetMs destrukturiert und nie benutzt: "Budget erschoepft" war von
// "0 Polls bei vorhandenem Budget" nicht unterscheidbar.
//
// Die Tests hier fahren den ECHTEN Draht: makeEventObserver (Produktion) ->
// collectCalendarEventLabels im 'fast'-Modus (Produktion, echter evaluate-
// Callback gegen ein Fake-DOM) -> waitForCalendarEvents (Produktion).
// Es gibt hier KEINE nachgebaute Beobachtung — genau das war bei Bug 2 der
// Grund, warum der Defekt ungetestet durchging.
// ===========================================================================

// Virtuelle Uhr: die Wartephase darf im Test nicht echt 20s dauern.
const virtualClock = (start = 1_000_000) => {
  let t = start;
  const sleeps = [];
  return {
    now: () => t,
    sleep: async (ms) => { sleeps.push(ms); t += ms; },
    advance: (ms) => { t += ms; },
    get slept() { return sleeps; }
  };
};

const withFakeDocument = async (doc, fn) => {
  const prev = globalThis.document;
  globalThis.document = doc;
  try { return await fn(); } finally { globalThis.document = prev; }
};

// Minimaler DOM-Knoten — der Kollektor darf nichts voraussetzen, was ein
// unit-test-Fake nicht liefert (tagName/className/innerText/attributes...).
const node = (over = {}) => ({
  tagName: over.tag || 'DIV',
  getAttribute: (n) => (over.attrs && over.attrs[n] !== undefined ? over.attrs[n] : null),
  getBoundingClientRect: () => over.rect || { left: 0, top: 0, width: 120, height: 24, x: 0, y: 0 },
  offsetParent: null,
  innerText: over.text || '',
  className: over.className || '',
  attributes: Object.entries(over.attrs || {}).map(([name, value]) => ({ name, value })),
  childElementCount: 0,
  children: [],
  ...over.extra
});

const EVENT_CHIP = { attrs: { 'aria-label': 'TeamSync, 10:00 to 11:00, Monday, September 28, 2026, Busy' } };
// Echter OWA-Spinner: <ms-Spinner class="ms-Spinner ms-Spinner-circle">. Die
// Klasse steht als echtes Attribut im DOM — der Selektor der Produktion
// ([class*="ms-Spinner" i]) greift genau darueber.
const SPINNER = { tag: 'MS-SPINNER', className: 'ms-Spinner ms-Spinner-circle', attrs: { class: 'ms-Spinner ms-Spinner-circle' } };
const DAY_HEADER = (label, left) => node({
  tag: 'TH', attrs: { role: 'columnheader', 'aria-label': label },
  rect: { left, top: 100, width: 100, height: 20, x: left, y: 100 }
});

// Mini-CSS-Matcher fuer genau die Selektorformen, die der Kollektor benutzt:
// '*', Tag, [attr], [attr="wert"], [attr*="teil"] (jeweils optional case-insensitive).
const matches = (el, sel) => {
  const s = String(sel).trim();
  if (s === '*') return true;
  const tag = /^([a-zA-Z][\w-]*)$/.exec(s);
  if (tag) return String(el.tagName || '').toLowerCase() === tag[1].toLowerCase();
  // [attr*="teil"] zuerst pruefen — sonst wuerde es als [attr] fehlgreifen.
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
  let anyHas = false;
  while ((m = attrHas.exec(s)) !== null) {
    anyHas = true;
    if (el.getAttribute && el.getAttribute(m[1]) != null) return true;
  }
  return false;
};

// Ein Dokument, dessen Zustand sich pro Zustandswechsel austauschen laesst.
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

// Der Produktions-Beobachter gegen ein Fake-DOM — genau die Komposition, die
// getFreeBusy benutzt (makeEventObserver -> collectCalendarEventLabels fast).
const productionObserver = () => makeEventObserver(() =>
  collectCalendarEventLabels(fakePage, { weekStartIso: '2026-09-28', fast: true }));

// Warteschleife mit installiertem Fake-DOM fahren. Ohne withFakeDocument
// laeuft collectCalendarEventLabels gegen ein NICHTS vorhandenes document —
// das waere eine Luege, kein Test.
const runWait = (doc, opts = {}) => withFakeDocument(doc, () => waitForCalendarEvents({
  observe: makeEventObserver(opts.collect || (() =>
    collectCalendarEventLabels(fakePage, { weekStartIso: '2026-09-28', fast: true }))),
  budgetMs: 20000,
  pollMs: FREEBUSY_EVENT_WAIT_POLL_MS,
  settleMs: FREEBUSY_EVENT_WAIT_SETTLE_MS,
  now: opts.clock ? opts.clock.now : virtualClock().now,
  sleep: opts.clock ? opts.clock.sleep : virtualClock().sleep
}));

// ===========================================================================
describe('Bug 3 — Lade-Indikator ist eine eigene Kategorie, kein Event', () => {

  test('isLoadIndicatorNode erkennt den echten OWA-Spinner am ms-Spinner-circle', () => {
    assert.equal(isLoadIndicatorNode({ tag: 'MS-SPINNER', cls: 'ms-Spinner ms-Spinner-circle' }), true);
    assert.equal(isLoadIndicatorNode({ tag: 'DIV', cls: 'ms-Spinner-circle' }), true);
    assert.equal(isLoadIndicatorNode({ tag: 'DIV', role: 'progressbar' }), true);
    assert.equal(isLoadIndicatorNode({ tag: 'DIV', attrs: undefined, data: { 'data-loading': 'true' } }), true);
    assert.equal(isLoadIndicatorNode({ tag: 'DIV', ariaLabel: 'Wird geladen…' }), true);
  });

  test('isLoadIndicatorNode lässt echte Events und Kalender-Nodes durch', () => {
    assert.equal(isLoadIndicatorNode(null), false);
    assert.equal(isLoadIndicatorNode(undefined), false);
    assert.equal(isLoadIndicatorNode({ tag: 'DIV', cls: 'ms-Calendar-week', text: 'Work week' }), false);
    assert.equal(isLoadIndicatorNode({ tag: 'DIV', ariaLabel: 'TeamSync, 10:00 to 11:00, Monday, September 28, 2026' }), false);
    assert.equal(isLoadIndicatorNode({ tag: 'BUTTON', cls: 'ms-Button', title: 'Nächste Woche' }), false);
  });

  test('loadIndicatorReason nennt die Kategorie und NICHT "leeres Grid"', () => {
    const r = loadIndicatorReason({ eventElementsFound: 0, loading: true });
    assert.equal(r.loading, true);
    assert.equal(r.eventElementsFound, 0);
    assert.notEqual(r.category, 'empty');
  });

  test('der Selektor ist dokumentweit und kostenarm (kein Universal-Durchlauf)', () => {
    assert.ok(LOAD_INDICATOR_SELECTOR.includes('[role="progressbar"]'));
    assert.ok(/ms-Spinner/.test(LOAD_INDICATOR_SELECTOR));
    const parts = LOAD_INDICATOR_SELECTOR.split(',').map((s) => s.trim());
    assert.equal(parts.includes('*'), false, 'kein nackter Universalselektor im Poll-Pfad');
    assert.ok(parts.every((p) => p.startsWith('[') || p.startsWith('ms-')),
      `nur Attribut-/Tag-Selektoren: ${JSON.stringify(parts)}`);
    assert.ok(LOAD_INDICATOR_PATTERNS.cls && LOAD_INDICATOR_PATTERNS.tag && LOAD_INDICATOR_PATTERNS.role);
  });

  test('Spinner im Grid-Body zaehlt als loading=true und NICHT als Event', async () => {
    const doc = mutableDoc([
      // Zwei Tageskoepfe, damit region entsteht und der Struktur-Scan laeuft.
      DAY_HEADER('Monday', 100),
      DAY_HEADER('Tuesday', 200),
      node({ ...SPINNER, rect: { left: 150, top: 487, width: 16, height: 16, x: 150, y: 487 } })
    ]);
    await withFakeDocument(doc, async () => {
      const res = await collectCalendarEventLabels(fakePage, { weekStartIso: '2026-09-28' });
      assert.equal(res.loading, true, 'Spinner muss als Lade-Indikator erkannt werden');
      assert.equal(res.eventElementsFound, 0, 'Spinner ist kein Event');
      assert.equal(res.labels.length, 0);
      assert.equal(res.census.loadingNodes, 1, 'Spinner landet in der eigenen Kategorie');
      assert.equal(res.census.structureNodes, 0, 'Spinner darf NICHT als Struktur-/Geometrie-Kandidat durchrutschen');
      assert.ok(res.loadingIndicators.length >= 1, 'Spinner wird benannt');
      assert.ok(res.loadingIndicators.every((x) => x.reason === 'load-indicator'));
    });
  });

  test('Spinner ueber die Attribut-Leiter wird ebenfalls nicht als Event gezaehlt', async () => {
    // Der Spinner haengt in der Praxis an aria-label, nicht nur an der Klasse.
    const spinnerViaAttr = node({ tag: 'DIV', className: 'ms-ProgressIndicator', attrs: { 'aria-label': 'Wird geladen' }, rect: { left: 10, top: 10, width: 30, height: 30, x: 10, y: 10 } });
    const doc = mutableDoc([spinnerViaAttr]);
    await withFakeDocument(doc, async () => {
      const res = await collectCalendarEventLabels(fakePage, { weekStartIso: '2026-09-28' });
      assert.equal(res.eventElementsFound, 0);
      assert.equal(res.loading, true);
    });
  });
});

// ===========================================================================
describe('Bug 3 — der Warte-Loop pollt (Live: polls 0 bei Budget 20000)', () => {

  test('REGRESSION: Spinner fuer N Polls, dann Events -> polls>0, waited>0, found', async () => {
    // Der Kern-Fall aus der Anforderung: ein Fake, der N Polls lang einen
    // Spinner zeigt und ERST DANACH Events liefert. Ein Fake, der sofort
    // Events liefert, faengt diesen Bug nicht.
    const SPINNER_POLLS = 6;
    let polls = 0;
    const doc = mutableDoc([node(SPINNER)]);
    const clock = virtualClock();

    const res = await runWait(doc, {
      clock,
      collect: async () => {
        polls++;
        // Waehrend der Ladephase nur der Spinner im Dokument, danach das Event.
        doc.set(polls <= SPINNER_POLLS ? [node(SPINNER)] : [node(EVENT_CHIP)]);
        // Echter Kollektor, 'fast'-Modus = Produktionspfad der Warte-Probe.
        return collectCalendarEventLabels(fakePage, { weekStartIso: '2026-09-28', fast: true });
      }
    });

    assert.ok(res.polls > 0, `polls muss > 0 sein (war ${res.polls})`);
    assert.ok(res.waitedMs > 0, `waitedMs muss > 0 sein (war ${res.waitedMs})`);
    assert.equal(res.found, true);
    assert.equal(res.outcome, 'found');
    assert.equal(res.loadingPolls, SPINNER_POLLS, 'jeder Spinner-Poll wird als "noch nicht fertig" gezaehlt');
    assert.equal(res.abortReason, null);
    assert.ok(res.waitedMs >= SPINNER_POLLS * FREEBUSY_EVENT_WAIT_POLL_MS,
      `hat den Spinner wirklich abgewartet: ${res.waitedMs}ms`);
  });

  test('REGRESSION: die stale Vorab-Meldung darf den Poll NICHT mehr ueberspringen', async () => {
    // Live-Situation: die schnelle Vorab-Probe meldete "Events da" (stale),
    // die abschliessende Sammlung fand 0. Genau diese stale Meldung hat den
    // Poll uebersprungen. Die Wartephase darf sich auf KEINE Vorab-Meldung
    // verlassen — sie beobachtet selbst.
    let polls = 0;
    const doc = mutableDoc([node(EVENT_CHIP)]);
    const clock = virtualClock();
    const res = await runWait(doc, {
      clock,
      collect: async () => {
        polls++;
        doc.set(polls <= 4 ? [node(SPINNER)] : [node(EVENT_CHIP)]);
        return collectCalendarEventLabels(fakePage, { weekStartIso: '2026-09-28', fast: true });
      }
    });
    assert.ok(polls >= 5, `es wurde weiter gepollt trotz "Vorab-Events" (polls=${polls})`);
    assert.equal(res.found, true);
    assert.equal(res.outcome, 'found');
  });

  test('Budget wird tatsaechlich ausge schoepft, wenn dauerhaft nur der Spinner da ist', async () => {
    const doc = mutableDoc([node(SPINNER)]);
    const clock = virtualClock();
    const res = await runWait(doc, { clock });
    assert.equal(res.found, false);
    assert.equal(res.outcome, 'budget-exhausted');
    assert.equal(res.abortReason, EVENT_WAIT_ABORT_REASONS.budgetExhausted);
    assert.equal(res.abortReason, 'wait-budget-exhausted');
    assert.ok(res.polls > 30, `bei 500ms Poll ueber 20s sind >30 Polls faellig (war ${res.polls})`);
    assert.ok(res.waitedMs >= 20000 - FREEBUSY_EVENT_WAIT_POLL_MS,
      `das Budget wurde ausge schoepft (war ${res.waitedMs}ms)`);
    assert.equal(res.loadingPolls, res.polls, 'jeder Poll sah den Spinner');
    assert.ok(clock.slept.length > 30, 'es wurde auch wirklich geschlafen');
  });

  test('frueher Abbruch NUR bei belegter Leerheit: Ruhe-Fenster muss vollstaendig sein', async () => {
    // Ohne Events und ohne Spinner ist die Ansicht gerendert — aber EIN
    // spinner-freier Blick ist kein Leer-Beweis (direkt nach Klick ist das
    // Grid kurz leer, bevor es befuellt wird). Erst nach settleMs wird
    // 'rendered-empty-proven' gemeldet.
    const doc = mutableDoc([]);
    const clock = virtualClock();
    const res = await runWait(doc, { clock });
    assert.equal(res.found, false);
    assert.equal(res.outcome, 'not-found');
    assert.equal(res.abortReason, 'rendered-empty-proven');
    assert.ok(res.polls >= 2, `nicht nach einem einzigen Poll beenden (polls=${res.polls})`);
    assert.ok(res.waitedMs >= FREEBUSY_EVENT_WAIT_SETTLE_MS,
      `Ruhe-Fenster wurde abgewartet (war ${res.waitedMs}ms)`);
    assert.ok(res.waitedMs < 20000, 'und trotzdem deutlich vor dem Budgetende');
    assert.equal(res.loadingPolls, 0);
  });

  test('Spinner bricht das Ruhe-Fenster ab: kein vorzeitiger Leer-Beweis', async () => {
    // Phase 1: 2 Polls ruhig (Grid noch nicht befuellt) -> Phase 2: Spinner.
    // Ohne Zuruecksetzen des Ruhe-Fensters wuerde hier faelschlich "leer"
    // gemeldet, obwohl die Ansicht noch laedt.
    const doc = mutableDoc([]);
    let n = 0;
    const clock = virtualClock();
    const res = await runWait(doc, {
      clock,
      collect: async () => {
        n++;
        doc.set(n <= 2 ? [] : [node(SPINNER)]);
        return collectCalendarEventLabels(fakePage, { weekStartIso: '2026-09-28', fast: true });
      }
    });
    assert.notEqual(res.abortReason, 'rendered-empty-proven',
      'ein zwischenzeitlicher Spinner macht das Ruhe-Fenster ungueltig');
    assert.equal(res.abortReason, 'wait-budget-exhausted');
    assert.ok(res.loadingPolls > 0, 'Spinner wurden gezaehlt');
  });

  test('Events sofort da: 1 Poll, 0ms — der Poll darf nicht kuenstlich warten', async () => {
    const doc = mutableDoc([node(EVENT_CHIP)]);
    const clock = virtualClock();
    const res = await runWait(doc, { clock });
    assert.equal(res.polls, 1);
    assert.equal(res.found, true);
    assert.equal(res.outcome, 'found');
    assert.equal(res.waitedMs, 0);
  });

  test('kein Budget -> ehrlich benennen, NICHT "0 Polls bei 20s"', async () => {
    const clock = virtualClock();
    const res = await waitForCalendarEvents({
      observe: async () => ({ eventElementsFound: 0, loading: false }),
      budgetMs: eventWaitBudgetMs(3000), // -> 0
      now: clock.now, sleep: clock.sleep
    });
    assert.equal(res.polls, 0);
    assert.equal(res.outcome, 'budget-exhausted');
    assert.equal(res.abortReason, EVENT_WAIT_ABORT_REASONS.budgetUnavailable);
  });

  test('der Beobachter meldet loading aus dem echten Kollektor', async () => {
    const doc = mutableDoc([node(SPINNER)]);
    const observe = productionObserver();
    const a = await withFakeDocument(doc, () => observe());
    assert.deepEqual(a, { eventElementsFound: 0, loading: true },
      'Spinner -> noch nicht fertig, nicht "leer"');
    doc.set([node(EVENT_CHIP)]);
    const b = await withFakeDocument(doc, () => observe());
    assert.equal(b.eventElementsFound, 1);
    assert.equal(b.loading, false);
    assert.equal(typeof observe.lastSelector, 'function', 'der zuletzt erfolgreiche Selektor ist abrufbar');
    assert.ok(observe.lastSelector(), 'er wurde aus der Leiter uebernommen');
  });

  test('loadIndicatorReason trennt die vier Zustaende der Wartephase', () => {
    assert.equal(loadIndicatorReason({ eventElementsFound: 6, loading: false }).category, 'events');
    assert.equal(loadIndicatorReason({ eventElementsFound: 0, loading: true }).category, 'still-loading');
    assert.equal(loadIndicatorReason({ eventElementsFound: 0, loading: false }).category, 'empty-candidate');
    // Der Live-Zustand: Spinner + 0 Events ist "still-loading", nicht leer.
    assert.notEqual(loadIndicatorReason({ eventElementsFound: 0, loading: true }).category, 'empty-candidate');
  });
});

// ===========================================================================
describe('Bug 3 — ehrlicher Abbruchgrund statt stillem Ende', () => {

  test('reconcileEventWait: 0 Events + Budget + 0 Polls -> not-polled (Logikfehler)', () => {
    const r = reconcileEventWait({ wait: { polls: 0, outcome: 'not-found' }, eventElementsFound: 0, waitBudgetMs: 20000 });
    assert.equal(r.outcome, 'not-polled');
    assert.equal(r.abortReason, 'wait-not-polled');
    assert.ok(EVENT_WAIT_OUTCOMES.includes('not-polled'));
  });

  test('reconcileEventWait: das ist genau der Live-Zustand polls=0/budget=20000', () => {
    // Live: waitedForEventsMs 0, eventWaitPolls 0, eventWaitBudgetMs 20000.
    const r = reconcileEventWait({ wait: { polls: 0, waitedMs: 0 }, eventElementsFound: 0, waitBudgetMs: 20000 });
    assert.equal(r.outcome, 'not-polled', 'haette vorher als not-found durchrutschen muessen');
  });

  test('reconcileEventWait: Events in der Abschluss-Sammlung -> found', () => {
    const r = reconcileEventWait({ wait: { polls: 3, outcome: 'found' }, eventElementsFound: 6, waitBudgetMs: 20000 });
    assert.equal(r.outcome, 'found');
    assert.equal(r.abortReason, null);
  });

  test('reconcileEventWait: gepollt, Budget erschoepft -> budget-exhausted', () => {
    const r = reconcileEventWait({
      wait: { polls: 40, outcome: 'budget-exhausted', abortReason: 'wait-budget-exhausted' },
      eventElementsFound: 0, waitBudgetMs: 20000
    });
    assert.equal(r.outcome, 'budget-exhausted');
    assert.equal(r.abortReason, 'wait-budget-exhausted');
  });

  test('reconcileEventWait: belegt leer -> not-found mit Beleg-Grund', () => {
    const r = reconcileEventWait({
      wait: { polls: 3, outcome: 'not-found', abortReason: 'rendered-empty-proven' },
      eventElementsFound: 0, waitBudgetMs: 20000
    });
    assert.equal(r.outcome, 'not-found');
    assert.equal(r.abortReason, 'rendered-empty-proven');
  });

  test('reconcileEventWait: kein Budget + keine Polls ist kein "not-polled"-Fehler', () => {
    const r = reconcileEventWait({ wait: { polls: 0, abortReason: 'wait-budget-unavailable' }, eventElementsFound: 0, waitBudgetMs: 0 });
    assert.equal(r.outcome, 'budget-exhausted');
    assert.equal(r.abortReason, 'wait-budget-unavailable');
  });

  test('effectiveAbortReason: Navigation "none" + Wartephase gescheitert -> warte-Grund', () => {
    assert.equal(effectiveAbortReason({ navReason: 'none', eventWaitReason: 'wait-budget-exhausted', eventElementsFound: 0 }),
      'wait-budget-exhausted');
    assert.equal(effectiveAbortReason({ navReason: 'none', eventWaitReason: 'wait-not-polled', eventElementsFound: 0 }),
      'wait-not-polled');
  });

  test('effectiveAbortReason: beide Gruende sichtbar; Events gefunden -> kein Warte-Grund', () => {
    assert.equal(effectiveAbortReason({ navReason: 'max-clicks', eventWaitReason: 'wait-budget-exhausted', eventElementsFound: 0 }),
      'max-clicks+wait-budget-exhausted');
    assert.equal(effectiveAbortReason({ navReason: 'none', eventWaitReason: 'wait-budget-exhausted', eventElementsFound: 4 }), 'none');
    assert.equal(effectiveAbortReason({ navReason: 'none', eventWaitReason: null, eventElementsFound: 0 }), 'none');
  });
});

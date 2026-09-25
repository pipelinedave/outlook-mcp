import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
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
  EVENT_WAIT_STATES,
  eventWaitBudgetMs,
  eventWaitWarning,
  countLabelsInRange,
  eventWaitObservationState,
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

// BUG 4: Der angefragte Zeitraum ist fuer die Wartephase Pflicht — ohne ihn kann
// sie keinen datumsqualifizierten Beleg erbringen und meldet 'wait-target-unknown'
// statt still 'found'. Die Beobachter im Test werden deshalb mit DIESEM Range
// gebaut, genau wie getFreeBusy es tut.
const TARGET_WEEK_START = '2026-09-28';
const TEST_RANGE = {
  rangeStartMs: new Date('2026-09-28T00:00:00').getTime(),
  rangeEndMs: new Date('2026-10-05T00:00:00').getTime()
};

// Der Produktions-Beobachter gegen ein Fake-DOM — genau die Komposition, die
// getFreeBusy benutzt (makeEventObserver -> collectCalendarEventLabels fast).
const productionObserver = (opts = {}) => makeEventObserver(() =>
  collectCalendarEventLabels(fakePage, { weekStartIso: TARGET_WEEK_START, fast: true }), {
    readWeek: opts.readWeek || null,
    range: opts.range === null ? null : (opts.range || TEST_RANGE),
    defaultYear: 2026,
    now: opts.now
  });

// Warteschleife mit installiertem Fake-DOM fahren. Ohne withFakeDocument
// laeuft collectCalendarEventLabels gegen ein NICHTS vorhandenes document —
// das waere eine Luege, kein Test.
const runWait = (doc, opts = {}) => withFakeDocument(doc, () => waitForCalendarEvents({
  observe: makeEventObserver(opts.collect || (() =>
    collectCalendarEventLabels(fakePage, { weekStartIso: TARGET_WEEK_START, fast: true })), {
      readWeek: opts.readWeek || null,
      range: opts.range === null ? null : (opts.range || TEST_RANGE),
      defaultYear: 2026
    }),
  budgetMs: opts.budgetMs === undefined ? 20000 : opts.budgetMs,
  pollMs: FREEBUSY_EVENT_WAIT_POLL_MS,
  settleMs: FREEBUSY_EVENT_WAIT_SETTLE_MS,
  targetStart: opts.targetStart === undefined ? TARGET_WEEK_START : opts.targetStart,
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
    // BUG 4: Die Beobachtung traegt jetzt mehr als die Knotenzahl — genau das
    // war der Defekt. Der deepEqual bleibt, aber auf der neuen Form.
    assert.deepEqual(a, {
      eventElementsFound: 0,
      loading: true,
      weekHeader: null,
      inRangeLabels: 0,
      labelsOutOfRange: 0,
      labelsUndated: 0
    }, 'Spinner -> noch nicht fertig, nicht "leer"');
    doc.set([node(EVENT_CHIP)]);
    const b = await withFakeDocument(doc, () => observe());
    assert.equal(b.eventElementsFound, 1);
    assert.equal(b.loading, false);
    // EVENT_CHIP traegt "Monday, September 28, 2026" -> liegt im Test-Range.
    assert.equal(b.inRangeLabels, 1, 'BUG 4: das Label AUS dem angefragten Zeitraum wird gezaehlt');
    assert.equal(typeof observe.lastSelector, 'function', 'der zuletzt erfolgreiche Selektor ist abrufbar');
    assert.ok(observe.lastSelector(), 'er wurde aus der Leiter uebernommen');
  });

  test('BUG 4: ohne konfigurierten Range bleibt inRangeLabels null (unmeasured)', async () => {
    // Der Beobachter MUSS die Unmessbarkeit benennen, statt 0 zu liefern:
    // eine gemeldete 0 waere von einem echten "0 Labels im Zeitraum" nicht
    // unterscheidbar — und genau diese Verwechslung war der Live-Bug.
    const doc = mutableDoc([node(EVENT_CHIP)]);
    const observe = productionObserver({ range: null });
    const o = await withFakeDocument(doc, () => observe());
    assert.equal(o.eventElementsFound, 1, 'der Knoten ist da …');
    assert.equal(o.inRangeLabels, null, '… aber es ist nicht messbar, ob er im Zeitraum liegt');
    assert.equal(o.labelsOutOfRange, null);
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
// Verdrahtungs-Waechter
// ---------------------------------------------------------------------------
// Eigene Klasse, mit einem Mutationsnachweis entwickelt: eine erste Fassung
// dieser Tests bestand NUR aus waitForCalendarEvents + Kollektor und wurde
// GRUEN, obwohl in getFreeBusy der stale-Boolean-Kurzschluss wiederhergestellt
// war — also genau der Live-Bug. Die Funktionen koennen also voellig korrekt
// sein und der Draht trotzdem falsch. Deshalb wird die Verdrahtung in
// getFreeBusy hier mitgeprueft.
// ===========================================================================
describe('Bug 3 — Verdrahtung: getFreeBusy pollt bedingungslos', () => {
  const src = readFileSync(fileURLToPath(new URL('../src/owa.js', import.meta.url)), 'utf8');
  // Der Block von der Budget-Berechnung bis zur Abschluss-Sammlung.
  const i = src.indexOf('const eventWaitBudget = eventWaitBudgetMs(remaining());');
  assert.ok(i > 0, 'eventWaitBudget nicht gefunden');
  const block = src.slice(i, src.indexOf('const rawLabels = eventCollection.labels;', i));

  test('die Wartephase wird BEDINGUNGSLOS gestartet, nicht hinter hasEvents', () => {
    assert.ok(/const eventWait = await waitForCalendarEvents\(\{/.test(block),
      'der Aufruf darf nicht in einem if(haEvents) versteckt sein');
    assert.ok(!/if \(hasEvents\)\s*\{[\s\S]{0,200}waitForCalendarEvents/.test(block),
      'KEIN if(hasEvents)-Kurzschluss um die Wartephase — das war die Root Cause');
    assert.ok(!/outcome: 'pending'/.test(block),
      "das 'pending'-Shortcut-Objekt (polls:0, waitedMs:0) darf nicht zurueckkehren");
  });

  test('der Beobachter kommt aus makeEventObserver, nicht aus der boolean-Vorabprobe', () => {
    assert.ok(/makeEventObserver\(\(preferSelector\) => collectCalendarEventLabels\(page/.test(block),
      'der Beobachter muss der exportierte Produktions-Beobachter sein');
    assert.ok(/preferSelector: observeCalendarEvents\.lastSelector\(\)/.test(src),
      'die Abschluss-Sammlung nutzt den Selector-Memory des Beobachters');
  });

  test('die Vorab-Probe hatEventElements steuert NUR noch GUARD 2', () => {
    // Die Vorab-Probe darf existieren (GUARD 2 braucht sie), aber sie darf die
    // Wartephase nicht mehr steuern.
    assert.ok(/hasEvents = hasEvents \|\| \(await hasEventElements\(\)\)/.test(src),
      'GUARD-2-Aktualisierung bleibt erhalten');
    const guard = src.slice(src.indexOf('if (!rendered.startDate && !hasEvents) {'), src.indexOf('const eventWaitBudget'));
    assert.ok(!/waitForCalendarEvents/.test(guard), 'die Vorab-Probe fuehrt nicht zur Wartephase');
  });

  // ---------------------------------------------------------------------
  // BUG 4/5 — Verdrahtung. Die Funktionen koennen voellig korrekt sein und
  // der Draht trotzdem wochenblind: dann baute getFreeBusy weiterhin eine
  // Beobachtung OHNE Zielzeitraum und OHNE Wochen-Header — also genau die
  // Konfiguration, die 'wait-target-unknown' ausloest. Das ist der Draht,
  // den der Kontrollfall in Runde 5 braucht.
  // ---------------------------------------------------------------------
  test('BUG 4: der Beobachter bekommt den Range UND den Wochen-Header', () => {
    const i = src.indexOf('const observeCalendarEvents = makeEventObserver(');
    assert.ok(i > 0, 'makeEventObserver wird in getFreeBusy gerufen');
    const j = src.indexOf('const eventWait = await waitForCalendarEvents({', i);
    assert.ok(j > i, 'die Wartephase folgt auf den Beobachter');
    const wiring = src.slice(i, j);
    assert.ok(/readWeek: \(\) => readRenderedCalendarWeek\(page\)/.test(wiring),
      'der Beobachter MUSS den Wochen-Header lesen — ohne ihn gibt es Belegkriterium (1) nicht');
    assert.ok(/range: \{ rangeStartMs, rangeEndMs \}/.test(wiring),
      'der Beobachter MUSS den angefragten Zeitraum kennen — ohne ihn gibt es Belegkriterium (2) nicht');
    assert.ok(/defaultYear/.test(wiring), 'Labels ohne Jahr brauchen das Fallback-Jahr');
  });

  // MUTATIONSNACHWEIS fuer diesen Draht: mit targetStart: null bleibt die
  // Testsuite GRUEN (die Funktionen sind dann korrekt, aber der Aufrufer
  // weekenblind — genau die Konfiguration, die 'wait-target-unknown'
  // ausloest). Dieser Test schliesst diese Luecke.
  test('BUG 4: die Wartephase bekommt die Zielwoche durchgereicht', () => {
    const i = src.indexOf('const eventWait = await waitForCalendarEvents({');
    assert.ok(i > 0, 'die Wartephase wird gestartet');
    const call = src.slice(i, src.indexOf('});', i) + 3);
    assert.ok(/targetStart: rangeStartDate/.test(call),
      'ohne targetStart meldet die Wartephase wait-target-unknown statt found');
    // Auch indirekt: ein umbenanntes/verschobenes rangeStartDate darf nicht
    // unbemerkt durchrutschen.
    assert.ok(call.includes('targetStart:'), 'der Schluessel selbst muss vorhanden sein');
    assert.ok(!/targetStart:\s*(null|undefined|0|''|"")/.test(call),
      'und er darf nicht leer sein — das waere der weekenblinde Aufruf');
  });

  test('BUG 4: die Wartephase wird mit konfiguriertem Ziel auch wirklich genutzt', async () => {
    // Fuehrt den ECHTEN Aufrufweg mit dem Ziel nach — der Kern des
    // Kontrolllaufs. Ohne Ziel: 'wait-target-unknown'. Mit Ziel: datums-
    // qualifiziertes found.
    // EVENT_CHIP traegt "Monday, September 28, 2026" -> liegt im Test-Range.
    const doc = mutableDoc([node(EVENT_CHIP)]);
    const withTarget = await runWait(doc, { clock: virtualClock() });
    assert.equal(withTarget.abortReason, null, 'mit Ziel wird belegt');
    assert.equal(withTarget.dateQualified, true);

    const withoutTarget = await runWait(doc, { clock: virtualClock(), targetStart: null, range: null });
    assert.equal(withoutTarget.abortReason, EVENT_WAIT_ABORT_REASONS.targetUnknown,
      'ohne Ziel wird der Konfigurationsfehler benannt, nicht still gepollt');
    assert.equal(withoutTarget.dateQualified, false);
  });

  test('BUG 4: die Range-Werte werden VOR der Wartephase berechnet (keine TDZ)', () => {
    // Zwei der drei Live-Defekte waren TDZ-Fehler. Hier waere es derselbe:
    // rangeStartMs wird im Wette-Block benutzt, aber erst weiter unten gebildet.
    const iWait = src.indexOf('const observeCalendarEvents = makeEventObserver(');
    const iRange = src.indexOf('const rangeStartMs = new Date(rangeStartIso).getTime();');
    assert.ok(iRange > 0, 'rangeStartMs wird gebildet');
    assert.ok(iRange < iWait,
      `der Range muss VOR dem Beobachter gebildet sein (range@${iRange}, warte@${iWait})`);
    // Und genau einmal — eine spaetere Zweitbildung waere driftanfaellig.
    assert.equal((src.match(/const rangeStartMs = new Date\(rangeStartIso\)\.getTime\(\);/g) || []).length, 1,
      'genau eine Bildung (die der Abschluss-Extraktion dient)');
  });

  test('BUG 4: die Abschluss-Extraktion benutzt dieselben Range-Werte', () => {
    assert.ok(/partitionSlotsByRange\(dedupedSlots, rangeStartMs, rangeEndMs\)/.test(src),
      'sonst wuerde die Wartephase nach einer anderen Definition von "im Zeitraum" arbeiten');
  });

  test('BUG 4/5: die neuen Ehrlichkeits-Kennzahlen werden berichtet', () => {
    // Ohne diese Felder laesst sich im naechsten Live-Lauf NICHT entscheiden,
    // ob der Wartepfad gearbeitet hat.
    for (const f of ['diag.eventWaitDateQualified', 'diag.eventWaitInRangeLabels',
      'diag.eventWaitOutsideRangePolls', 'diag.viewWaitMsSpent']) {
      assert.ok(src.includes(f), `${f} fehlt in der Diagnostik`);
    }
    assert.ok(/eventWaitWarning\(\{[^}]*abortReason: eventWait\.abortReason/.test(src),
      'der Warte-Abbruchgrund muss auch in den Hinweistext');
  });

  test('BUG 5: der Render-Wait verlaesst "Body hat Events" nicht mehr sofort', () => {
    // Live-Befund Runde 4: `else if (events) break;` verliess die Ansicht nach
    // dem ersten Poll, obwohl der Header fehlte. Genau daher kamen die 13ms.
    const i = src.indexOf('const waitForCalendarRendered = async (budgetMs) => {');
    assert.ok(i > 0, 'waitForCalendarRendered existiert');
    const fn = src.slice(i, src.indexOf('\n    };', i));
    assert.ok(!/else if \(events\) \{\s*\n\s*break;/.test(fn),
      'der Premature-Exit "kein Header + Events -> fertig" IST der Live-Bug');
    assert.ok(/eventsSeenAt/.test(fn), 'stattdessen bekommt dieser Fall eine Gnadenfrist');
    assert.ok(/if \(!eventsSeenAt \|\| loading\) eventsSeenAt = Date\.now\(\);/.test(fn),
      'ein Spinner verlaengert die Gnadenfrist — "noch nicht fertig" ist ein eigener Zustand');
    assert.ok(/if \(Date\.now\(\) - eventsSeenAt >= EVENT_GRACE_MS\) break;/.test(fn),
      'erst nach der Gnadenfrist wird Zustand 1 gemeldet');
  });

  test('BUG 5: der Aufrufer kann Zustand 1 nicht mehr durchreichen', () => {
    // Im Delta-Pfad (nach einem Klick) war es ein break mit Abbruchgrund. Der
    // Guard-Pfad gibt jetzt continue zurueck — die Schleife bewertet neu.
    const delta = src.slice(src.indexOf('const r = await clickWeekButton(direction, rendered.startDate);'),
      src.indexOf('if (r.targetReached) {', src.indexOf('const r = await clickWeekButton(direction, rendered.startDate);')));
    assert.ok(!/abortReason = 'view-not-ready'/.test(delta),
      'der Klickpfad darf den Zustand 1 nicht als Abbruch durchreichen');
    assert.ok(/continue;/.test(delta), 'sondern zurueck in die Zustandsbewertung');
  });

  test('BUG 5: der Zustand-1-Zweig steht VOR der Delta-Berechnung', () => {
    // Reihenfolge ist die Korrektheit: Zustand 1 muss VOR jeder Delta-
    // Berechnung liegen, sonst rechnet der Pfad mit einem startDate, das
    // nicht existiert.
    const loopStart = src.indexOf('for (let step = 0; step < 30; step++) {');
    assert.ok(loopStart > 0, 'die Schritt-Schleife existiert');
    const iState = src.indexOf('const view = weekViewState({', loopStart);
    const iTarget = src.indexOf('WEEK_VIEW_STATES.targetWeek) {', iState);
    const iNotReady = src.indexOf('if (view.state === WEEK_VIEW_STATES.notReady) {', iState);
    const iDelta = src.indexOf('const deltaWeeks', loopStart);
    assert.ok(iState > loopStart && iTarget > iState && iNotReady > iTarget && iDelta > iNotReady,
      `Reihenfolge state@${iState} < target@${iTarget} < notReady@${iNotReady} < delta@${iDelta}`);
  });

  test('BUG 4: die neue Kontrollfall-Ursache wird im Ergebnis benannt', () => {
    assert.ok(/eventWait\.outsideRangePolls \? /.test(src),
      'die Zahl der "Events ausserhalb"-Polls gehoert in den Hinweistext');
    assert.ok(/rangeStartDate}–\$\{rangeEndDate\}/.test(src),
      'GUARD 2 nennt den konkreten Zeitraum, um den es ging');
  });

  test('resolveEventWaitOutcome wird in der Verdrahtung nicht mehr benutzt', () => {
    // Die alte Funktion hat polls und waitBudgetMs destrukturiert und nie
    // benutzt. Sie darf nicht mehr die Entscheidung treffen.
    const j = src.indexOf('const eventWaitReconciled = reconcileEventWait({');
    assert.ok(j > 0, 'reconcileEventWait muss die Entscheidung treffen');
    assert.ok(!/resolveEventWaitOutcome\(\{/.test(src.slice(src.indexOf('const weekVerified'), src.indexOf('// Defekt 2: dataQuality'))),
      'resolveEventWaitOutcome darf die Wartephase nicht mehr entscheiden');
  });

  test('der ehrliche Abbruchgrund wird auch berichtet, wenn die Navigation sauber war', () => {
    assert.ok(/const reportedAbortReason = effectiveAbortReason\(\{/.test(src));
    assert.ok(/abortReason: reportedAbortReason/.test(src), 'die Diagnostik muss ihn melden');
    assert.ok(/eventWaitAbortReason: eventWait\.abortReason/.test(src), 'und separat als Warte-Grund');
  });
});

// ===========================================================================
// Bug 4 — der Warte-Loop ist wochenblind (Live: polls 1, waitedMs 13,
//          eventWaitOutcome 'found' bei 18 Events der FALSCHEN Woche)
// ===========================================================================
// Live-Kennzahlen (Angefragt: KW 40 ab 2026-09-28):
//     eventWaitPolls: 1   waitedForEventsMs: 13   eventWaitBudgetMs: 20000
//     eventWaitOutcome: 'found'   abortReason: null
// 13/20000ms = 0,07 % des Budgets. `eventElementsFound > 0` war mit 18
// Event-Knoten der KW 39 erfuellt — die Abbruchbedingung war reine
// KNOTENZAHL und kann das Datum nicht unterscheiden. Der Range-Filter
// ('date-out-of-range') laeuft erst in der Abschluss-Sammlung, die
// Wartephase erfuhr die falsche Woche also nie.
//
// Diese Tests fahren den ECHTEN Draht: makeEventObserver (Produktion) ->
// collectCalendarEventLabels 'fast' (Produktion, echter evaluate-Callback
// gegen ein Fake-DOM) -> waitForCalendarEvents (Produktion).
// ===========================================================================
describe('Bug 4 — der Warte-Loop ist datumsqualifiziert', () => {
  // Ein Event der FALSCHEN Woche (KW 39, 21.09.) — exakt der Live-Fall.
  const WRONG_WEEK_CHIP = { attrs: { 'aria-label': 'MKK Nachbereitung Update Prod, 08:00 to 10:00, Monday September 21 2026, Busy' } };
  const TARGET_CHIP = EVENT_CHIP;
  // Der Wochen-Header der ZIELWOCHE, gerendert (KW 40, 28.09.–04.10.).
  const TARGET_HEADER = { text: '28. September – 4. Oktober', startDate: '2026-09-28', endDate: '2026-10-04' };

  test('KONTROLLFALL (Live): Events der FALSCHEN Woche -> KEIN found, Poll bis zum Budget', async () => {
    // Das ist der Kern der Anforderung. 18 Knoten der KW 39 bei angefragter
    // KW 40 duerfen NICHT als 'found' durchgehen.
    const doc = mutableDoc([node(WRONG_WEEK_CHIP), node(WRONG_WEEK_CHIP), node(WRONG_WEEK_CHIP)]);
    const clock = virtualClock();
    const res = await runWait(doc, {
      clock,
      // 18 Knoten, alle ausserhalb des angefragten Zeitraums (Live-Zahl).
      collect: async () => collectCalendarEventLabels(
        { evaluate: async (fn, arg) => fn(arg) },
        { weekStartIso: TARGET_WEEK_START, fast: true }
      ),
      readWeek: async () => null     // Header nie gerendert (Live-Zustand)
    });
    assert.equal(res.found, false, 'BUG 4: 18 Events der falschen Woche sind kein "found"');
    assert.notEqual(res.outcome, 'found');
    assert.equal(res.dateQualified, false);
    // Der Loop pollt BIS ZUM BUDGET (nicht 13ms/1 Poll).
    assert.ok(res.polls > 30, `es wurde bis zum Budget gepollt (war ${res.polls} Polls)`);
    assert.ok(res.waitedMs >= 20000 - FREEBUSY_EVENT_WAIT_POLL_MS,
      `das Budget wurde ausge schoepft (war ${res.waitedMs}ms, Live: 13ms)`);
    assert.ok(res.outsideRangePolls > 30, 'jeder Poll sah "Events ausserhalb des Zeitraums"');
    // Und der Abbruchgrund benennt genau das — nicht 'rendered-empty-proven'
    // (das waere eine Falschaussage: die Ansicht war nicht leer).
    assert.equal(res.abortReason, EVENT_WAIT_ABORT_REASONS.eventsOutsideRange);
    assert.equal(res.abortReason, 'events-outside-requested-range');
    assert.notEqual(res.abortReason, EVENT_WAIT_ABORT_REASONS.renderedEmpty,
      'Events ausserhalb des Zeitraums sind KEIN Leer-Beweis');
  });

  test('KONTROLLFALL: die Woche wechselt waehrend des Wartens -> found erst dann', async () => {
    // Der Kalender darf zwischen den Polls umschalten. 3 Polls falsche Woche,
    // dann ein Label aus dem angefragten Zeitraum.
    let n = 0;
    const doc = mutableDoc([node(WRONG_WEEK_CHIP)]);
    const clock = virtualClock();
    const res = await runWait(doc, {
      clock,
      collect: async () => {
        n++;
        doc.set(n <= 3 ? [node(WRONG_WEEK_CHIP)] : [node(TARGET_CHIP)]);
        return collectCalendarEventLabels(
          { evaluate: async (fn, arg) => fn(arg) },
          { weekStartIso: TARGET_WEEK_START, fast: true }
        );
      },
      readWeek: async () => null
    });
    assert.equal(res.found, true, 'nach dem Wechsel ist der Zeitraum belegt');
    assert.equal(res.outcome, 'found');
    assert.equal(res.dateQualified, true);
    assert.equal(res.evidence, 'in-range-label');
    assert.ok(res.polls >= 4, `es wurde ueber die Polls hinweg weitergemacht (war ${res.polls})`);
    assert.ok(res.outsideRangePolls >= 3, 'die ersten Polls waren ausserhalb — und wurden nicht als Treffer gezaehlt');
  });

  test('der Wochen-Header der Zielwoche qualifiziert auch OHNE Event-Label', async () => {
    // Belegkriterium (1): ein datierbarer Header der Zielwoche. Das ist der
    // Fall einer leeren Zielwoche — 'found' heisst hier "datumsqualifiziert
    // belegt", nicht "Events vorhanden".
    const doc = mutableDoc([]);
    const clock = virtualClock();
    const res = await runWait(doc, { clock, readWeek: async () => TARGET_HEADER });
    assert.equal(res.found, true);
    assert.equal(res.outcome, 'found');
    assert.equal(res.evidence, 'week-header');
    assert.equal(res.dateQualified, true);
  });

  test('der Wochen-Header einer FALSCHEN Woche qualifiziert NICHT', async () => {
    // Header gerendert, aber KW 39 statt KW 40 -> kein 'found'. Sonst waere der
    // Header-Pfad genauso wochenblind wie der Knotenzaehl-Pfad.
    const WRONG_HEADER = { text: '21.–27. September', startDate: '2026-09-21', endDate: '2026-09-27' };
    const doc = mutableDoc([node(WRONG_WEEK_CHIP)]);
    const clock = virtualClock();
    const res = await runWait(doc, { clock, readWeek: async () => WRONG_HEADER });
    assert.equal(res.found, false, 'ein Header der falschen Woche ist kein Beleg');
    assert.equal(res.outcome, 'budget-exhausted');
    assert.equal(res.abortReason, EVENT_WAIT_ABORT_REASONS.eventsOutsideRange);
  });

  test('OHNE Zielzeitraum: kein "found", sondern ein benannter Konfigurationsfehler', async () => {
    // Ein weekenblind konfigurierter Aufruf darf nicht still Erfolg melden.
    const doc = mutableDoc([node(TARGET_CHIP), node(TARGET_CHIP)]);
    const clock = virtualClock();
    const res = await runWait(doc, { clock, targetStart: null, range: null });
    assert.equal(res.found, false, 'ohne Zielwoche kann nichts belegt werden');
    assert.equal(res.outcome, 'not-found');
    assert.equal(res.abortReason, EVENT_WAIT_ABORT_REASONS.targetUnknown);
    assert.equal(res.polls, 0, 'und es wird auch nicht 20s auf eine Beleglosigkeit gepollt');
  });

  test('der Beobachter liest den Wochen-Header nur, wenn die Labels nicht qualifizieren', async () => {
    // Der Header-Read ist teuer (document.body.innerText). Er wird gedrosselt
    // und nur gebraucht, wenn inRangeLabels === 0.
    const doc = mutableDoc([node(TARGET_CHIP)]);
    let headerReads = 0;
    const observe = productionObserver({ readWeek: async () => { headerReads++; return TARGET_HEADER; } });
    const withLabel = await withFakeDocument(doc, () => observe());
    assert.equal(withLabel.inRangeLabels, 1, 'das Label qualifiziert allein');
    assert.equal(headerReads, 0, 'kein Header-Read noetig, es ist bereits belegt');
    assert.equal(withLabel.weekHeader, null);

    doc.set([node(WRONG_WEEK_CHIP)]);
    const withoutLabel = await withFakeDocument(doc, () => observe());
    assert.equal(withoutLabel.inRangeLabels, 0);
    assert.equal(headerReads, 1, 'jetzt wird der Header gelesen');
    assert.equal(withoutLabel.weekHeader.startDate, '2026-09-28');
  });

  test('countLabelsInRange benutzt dieselbe Prueflogik wie die Abschluss-Extraktion', () => {
    // Wenn die Wartephase eine andere Definition haette als die Extraktion,
    // koennte sie "fertig" melden, fuer das es am Ende keine Slots gibt.
    const labels = [
      'MKK Nachbereitung Update Prod, 08:00 to 10:00, Monday September 21 2026, Busy',   // KW 39
      'TeamSync, 10:00 to 11:00, Monday, September 28, 2026, Busy',                      // KW 40
      'Urlaub, all day event, Wednesday September 23 2026, Free'                        // KW 39, all-day
    ];
    const c = countLabelsInRange(labels, {
      rangeStartMs: TEST_RANGE.rangeStartMs,
      rangeEndMs: TEST_RANGE.rangeEndMs,
      defaultYear: 2026
    });
    assert.equal(c.inRange, 1, 'genau das Label der Zielwoche');
    assert.equal(c.outOfRange, 2, 'die beiden aus KW 39');
    assert.equal(c.undated, 0);
    assert.equal(c.unmeasured, false);
    // Ohne Range wird ausdruecklich "unmeasured" gemeldet, nicht 0.
    const u = countLabelsInRange(labels, { defaultYear: 2026 });
    assert.equal(u.unmeasured, true);
    assert.equal(u.inRange, 0);
  });

  test('eventWaitObservationState: fuenf Zustaende, disjunkt', () => {
    const q = (args) => eventWaitObservationState({ targetConfigured: true, targetStart: TARGET_WEEK_START, ...args });
    // Belegt, Header der Zielwoche
    assert.equal(q({ weekHeader: TARGET_HEADER, eventElementsFound: 0 }).category, EVENT_WAIT_STATES.targetWeek);
    // Belegt, Label im Zeitraum
    assert.equal(q({ inRangeLabels: 1, eventElementsFound: 4 }).category, EVENT_WAIT_STATES.inRange);
    // Spinner
    assert.equal(q({ loading: true, eventElementsFound: 0 }).category, EVENT_WAIT_STATES.stillLoading);
    // Events, alle ausserhalb — der Live-Fall
    assert.equal(q({ eventElementsFound: 18, inRangeLabels: 0 }).category, EVENT_WAIT_STATES.outsideRange);
    assert.equal(q({ eventElementsFound: 18, inRangeLabels: 0 }).dateQualified, false);
    // Nichts
    assert.equal(q({ eventElementsFound: 0 }).category, EVENT_WAIT_STATES.emptyCandidate);
    // Ohne Ziel: nie belegt
    assert.equal(eventWaitObservationState({ targetConfigured: false, eventElementsFound: 18, inRangeLabels: 3 }).category,
      EVENT_WAIT_STATES.targetUnknown);
    // Nur die ersten beiden sind datumsqualifiziert.
    assert.equal(q({ weekHeader: TARGET_HEADER }).dateQualified, true);
    assert.equal(q({ inRangeLabels: 1 }).dateQualified, true);
    for (const cat of [EVENT_WAIT_STATES.stillLoading, EVENT_WAIT_STATES.outsideRange, EVENT_WAIT_STATES.emptyCandidate]) {
      assert.equal(q({ eventElementsFound: 3, loading: cat === EVENT_WAIT_STATES.stillLoading }).dateQualified, false,
        `${cat} darf nicht qualifizieren`);
    }
  });

  test('der Header dominiert: Zielwoche-Header + 0 Labels -> found, nicht "leer"', async () => {
    const doc = mutableDoc([]);
    const clock = virtualClock();
    const res = await runWait(doc, { clock, readWeek: async () => TARGET_HEADER });
    // Ohne diese Regel wuerde der Leer-Beweis 'rendered-empty-proven' greifen
    // und dataQuality 'unverified-empty' melden — bei nachweislich richtiger,
    // leerer Woche.
    assert.notEqual(res.abortReason, EVENT_WAIT_ABORT_REASONS.renderedEmpty);
    assert.equal(res.outcome, 'found');
  });

  test('reconcileEventWait: datumsqualifiziert + 0 Knoten -> kein "budget-exhausted"', async () => {
    // Der Header war da, das Grid nicht lesbar. 'budget-exhausted' waere eine
    // Falschaussage (das Budget wurde nie ausgeschoepft).
    const r = reconcileEventWait({
      wait: { polls: 4, outcome: 'found', dateQualified: true, abortReason: null },
      eventElementsFound: 0, waitBudgetMs: 20000
    });
    assert.equal(r.outcome, 'found');
    assert.equal(r.abortReason, EVENT_WAIT_ABORT_REASONS.noNodesInTargetWeek);
    assert.equal(r.abortReason, 'no-event-nodes-in-target-week');
    assert.notEqual(r.abortReason, EVENT_WAIT_ABORT_REASONS.budgetExhausted);
  });

  test('reconcileEventWait: "wait-target-unknown" ist KEIN "not-polled"-Fehler', () => {
    // polls 0 bei Budget 20000 ist ein Logikfehler — AUSSER die Wartephase
    // durfte gar nicht starten, weil kein Ziel konfiguriert war.
    const r = reconcileEventWait({
      wait: { polls: 0, outcome: 'not-found', abortReason: EVENT_WAIT_ABORT_REASONS.targetUnknown },
      eventElementsFound: 0, waitBudgetMs: 20000
    });
    assert.equal(r.outcome, 'not-found');
    assert.equal(r.abortReason, EVENT_WAIT_ABORT_REASONS.targetUnknown);
    assert.notEqual(r.outcome, 'not-polled');
  });

  test('reconcileEventWait: Events ausserhalb des Zeitraums bleibt unterscheidbar', () => {
    const r = reconcileEventWait({
      wait: { polls: 40, outcome: 'budget-exhausted', abortReason: EVENT_WAIT_ABORT_REASONS.eventsOutsideRange },
      eventElementsFound: 0, waitBudgetMs: 20000
    });
    assert.equal(r.abortReason, 'events-outside-requested-range');
    assert.notEqual(r.abortReason, EVENT_WAIT_ABORT_REASONS.renderedEmpty);
  });

  test('der Hinweistext nennt je Abbruchgrund etwas anderes', () => {
    const outside = eventWaitWarning({ outcome: 'budget-exhausted', abortReason: EVENT_WAIT_ABORT_REASONS.eventsOutsideRange });
    assert.ok(outside && /Event-Knoten im DOM sichtbar/.test(outside), outside);
    assert.ok(!/KEINE Event-Knoten im DOM gelesen/.test(outside), 'das waere falsch — es waren ja Knoten da');
    assert.ok(/weder ein Treffer noch ein Leer-Beweis/.test(outside), 'beides wird gesagt');
    const noNodes = eventWaitWarning({ outcome: 'found', abortReason: EVENT_WAIT_ABORT_REASONS.noNodesInTargetWeek });
    assert.ok(noNodes && /nicht auslesbares Grid/.test(noNodes), noNodes);
    assert.ok(/KEINEN leeren Kalender/.test(noNodes), '0 Slots heissen nicht "frei"');
    const unknown = eventWaitWarning({ outcome: 'not-found', abortReason: EVENT_WAIT_ABORT_REASONS.targetUnknown });
    assert.ok(unknown && /Konfigurationsfehler/.test(unknown), unknown);
    // Die alten Faelle bleiben unveraendert.
    assert.equal(eventWaitWarning({ outcome: 'found', abortReason: null }), null);
    assert.ok(/KEINE Event-Knoten im DOM gelesen/.test(eventWaitWarning({ outcome: 'not-found', abortReason: null })));
    assert.ok(/Budget/.test(eventWaitWarning({ outcome: 'budget-exhausted', abortReason: null })));
  });

  test('das Ruhe-Fenster wird von "Events ausserhalb" NICHT ausgeloest', async () => {
    // Sonst wuerde nach settleMs 'rendered-empty-proven' gemeldet — bei 18
    // sichtbaren Events. Das ist die schlimmste moegliche Falschaussage:
    // "Kalender leer".
    let n = 0;
    const doc = mutableDoc([node(WRONG_WEEK_CHIP)]);
    const clock = virtualClock();
    const res = await runWait(doc, {
      clock,
      collect: async () => {
        n++;
        // 1 Poll ganz leer, danach dauerhaft Events der falschen Woche.
        doc.set(n <= 1 ? [] : [node(WRONG_WEEK_CHIP)]);
        return collectCalendarEventLabels(
          { evaluate: async (fn, arg) => fn(arg) },
          { weekStartIso: TARGET_WEEK_START, fast: true }
        );
      },
      readWeek: async () => null
    });
    assert.notEqual(res.abortReason, EVENT_WAIT_ABORT_REASONS.renderedEmpty);
    assert.equal(res.abortReason, EVENT_WAIT_ABORT_REASONS.eventsOutsideRange);
  });
});

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

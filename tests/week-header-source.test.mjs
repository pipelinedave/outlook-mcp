import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
  isWeekNavLabel,
  isWeekHeaderShaped,
  isDisplayedWeekSource,
  isHeaderReadVerified,
  parseWeekHeaderText,
  resolveRenderedWeekFromCandidates,
  readRenderedCalendarWeek,
  collectCalendarCandidates,
  collectCalendarEventLabels,
  renderedWeekContains,
  WEEK_NAV_LABEL_PATTERNS,
  WEEK_READ_SOURCES
} from '../src/owa.js';

// ===========================================================================
// FIX 1 (Live-Lauf 5, 26.09.2026): der angezeigte Wochen-Header wurde aus dem
// ARIA-Label des "Go to next week"-Buttons gelesen.
//
// Belege, beide aus demselben Lauf (Commit 3907f0d):
//   Aufruf 1: angefragt 21.09.–02.10., View stand KORREKT auf KW 39
//             (21.–27.09.) -> renderedWeek = 2026-09-28–10-04 (= next-Label)
//             -> Abbruch 'max-clicks' nach 6 Klicks (prev/next/prev/next),
//             budget.usedMs 8706. Die View war richtig und wurde weggezogen.
//   Aufruf 2: angefragt KW 40, View stand auf KW 39. next-Label =
//             "28 September – 4 October, 2026" = Zielwoche -> weekVerified:true,
//             dataQuality:'verified', 0 Klicks, View blieb KW 39 -> FALSCHPASS
//             (alle 10 Events als 'date-out-of-range' verworfen).
// Der echte Header lag im DOM: "21–27 September, 2026, Jump to a specific
// date or date range." (headerCandidates[3]) und als VERWORFENER Knoten
// "21–27 September, 2026 (Week 39) " mit reason 'header-misclassified'.
//
// Diese Tests fahren den ECHTEN Pfad: collectCalendarCandidates (Produktion,
// echter evaluate-Callback gegen ein Fake-DOM) -> resolveRenderedWeekFromCandidates
// -> readRenderedCalendarWeek. Keine nachgebaute Entschcheidung.
// ===========================================================================

// ---------------------------------------------------------------------------
// Mini-DOM. Der evaluate-Callback von collectCalendarCandidates nutzt
// document.querySelectorAll (Komma-Listen, Nachfahren, Tag, [attr], [attr="w"],
// [attr*="teil"] mit 'i'), window.getComputedStyle, offsetParent,
// getBoundingClientRect und document.body.innerText. Der Matcher bildet genau
// diese Selektorformen ab — er "versteht" keinen Selektor, den der Produktions-
// code nicht verwendet.
// ---------------------------------------------------------------------------
const MONTHISH = '[A-Za-z\\u00e4\\u00f6\\u00fc\\u00c4\\u00d6\\u00dc\\u00df]';

function simpleMatches(el, sel) {
  const s = String(sel).trim();
  if (!s) return false;
  if (s === '*') return true;
  const parts = s.match(/^([a-zA-Z][\w-]*)?((?:\[[^\]]*\])*)$/);
  if (!parts) return false;
  const [, tag, attrBlock] = parts;
  if (tag && String(el.tagName || '').toLowerCase() !== tag.toLowerCase()) return false;
  const attrRe = /\[([\w-]+)\s*(\*=|=)?\s*"([^"]*)"\s*(i)?\]/g;
  let m;
  while ((m = attrRe.exec(attrBlock || '')) !== null) {
    const [, name, op, value, ci] = m;
    const actual = el.getAttribute ? el.getAttribute(name) : null;
    if (actual == null) return false;
    const hay = ci ? String(actual).toLowerCase() : String(actual);
    const nee = ci ? value.toLowerCase() : value;
    if (op === '*=') { if (hay.indexOf(nee) < 0) return false; }
    else if (hay !== nee) return false;
  }
  return true;
}

function selectorMatches(el, sel) {
  // Komma-Listen: ein Selektor passt, wenn eine beliebige Alternative passt.
  const alternatives = String(sel).split(',').map((s) => s.trim()).filter(Boolean);
  for (const alt of alternatives) {
    const chain = alt.split(/\s+/).map((s) => s.trim());
    let node = el;
    let i = chain.length - 1;
    if (!simpleMatches(el, chain[i])) continue;
    i--;
    node = el.parent || null;
    let ok = true;
    while (i >= 0) {
      let found = false;
      while (node) {
        if (simpleMatches(node, chain[i])) { found = true; node = node.parent || null; break; }
        node = node.parent || null;
      }
      if (!found) { ok = false; break; }
      i--;
    }
    if (ok) return true;
  }
  return false;
}

const NODE = (over = {}) => ({
  tagName: over.tag || 'DIV',
  parent: over.parent || null,
  attrs: over.attrs || {},
  innerText: over.text || '',
  get textContent() { return this.innerText; },
  offsetParent: over.offsetParent === undefined ? {} : over.offsetParent,
  getBoundingClientRect: () => over.rect || { x: 0, y: 0, left: 0, top: 0, width: 300, height: 40 },
  getAttribute(name) {
    if (name === 'class') return over.className || this.attrs.class || null;
    if (name === 'id') return over.id || null;
    return this.attrs[name] !== undefined ? this.attrs[name] : null;
  }
});

const docOf = (nodes) => ({
  querySelectorAll: (sel) => nodes.filter((el) => selectorMatches(el, sel)),
  querySelector: (sel) => nodes.find((el) => selectorMatches(el, sel)) || null,
  body: { innerText: (nodes.map((n) => n.innerText).join('\n')) }
});

const fakePage = (nodes) => ({
  evaluate: async (fn, arg) => fn(arg),
  url: () => 'https://outlook.office.com/calendar/view/week?startdate=2026-09-21',
  __nodes: nodes,
  get __doc() { return docOf(this.__nodes); }
});

const withDom = async (nodes, fn) => {
  const page = fakePage(nodes);
  const prevDoc = globalThis.document;
  const prevWin = globalThis.window;
  globalThis.document = page.__doc;
  globalThis.window = { getComputedStyle: () => ({ display: 'block', visibility: 'visible', position: 'static' }) };
  try { return await fn(page); } finally {
    globalThis.document = prevDoc;
    globalThis.window = prevWin;
  }
};

// --- Die Live-Labels aus Lauf 5 (woertlich) ---------------------------------
// Der Navigations-Button, der als Header gelesen wurde:
const NEXT_BTN = 'Go to next week 28 September – 4 October, 2026';
// Der ECHTE angezeigte Header, der die ganze Zeit im DOM lag:
const DISPLAYED = '21–27 September, 2026, Jump to a specific date or date range.';
// Derselbe Header in der zweiten, in Lauf 5 als 'header-misclassified'
// VERWORFENEN Form:
const DISPLAYED_WEEKNUM = '21–27 September, 2026 (Week 39) ';
const PREV_BTN = 'Go to previous week 14–20 September, 2026 (Week 38)';

// Ein DOM, das genau dem Live-Zustand entspricht: echter Header + Navigations-
// Buttons. Der Header liegt in der Banner-Button-Region (so rendert OWA ihn).
const liveNodes = () => {
  const banner = NODE({ tag: 'DIV', attrs: { role: 'banner' } });
  const headerBtn = NODE({ tag: 'BUTTON', parent: banner, attrs: { 'aria-label': DISPLAYED } });
  const prevBtn = NODE({ tag: 'BUTTON', parent: banner, attrs: { 'aria-label': PREV_BTN } });
  const nextBtn = NODE({ tag: 'BUTTON', parent: banner, attrs: { 'aria-label': NEXT_BTN } });
  return [banner, headerBtn, prevBtn, nextBtn];
};

// ===========================================================================
describe('FIX 1 — isWeekNavLabel: Navigation ist per Definition kein Header', () => {

  test('EN-Navigations-Labels werden erkannt (belegte Live-Form)', () => {
    assert.equal(isWeekNavLabel(NEXT_BTN), true);
    assert.equal(isWeekNavLabel('Go to previous week 14–20 September, 2026 (Week 38)'), true);
    assert.equal(isWeekNavLabel('Go to next week 5–11 October, 2026 (Week 41)'), true);
    assert.equal(isWeekNavLabel('Go to the next week'), true);
  });

  test('DE-Navigations-Labels werden erkannt', () => {
    assert.equal(isWeekNavLabel('Nächste Woche'), true);
    assert.equal(isWeekNavLabel('Zur vorherigen Woche'), true);
    assert.equal(isWeekNavLabel('Nächsten Woche springen'), true);
    assert.equal(isWeekNavLabel('Wochennummer 41'), true);
  });

  test('EN: "Week number 41" ist Navigation', () => {
    assert.equal(isWeekNavLabel('Week number 41'), true);
    assert.equal(isWeekNavLabel('Go to week number 39'), true);
  });

  test('ein nacktes "Woche 41" ist AMBIGU und wird NICHT als Navigation gewertet', () => {
    // Bewusste Entscheidung: nur die belegten Navigations-Absichten
    // ("go to … week", "previous/next week", "week number NN", "Nächste/
    // Vorherige Woche") schliessen einen Header aus. Ein nacktes "Woche 41"
    // kann in einem Build auch der ANGEZEIGTE Header sein — es als Navigation
    // zu werten wuerde den echten Header blind machen (derselbe Fehlertyp wie
    // in Lauf 5, nur in die Gegenrichtung).
    assert.equal(isWeekNavLabel('Woche 41'), false);
    assert.equal(isWeekNavLabel('Week 41'), false);
    assert.equal(isWeekNavLabel('21–27 September, 2026 (Week 39)'), false);
  });

  test('der ANGEZEIGTE Header ist KEIN Nav-Label (beide Live-Formen)', () => {
    assert.equal(isWeekNavLabel(DISPLAYED), false,
      '"Jump to a specific date or date range" ist der Header, nicht die Navigation');
    assert.equal(isWeekNavLabel(DISPLAYED_WEEKNUM), false,
      '"(Week 39)" ohne Navigations-Absicht ist der Header — es wurde in Lauf 5 verworfen');
    assert.equal(isWeekNavLabel('28 September – 4 October, 2026'), false);
    assert.equal(isWeekNavLabel('28. September – 4. Oktober 2026'), false);
    assert.equal(isWeekNavLabel('Work week'), false);
  });

  test('Leer/Null/nicht-Text -> kein Nav-Label (defensiv)', () => {
    for (const v of [null, undefined, '', '   ', 0, false]) {
      assert.equal(isWeekNavLabel(v), false, `${JSON.stringify(v)}`);
    }
  });

  test('die Nav-Patterns sind Strings (serialisierbar ueber die evaluate-Grenze)', () => {
    assert.ok(Array.isArray(WEEK_NAV_LABEL_PATTERNS));
    assert.ok(WEEK_NAV_LABEL_PATTERNS.every((p) => typeof p === 'string'),
      'RegExp-Objekte koennen die page.evaluate-Grenze nicht passieren');
    for (const p of WEEK_NAV_LABEL_PATTERNS) assert.doesNotThrow(() => new RegExp(p, 'i'));
  });
});

// ===========================================================================
describe('FIX 1 — Bereich-Kriterium: ein Header ist ein ableitbarer Datumsbereich', () => {

  test('beide Live-Header-Formen liefern startDate+endDate', () => {
    const a = parseWeekHeaderText(DISPLAYED, new Date(2026, 8, 26));
    assert.equal(a.startDate, '2026-09-21');
    assert.equal(a.endDate, '2026-09-27');
    // Die Form, die vorher als 'header-misclassified' verworfen wurde — sie war
    // fuer den Parser UNLESBAR, deshalb griff der next-Button.
    const b = parseWeekHeaderText(DISPLAYED_WEEKNUM, new Date(2026, 8, 26));
    assert.equal(b.startDate, '2026-09-21');
    assert.equal(b.endDate, '2026-09-27');
  });

  test('der Nav-Button-Text IST zwar parsebar, aber kein Header (isWeekHeaderShaped=false)', () => {
    // Genau die Verwechslung: Form identisch, Absicht verschieden.
    assert.notEqual(parseWeekHeaderText(NEXT_BTN, new Date(2026, 8, 26)), null,
      'der Nav-Text ist parsebar — deshalb war er so gefaehrlich');
    assert.equal(isWeekHeaderShaped(NEXT_BTN), false, '… und darf trotzdem kein Header sein');
    assert.equal(isWeekHeaderShaped(DISPLAYED), true);
    assert.equal(isWeekHeaderShaped(DISPLAYED_WEEKNUM), true);
  });

  test('isWeekHeaderShaped verlangt einen BEREICH (kein Einzel-Datum, kein Event)', () => {
    assert.equal(isWeekHeaderShaped('Montag, 28. September'), false, 'Einzel-Datum ist kein Bereich');
    assert.equal(isWeekHeaderShaped('28.9.2026'), false);
    assert.equal(isWeekHeaderShaped('TeamSync, 10:00 to 11:00, Monday, September 28, 2026'), false);
    assert.equal(isWeekHeaderShaped('28.9. – 4.10.2026'), true, 'numerischer Bereich zaehlt');
    assert.equal(isWeekHeaderShaped(null), false);
  });
});

// ===========================================================================
// Der ECHTE aufrufende Pfad: collectCalendarCandidates (Browser-evaluate) ->
// resolveRenderedWeekFromCandidates -> readRenderedCalendarWeek.
// ===========================================================================
describe('FIX 1 — ECHTER Pfad: readRenderedCalendarWeek gegen das Live-DOM', () => {

  test('REGRESSION (Lauf 5 Aufruf 1): echter Header + next-Button -> die ANGEZEIGTE Woche', async () => {
    const r = await withDom(liveNodes(), (page) => readRenderedCalendarWeek(page));
    assert.equal(r.startDate, '2026-09-21', 'die View stand auf KW 39 — das ist der Header');
    assert.equal(r.endDate, '2026-09-27');
    assert.equal(r.source, WEEK_READ_SOURCES.displayed);
    assert.equal(r.readError, undefined, 'es wurde ein Header gelesen — kein Fehler');
    assert.equal(r.text, DISPLAYED, 'der Header-Text ist der angezeigte, nicht der next-Button');
    assert.ok(!/Go to next week/.test(r.text), 'der Navigations-Text darf NIE der Header sein');
  });

  test('REGRESSION (Lauf 5 Aufruf 2): die Ansicht war KW 39, angefragt KW 40 -> NICHT zertifiziert', async () => {
    // Der Falschpass. Vor dem Fix: renderedWeekContains(rendered,'2026-09-28')===true,
    // weil der next-Button die Zielwoche nannte -> weekVerified:true bei
    // View auf KW 39.
    const r = await withDom(liveNodes(), (page) => readRenderedCalendarWeek(page));
    assert.equal(renderedWeekContains(r, '2026-09-28'), false,
      'KW 39 darf KW 40 (28.09.–04.10.) nicht als "enthalten" melden');
    assert.equal(isHeaderReadVerified(r), true, 'der Header selbst IST gelesen (KW 39)');
    // Die Richtungs-Aussage der Navigation ist damit korrekt: KW 39 liegt
    // EINE Woche vor der Zielwoche.
    assert.equal(
      Math.round((new Date('2026-09-28T00:00:00') - new Date(r.startDate + 'T00:00:00')) / (7 * 86400000)), 1
    );
  });

  test('DOM OHNE angezeigten Header (nur Nav-Buttons) -> headerReadError, KEIN Button-Text als Woche', async () => {
    const banner = NODE({ tag: 'DIV', attrs: { role: 'banner' } });
    const nodes = [
      banner,
      NODE({ tag: 'BUTTON', parent: banner, attrs: { 'aria-label': PREV_BTN } }),
      NODE({ tag: 'BUTTON', parent: banner, attrs: { 'aria-label': NEXT_BTN } })
    ];
    const r = await withDom(nodes, (page) => readRenderedCalendarWeek(page));
    // Der Kern: das Klickziel des next-Buttons (KW 40) wird NIE zur gerenderten
    // Woche. Was bleibt, ist hoechstens die ABGEITUNG zwischen prev und next
    // (hier KW 39) — und die ist als Schätzung gekennzeichnet.
    assert.notEqual(r.startDate, '2026-09-28',
      'der next-Button darf NIE die gerenderte Woche werden');
    assert.equal(r.source, WEEK_READ_SOURCES.navDerived, 'nur eine Ableitung, kein gelesener Header');
    assert.equal(r.derivedFrom, 'nav-buttons');
    assert.equal(isHeaderReadVerified(r), false, 'FIX 1.4: weekVerified muss false sein');
    assert.equal(r.readError, 'no-displayed-header', 'der Grund wird BENANNT');
    assert.ok(Array.isArray(r.navOnlyCandidates) && r.navOnlyCandidates.includes(NEXT_BTN),
      'die Nav-Labels sind als solche sichtbar — sie sind nur kein Header');
  });

  test('die Sammelstufe trennt Nav- von Header-Kandidaten (Diagnose-Honestaet)', async () => {
    const c = await withDom(liveNodes(), (page) => collectCalendarCandidates(page));
    assert.ok(c.headerTexts.includes(DISPLAYED), 'der angezeigte Header ist ein Header-Kandidat');
    assert.ok(!c.headerTexts.includes(NEXT_BTN), 'der next-Button ist KEIN Header-Kandidat');
    assert.ok(!c.headerTexts.includes(PREV_BTN), 'der prev-Button auch nicht');
    assert.ok(c.navTexts.includes(NEXT_BTN) && c.navTexts.includes(PREV_BTN),
      'beide Nav-Labels stehen in weekNavCandidates (diagnostisch sichtbar)');
  });

  test('der verworfene echte Header wird als Header erkannt, nicht als Muell (header-misclassified)', async () => {
    // Der Knoten "21–27 September, 2026 (Week 39) " stand im Live-DOM und wurde
    // vom Kollektor als Event-Kandidat verworfen (reason 'header-misclassified').
    // Er ist der Wochen-Header.
    const chip = NODE({ tag: 'DIV', attrs: { role: 'button', 'aria-label': DISPLAYED_WEEKNUM } });
    const r = await withDom([chip], (page) => collectCalendarEventLabels(page, { weekStartIso: '2026-09-21' }));
    assert.deepEqual(r.weekHeaderLabels, [DISPLAYED_WEEKNUM.trim()]);
    assert.ok(!r.headerLikeLabels.includes(DISPLAYED_WEEKNUM.trim()),
      'er darf NICHT mehr als "header-misclassified" (Müll) geführt werden');
    assert.equal(r.eventElementsFound, 0, 'ein Header ist kein Event — die Extraktion bleibt unberührt');
  });

  test('der Jump-to-Header wird ebenfalls als Header und nicht als Muell gefuehrt', async () => {
    const chip = NODE({ tag: 'DIV', attrs: { role: 'button', 'aria-label': DISPLAYED } });
    const r = await withDom([chip], (page) => collectCalendarEventLabels(page, { weekStartIso: '2026-09-21' }));
    assert.deepEqual(r.weekHeaderLabels, [DISPLAYED]);
    assert.equal(r.headerLikeLabels.length, 0);
  });

  test('echte Events bleiben Events (FIX 1 fasst die Extraktion nicht an)', async () => {
    const chip = NODE({ tag: 'DIV', attrs: { 'aria-label': 'TeamSync, 10:00 to 11:00, Monday, September 28, 2026, Busy' } });
    const r = await withDom([chip], (page) => collectCalendarEventLabels(page, { weekStartIso: '2026-09-28' }));
    assert.equal(r.eventElementsFound, 1, 'der belegte Extraktionspfad funktioniert unverändert');
    assert.deepEqual(r.labels, ['TeamSync, 10:00 to 11:00, Monday, September 28, 2026, Busy']);
    assert.equal(r.weekHeaderLabels.length, 0);
  });
});

// ===========================================================================
describe('FIX 1 — Herkunft des Wochenwerts (source) und weekVerified', () => {

  test('nur ANGEZEIGTE Herkunft gilt als Wochen-Nachweis', () => {
    assert.equal(isDisplayedWeekSource(WEEK_READ_SOURCES.displayed), true);
    assert.equal(isDisplayedWeekSource(WEEK_READ_SOURCES.column), true);
    assert.equal(isDisplayedWeekSource(WEEK_READ_SOURCES.navDerived), false, 'ABGELEITET ist nicht GELESEN');
    assert.equal(isDisplayedWeekSource(WEEK_READ_SOURCES.singleDate), false);
    assert.equal(isDisplayedWeekSource(WEEK_READ_SOURCES.numericDay), false);
    assert.equal(isDisplayedWeekSource(null), false);
  });

  test('isHeaderReadVerified: nav-derivation ist KEIN Nachweis (FIX 1.4)', () => {
    const derived = resolveRenderedWeekFromCandidates({
      headerTexts: [], dayHeaderTexts: [],
      navTexts: ['Go to previous week 14–20 September, 2026 (Week 38)', 'Go to next week 28 September – 4 October, 2026']
    }, new Date(2026, 8, 26));
    // Die Ableitung liefert plausibel KW 39 …
    assert.equal(derived.startDate, '2026-09-21');
    assert.equal(derived.source, WEEK_READ_SOURCES.navDerived);
    // … ist aber eine SCHÄTZUNG und darf keine 'weekVerified' erkaufen.
    assert.equal(isHeaderReadVerified(derived), false);
  });

  test('der angezeigte Header IST ein Nachweis', () => {
    const read = resolveRenderedWeekFromCandidates({
      headerTexts: [DISPLAYED, NEXT_BTN], dayHeaderTexts: []
    }, new Date(2026, 8, 26));
    assert.equal(read.source, WEEK_READ_SOURCES.displayed);
    assert.equal(isHeaderReadVerified(read), true);
  });

  test('nichts Datierbares -> readError benannt, source null, kein "erfolgreiches" reading', () => {
    const out = resolveRenderedWeekFromCandidates({ headerTexts: ['Work week'], dayHeaderTexts: [] }, new Date(2026, 8, 26));
    assert.equal(out.startDate, null);
    assert.equal(out.source, null);
    assert.equal(out.readError, 'no-displayed-header');
    assert.equal(isHeaderReadVerified(out), false);
  });

  test('DER FEHLER ALS ANKER: die alte Aufloesung haette den next-Button genommen', () => {
    // Die exakte Vorher-Logik: erste parsebare Range ueber alle Kandidaten.
    const oldResolve = (headerTexts) => {
      for (const text of headerTexts) {
        const p = parseWeekHeaderText(text, new Date(2026, 8, 26));
        if (p) return p;
      }
      return null;
    };
    const headerTexts = [NEXT_BTN, DISPLAYED];
    const alt = oldResolve(headerTexts);
    const neu = resolveRenderedWeekFromCandidates({ headerTexts, dayHeaderTexts: [] }, new Date(2026, 8, 26));
    assert.equal(alt.startDate, '2026-09-28', 'alt: der next-Button (KW 40) — das war der Falschpass');
    assert.equal(neu.startDate, '2026-09-21', 'neu: der angezeigte Header (KW 39)');
    assert.notEqual(alt.startDate, neu.startDate);
  });
});

// ===========================================================================
// Verdrahtungs-Wächter: die Entscheidung muss im AUFRUFER verdrahtet sein.
// Ein korrektes isWeekNavLabel, das niemand benutzt, faengt den Live-Bug nicht.
// ===========================================================================
describe('FIX 1 — Verdrahtung in getFreeBusy', () => {
  const src = readFileSync(fileURLToPath(new URL('../src/owa.js', import.meta.url)), 'utf8');

  test('die Aufloesung filtert Nav-Labels VOR jeder Range-/Text-Heuristik', () => {
    const i = src.indexOf('export function resolveRenderedWeekFromCandidates(');
    assert.ok(i > 0);
    const fn = src.slice(i, src.indexOf('// Liest die gerenderte Kalenderwoche', i));
    const iFilter = fn.indexOf('headers.filter((t) => !isWeekNavLabel(t))');
    const iRung1 = fn.indexOf('parseWeekHeaderText(text, now)');
    assert.ok(iFilter > 0 && iRung1 > 0);
    assert.ok(iFilter < iRung1, 'der Nav-Filter muss VOR der Bereichs-/Text-Heuristik stehen');
  });

  test('weekVerified ist an die ANGEZEIGTE Herkunft gebunden, nicht an startDate/endDate', () => {
    assert.ok(src.includes('const weekVerified = isHeaderReadVerified(rendered);'),
      'der Nachweis muss isHeaderReadVerified sein — sonst erkauft die Nav-Ableitung weekVerified');
    assert.ok(!/const weekVerified = Boolean\(rendered\.startDate && rendered\.endDate\)/.test(src),
      'die alte Form (belegt im Falschpass Lauf 5) darf nicht zurueckkehren');
  });

  test('der Browser-Collector schliesst Nav-Labels von den Header-Kandidaten aus', () => {
    const i = src.indexOf('export async function collectCalendarCandidates(page)');
    assert.ok(i > 0);
    const fn = src.slice(i, src.indexOf('// Sammelt reiche Diagnostik', i));
    assert.ok(/pushHeader\(headerTexts, text\)/.test(fn), 'Schritt 1/2/3 muessen ueber pushHeader laufen');
    assert.ok(/if \(isNavLabel\(label\)\) push\(navTexts, label\)/.test(fn),
      'Nav-Labels muessen in navTexts landen');
    assert.ok(!/push\(headerTexts, label\)/.test(fn), 'kein direkter push in headerTexts mehr');
    // Und die Pattern kommen als Argument ueber die evaluate-Grenze.
    assert.ok(/\}, WEEK_NAV_LABEL_PATTERNS\)/.test(fn),
      'die Nav-Patterns muessen serialisiert uebergeben werden (kein ReferenceError)');
  });

  test('der Event-Kollektor verwirft den echten Header nicht mehr als "header-misclassified"', () => {
    assert.ok(/if \(isWeekHeaderShaped\(t\)\) return 'displayed-week-header';/.test(src));
    assert.ok(/weekHeaderLabels: cap\(dedupe\(weekHeaderLike\), 10\)/.test(src),
      'der Header muss als eigener Beleg zurueckkommen');
    assert.ok(!/for \(const label of eventCollection\.weekHeaderLabels\) reject\(/.test(src),
      'der Header wird NICHT verworfen');
  });

  test('der Hinweistext behauptet nicht mehr, die URL belege die View', () => {
    // Lauf 5: urlAfterNavigation korrekt (startdate=2026-09-28), View KW 39.
    // Die alte Formel behauptete "startdate-Param ist Primärsignal".
    assert.ok(!/startdate-Param ist Primärsignal/.test(src), 'die Falschaussage ist entfernt');
    assert.ok(/der angezeigte Wochen-Header ist der EINZIGE Nachweis/.test(src));
  });
});

import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { OwaClient } from '../src/owa.js';
import { browserManager } from '../src/browser.js';

// ===========================================================================
// ECHTER AUFRUFENDER PFAD: OwaClient.getFreeBusy().
//
// Bisher testete jede Datei die ENTSCHEIDUNGSFUNKTIONEN. Das ist genau die
// Fehlerklasse, die in Lauf 5 zugeschlagen hat: resolveDataQuality und
// isBelievableBusySlots waren "korrekt", verdrahtet war die alte
// Ein-Zeilen-Form. Diese Datei ruft die echte Methode auf und prueft das
// ERGEBNIS-OBJEKT, das der MCP-Client sieht.
//
// Der Browser wird durch einen Fake-`page` ersetzt (browserManager.ensureContext
// ist patchbar, weil browserManager eine mutierbare Instanz ist). page.evaluate
// ist der ECHTE evaluate-Callback gegen ein Fake-DOM — dieselbe Komposition
// wie in Produktion, nur ohne Chromium.
//
// Das Fake modelliert das, was Lauf 5 belegt hat:
//   - der Wochen-Header traegt die ANGEZEIGTE Woche (nicht das Klickziel),
//   - "Go to next/previous week" aendert die angezeigte Woche beim Klick,
//   - der harte Reload der startdate-URL aendert die View NICHT (belegt:
//     urlAfterNavigation korrekt, View unveraendert).
// ===========================================================================

function simpleMatches(el, sel) {
  const s = String(sel).trim();
  if (!s) return false;
  if (s === '*') return true;
  const m = /^([a-zA-Z][\w-]*)?((?:\[[^\]]*\])*)$/.exec(s);
  if (!m) return false;
  const [, tag, attrBlock] = m;
  if (tag && String(el.tagName || '').toLowerCase() !== tag.toLowerCase()) return false;
  const attrRe = /\[([\w-]+)\s*(\*=|=)?\s*"([^"]*)"\s*(i)?\]/g;
  let a;
  while ((a = attrRe.exec(attrBlock || '')) !== null) {
    const [, name, op, value, ci] = a;
    const actual = el.getAttribute ? el.getAttribute(name) : null;
    if (actual == null) return false;
    const hay = ci ? String(actual).toLowerCase() : String(actual);
    const nee = ci ? value.toLowerCase() : value;
    if (op === '*=') { if (hay.indexOf(nee) < 0) return false; }
    else if (hay !== nee) return false;
  }
  return true;
}

// Whitespace-Zerlegung NUR ausserhalb von Klammern/Quotes: die Produktions-
// Selektoren enthalten Leerzeichen im Attributwert ("[aria-label*="Next week" i]").
const splitTop = (s) => {
  const out = [];
  let depth = 0;
  let q = false;
  let cur = '';
  for (const ch of s) {
    if (ch === '"') q = !q;
    if (!q) {
      if (ch === '[') depth++;
      else if (ch === ']') depth--;
      else if (depth === 0 && /\s/.test(ch)) {
        if (cur) { out.push(cur); cur = ''; }
        continue;
      }
    }
    cur += ch;
  }
  if (cur) out.push(cur);
  return out;
};

function selectorMatches(el, sel) {
  for (const alt of String(sel).split(',').map((s) => s.trim()).filter(Boolean)) {
    const chain = splitTop(alt);
    let i = chain.length - 1;
    if (!simpleMatches(el, chain[i])) continue;
    let node = el.parent || null;
    i--;
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

const mkNode = (over = {}) => ({
  tagName: over.tag || 'DIV',
  parent: over.parent || null,
  attrs: over.attrs || {},
  innerText: over.text || '',
  get textContent() { return this.innerText; },
  offsetParent: over.offsetParent === undefined ? {} : over.offsetParent,
  getBoundingClientRect: () => over.rect || { x: 0, y: 0, left: 0, top: 0, width: 300, height: 40 },
  getAttribute(name) {
    if (name === 'class') return over.className || (this.attrs.class || null);
    return this.attrs[name] !== undefined ? this.attrs[name] : null;
  }
});

const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const addDays = (isoDate, n) => iso(new Date(new Date(isoDate + 'T00:00:00').getTime() + n * 86400000));
const MONTH_EN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// Die OWA-Form des angezeigten Headers (Live-Beleg 26.09.2026). WICHTIG: der
// Monat gehoert zu SEINEM Tag — eine Woche 28.09.–04.10. lautet
// "28 September – 4 October, 2026", nicht "28–4 September".
const weekLabel = (mondayIso) => {
  const a = new Date(mondayIso + 'T00:00:00');
  const b = new Date(a.getTime() + 6 * 86400000);
  return `${a.getDate()} ${MONTH_EN[a.getMonth()]} – ${b.getDate()} ${MONTH_EN[b.getMonth()]}, ${a.getFullYear()}`;
};

const buildCalendar = ({ displayedMonday, events = [], eventsByWeek = null, headerHint = true, miswired = null }) => {
  const state = { monday: displayedMonday, clicks: [], reloads: 0, gotos: [], headerGone: false };
  // Events der ANGEZEIGTEN Woche — so wie im echten Kalender: navigiert man
  // eine Woche vorwaerts, zeigen sich deren Termine.
  const labelsFor = () => (typeof eventsByWeek === 'function' ? eventsByWeek(state.monday) : events);
  const nodes = () => {
    const banner = mkNode({ tag: 'DIV', attrs: { role: 'banner' } });
    const list = [banner];
    // 'unreadable-after-click': der Header verschwindet nach dem ersten Klick
    // (nicht bewertbarer Klick). 'backwards-after-click': der Button laeuft in
    // die Gegenrichtung (uebersprungene Zielwoche).
    if (headerHint && !state.headerGone) {
      list.push(mkNode({
        tag: 'BUTTON', parent: banner,
        attrs: { 'aria-label': `${weekLabel(state.monday)}, Jump to a specific date or date range.` }
      }));
    }
    list.push(mkNode({ tag: 'BUTTON', parent: banner, attrs: { role: 'button', 'aria-label': `Go to previous week ${weekLabel(addDays(state.monday, -7))}` } }));
    list.push(mkNode({ tag: 'BUTTON', parent: banner, attrs: { role: 'button', 'aria-label': `Go to next week ${weekLabel(addDays(state.monday, 7))}` } }));
    for (const label of labelsFor()) list.push(mkNode({ tag: 'DIV', attrs: { role: 'button', 'aria-label': label } }));
    return list;
  };
  const doc = () => {
    const current = nodes();
    return {
      querySelectorAll: (sel) => current.filter((el) => selectorMatches(el, sel)),
      querySelector: (sel) => current.find((el) => selectorMatches(el, sel)) || null,
      body: { innerText: '' }
    };
  };
  let currentUrl = 'https://outlook.office.com/calendar/view/week';
  const page = {
    url: () => currentUrl,
    async goto(url) { state.gotos.push(url); currentUrl = url; return true; },
    // Belegt aus Lauf 5: der harte Reload aendert die View NICHT.
    async reload() { state.reloads++; return true; },
    async waitForTimeout() { /* keine echte Wartezeit noetig */ },
    // JEDES evaluate sieht den AKTUELLEN DOM-Zustand — genau wie im Browser.
    // Ohne das waere der Header nach einem Klick ein Snapshot und die
    // Verifikation am angezeigten Header koennte nie eine Aenderung sehen.
    async evaluate(fn, arg) { globalThis.document = doc(); return fn(arg); },
    locator(selector) {
      return {
        first() {
          return {
            async waitFor() {
              if (!doc().querySelectorAll(selector)[0]) throw new Error('not visible');
            },
            async click() {
              const el = doc().querySelectorAll(selector)[0];
              if (!el) throw new Error('not found');
              const label = String(el.getAttribute('aria-label') || '');
              if (/Go to next week/i.test(label)) {
                if (miswired === 'unreadable-after-click') state.headerGone = true;
                state.monday = addDays(state.monday, miswired === 'backwards-after-click' ? -7 : 7);
                state.clicks.push('next');
              }
              else if (/Go to previous week/i.test(label)) { state.monday = addDays(state.monday, -7); state.clicks.push('prev'); }
              else state.clicks.push('other');
            }
          };
        }
      };
    }
  };
  return { page, state, doc };
};

const realEnsure = browserManager.ensureContext.bind(browserManager);
let activeCalendar = null;
const prevDoc = globalThis.document;
const prevWin = globalThis.window;

describe('ECHTER PFAD: OwaClient.getFreeBusy — Ergebnis-Invarianten', () => {
  before(() => {
    browserManager.ensureContext = async () => ({ page: activeCalendar.page, context: {} });
    browserManager.resetForTest();
    globalThis.window = { getComputedStyle: () => ({ display: 'block', visibility: 'visible', position: 'static' }) };
  });
  after(() => {
    browserManager.ensureContext = realEnsure;
    globalThis.document = prevDoc;
    globalThis.window = prevWin;
  });

  const run = async (cal, start, end) => {
    activeCalendar = cal;
    const client = new OwaClient();
    return client.getFreeBusy({ start, end });
  };

  // =====================================================================
  // FIX 4 — busySlotsReliable vs. die eigene Widerspruchs-Warnung
  // =====================================================================
  test('FIX 4: 6 echte Event-Knoten, 0 Busy-Slots -> busySlotsReliable === false', async () => {
    // ZIELWOCHE steht im Header. Die 6 Knoten sind "Free"-Marker (Urlaub/Frei):
    // parsebar, also Event-Knoten, aber per Definition keine BELEGUNG.
    // dataQuality ist damit 'verified' (Woche verifiziert, Knoten gelesen) und
    // busySlots ist leer.
    // In Lauf 2 stand hier busySlotsReliable:true GLEICHZEITIG mit der eigenen
    // Warnung "18 Event-Element(e) im DOM gefunden, aber 0 Busy-Slots extrahiert".
    const events = [
      'Urlaub, all day event, Monday, September 28, 2026, Free',
      'Urlaub, all day event, Tuesday, September 29, 2026, Free',
      'Urlaub, all day event, Wednesday, September 30, 2026, Free',
      'Frei, all day event, Thursday, October 1, 2026, Free',
      'Frei, all day event, Friday, October 2, 2026, Free',
      'Frei, all day event, Saturday, October 3, 2026, Free'
    ];
    const cal = buildCalendar({ displayedMonday: '2026-09-28', events });
    const res = await run(cal, '2026-09-28', '2026-10-04');
    assert.equal(res.status, 'ok', JSON.stringify(res.error || res.note).slice(0, 400));
    assert.ok(res.diagnostics, 'es MUSS Diagnostik geben (unzuverlaessig)');
    assert.equal(res.diagnostics.eventElementsFound, 6, `6 Event-Knoten gelesen (war ${res.diagnostics.eventElementsFound})`);
    assert.equal(res.busySlots.length, 0);
    assert.match(res.diagnostics.warning || '', /Widerspruch: 6 Event-Element\(e\) im DOM gefunden, aber 0 Busy-Slots extrahiert/,
      'die eigene Widerspruchs-Warnung entsteht');
    assert.equal(res.busySlotsReliable, false,
      'INVARIANTE: Widerspruchswarnung => busySlotsReliable === false');
    // Die Woche selbst war verifiziert — der Fix darf nichts ueberkorrigieren.
    assert.equal(res.weekVerified, true);
    assert.equal(res.dataQuality, 'verified');
  });

  test('FIX 4: ohne Widerspruch bleibt busySlotsReliable true (kein Ueberkorrigieren)', async () => {
    const events = [
      'TeamSync, 10:00 to 11:00, Monday, September 28, 2026, Busy',
      'Review, 14:00 to 15:00, Tuesday, September 29, 2026, Busy'
    ];
    const cal = buildCalendar({ displayedMonday: '2026-09-28', events });
    const res = await run(cal, '2026-09-28', '2026-10-04');
    assert.equal(res.status, 'ok', JSON.stringify(res.error || res.note).slice(0, 400));
    assert.equal(res.busySlots.length, 2, `zwei Slots extrahiert (war ${res.busySlots.length})`);
    assert.equal(res.busySlotsReliable, true, 'ohne Widerspruch bleibt es zuverlaessig');
    assert.equal(res.diagnostics, undefined, 'und es wird keine Diagnostik beigelegt');
    assert.equal(cal.state.clicks.length, 0, 'die Zielwoche stand bereits im Header — 0 Klicks');
  });

  // =====================================================================
  // FIX 1/2/3 im Ergebnis: die Navigation ist gerichtet, verifiziert, begrenzt
  // =====================================================================
  test('FIX 1+3: eine Woche zurueck -> genau EIN Klick, dann Zielwoche gelesen', async () => {
    // Die View steht auf KW 39 (21.-27.09.), angefragt ist KW 40 (28.09.-04.10.).
    // Genau der Live-Fall aus Lauf 5 Aufruf 1 - dort wurden 6 Klicks in einer
    // Schleife (prev/next/prev/next) abgesetzt und dann 'max-clicks' gemeldet.
    const eventsByWeek = (mondayIso) => (mondayIso === '2026-09-28'
      ? ['TeamSync, 10:00 to 11:00, Monday, September 28, 2026, Busy',
        'Review, 14:00 to 15:00, Tuesday, September 29, 2026, Busy']
      : ['UKO TP GM, 11:00 to 12:00, Monday, September 21, 2026, Busy',
        'MKK Nachbereitung, 08:00 to 10:00, Monday, September 21, 2026, Busy']);
    const cal = buildCalendar({ displayedMonday: '2026-09-21', eventsByWeek });
    const res = await run(cal, '2026-09-28', '2026-10-04');
    // Ein Klick "next" bringt von KW 39 auf KW 40; danach wird die Zielwoche
    // am ANGEZEIGTEN Header verifiziert und es wird NICHT weitergeklickt.
    assert.deepEqual(cal.state.clicks, ['next'], `genau ein Klick in eine Richtung (war ${JSON.stringify(cal.state.clicks)})`);
    assert.equal(res.status, 'ok', JSON.stringify(res.error || res.note).slice(0, 400));
    assert.equal(res.renderedWeek.startDate, '2026-09-28', 'die Zielwoche ist angezeigt');
    assert.equal(res.renderedWeek.endDate, '2026-10-04');
    assert.equal(res.weekVerified, true);
    assert.equal(res.busySlots.length, 2, 'die Events der ZIELWOCHE sind extrahiert');
    assert.equal(res.busySlotsReliable, true);
    assert.equal(res.diagnostics, undefined, 'sauberer Lauf -> keine Diagnostik');
  });

  test('FIX 3: die Navigation ist auf 2 Klicks begrenzt (kein Blindklicken)', async () => {
    // Die View steht DREI Wochen vor der Zielwoche. Mit 6 Klicks (alt) waeren
    // das 3 Klicks vorwaerts; erlaubt sind laut Anforderung 2.
    const cal = buildCalendar({
      displayedMonday: '2026-09-07',
      events: ['TeamSync, 10:00 to 11:00, Monday, September 7, 2026, Busy']
    });
    const res = await run(cal, '2026-09-28', '2026-10-04');
    assert.ok(cal.state.clicks.length <= 2, `maximal 2 Klicks (war ${cal.state.clicks.length}: ${JSON.stringify(cal.state.clicks)})`);
    // Der ehrliche Abbruchgrund, kein stilles Durchrutschen in "falsche Daten".
    assert.notEqual(res.status, 'ok', 'die Zielwoche wurde nicht erreicht — das muss als Abbruch sichtbar sein');
    assert.ok(/max-clicks|too-far/.test((res.diagnostics && res.diagnostics.abortReason) || ''),
      `ehrlicher Abbruchgrund (war ${res.diagnostics && res.diagnostics.abortReason})`);
    assert.ok(!res.busySlots || res.busySlots.length === 0, 'es werden keine Daten der falschen Woche zurueckgegeben');
  });

  test('FIX 3: der Reload wird versucht, aber NICHT als Beleg der View gewertet', async () => {
    // Lauf 5: urlAfterNavigation korrekt (startdate=2026-09-28), View KW 39.
    // Der Reload darf die Beweislast nicht auf die URL verlagern.
    const cal = buildCalendar({
      displayedMonday: '2026-09-21',
      events: ['UKO TP GM, 11:00 to 12:00, Monday, September 21, 2026, Busy']
    });
    const res = await run(cal, '2026-09-28', '2026-10-04');
    // Egal ob der Reload zuegig: die View wird am Header verifiziert.
    assert.equal(res.renderedWeek.startDate, '2026-09-28', 'die View ist am Header verifiziert, nicht an der URL');
    assert.ok(!/Primärsignal/.test(res.note || ''), 'der Hinweistext behauptet keine URL-Garantie');
  });

  test('FIX 3: ein Button in die Gegenrichtung -> ehrlicher Abbruch, kein Weiterklicken', async () => {
    // "Go to next week" bewegt die View ZURUECK. Der Klick hat damit die
    // Zielwoche uebersprungen — weiterzuklicken waere Blindklicken (in Lauf 5
    // sind daraus 6 Klicks prev/next/prev/next geworden).
    const cal = buildCalendar({
      displayedMonday: '2026-09-21',
      miswired: 'backwards-after-click',
      events: ['UKO TP GM, 11:00 to 12:00, Monday, September 21, 2026, Busy']
    });
    const res = await run(cal, '2026-09-28', '2026-10-04');
    assert.equal(cal.state.clicks.length, 1, `genau EIN Klick, dann Abbruch (war ${JSON.stringify(cal.state.clicks)})`);
    assert.notEqual(res.status, 'ok', 'die Zielwoche wurde nicht erreicht');
    assert.equal(res.diagnostics.abortReason, 'week-overshoot',
      'der Abbruch benennt die tatsaechliche Ursache');
  });

  test('FIX 3: Header nach dem Klick unlesbar -> ehrlicher Abbruch, kein "changed"-Fortsetzen', async () => {
    // In Lauf 5 hielt genau das die Schleife am Leben: 'changed' ohne Zielbezug.
    const cal = buildCalendar({
      displayedMonday: '2026-09-21',
      miswired: 'unreadable-after-click',
      events: ['UKO TP GM, 11:00 to 12:00, Monday, September 21, 2026, Busy']
    });
    const res = await run(cal, '2026-09-28', '2026-10-04');
    assert.equal(cal.state.clicks.length, 1, `genau EIN Klick, dann Abbruch (war ${JSON.stringify(cal.state.clicks)})`);
    assert.equal(res.diagnostics.abortReason, 'click-unverified',
      'ein nicht bewertbarer Klick wird benannt, nicht durch Wiederholung "geprueft"');
    // Und das Ergebnis behauptet nichts, was es nicht weiss: ohne lesbaren
    // Header gibt es keine Wochen-Verifikation und keine belastbaren Slots.
    assert.equal(res.weekVerified, false);
    assert.equal(res.busySlotsReliable, false);
    assert.equal(res.diagnostics.headerReadError, 'no-displayed-header');
  });

  // =====================================================================
  // FIX 2 im Ergebnis: Header der falschen Woche + Events ausserhalb
  // =====================================================================
  test('FIX 2: Header der falschen Woche -> 0 Busy-Slots und unzuverlaessig', async () => {
    // Die View bleibt auf KW 39 (der Klickpfad erreicht sie nicht, weil der
    // Aufruf bei 1 Klick abbrichst bzw. der Header KW 39 sagt) und die 2 Events
    // liegen in KW 39. Nach FIX 2 ist das ein Gegenbeleg: busySlots bleibt
    // leer und unzuverlaessig.
    const cal = buildCalendar({
      displayedMonday: '2026-09-21',
      events: [
        'UKO TP GM, 11:00 to 12:00, Monday, September 21, 2026, Busy',
        'MKK Nachbereitung, 08:00 to 10:00, Tuesday, September 22, 2026, Busy'
      ]
    });
    const res = await run(cal, '2026-09-28', '2026-10-04');
    assert.equal(res.busySlots.length, 0, 'keine Events der falschen Woche als Belegung');
    assert.equal(res.busySlotsReliable, false);
    assert.ok(res.diagnostics, 'Diagnostik muss den Gegenbeleg zeigen');
  });
});

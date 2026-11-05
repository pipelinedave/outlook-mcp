import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { OwaClient } from '../src/owa.js';
import { browserManager } from '../src/browser.js';

// ===========================================================================
// ECHTER AUFRUFENDER PFAD: OwaClient.listCalendarEvents().
//
// Das READ-ONLY Kalender-Lese-Tool (gbrain #5916, otto Termin-Briefings) nutzt
// dieselben verifizierten Navigations-Primitive wie getFreeBusy, gibt aber
// ALLE Termine zurueck (busy, free, tentative, Ganztag), nicht nur Busy-Slots.
//
// Der Browser wird wie in getfreebusy-result.test.mjs durch einen Fake-`page`
// ersetzt (browserManager.ensureContext ist patchbar). page.evaluate ist der
// ECHTE evaluate-Callback gegen ein Fake-DOM.
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

const weekLabel = (mondayIso) => {
  const a = new Date(mondayIso + 'T00:00:00');
  const b = new Date(a.getTime() + 6 * 86400000);
  return `${a.getDate()} ${MONTH_EN[a.getMonth()]} – ${b.getDate()} ${MONTH_EN[b.getMonth()]}, ${a.getFullYear()}`;
};

const buildCalendar = ({ displayedMonday, events = [], eventsByWeek = null, hydrateTo = null }) => {
  // hydrateTo (Runde 3): { afterMs, labels } — bis afterMs werden `events`
  // geliefert, danach `labels`. Simuliert OWA-Lazy-Hydration: Header sofort
  // da, Event-Grid faellt gestaffelt ein. Die Uhr ist VIRTUELL: das Fake-
  // page.waitForTimeout dreht state.fakeMs weiter (reale Wanduhr bleibt in
  // Tests bei ~0 — daher nicht Date.now()).
  const state = { monday: displayedMonday, clicks: [], fakeMs: 0, startFakeMs: null };
  const labelsFor = () => {
    if (hydrateTo && Array.isArray(hydrateTo.labels)) {
      if (state.startFakeMs === null) state.startFakeMs = state.fakeMs;
      return (state.fakeMs - state.startFakeMs) >= hydrateTo.afterMs ? hydrateTo.labels : events;
    }
    return typeof eventsByWeek === 'function' ? eventsByWeek(state.monday) : events;
  };
  const nodes = () => {
    const banner = mkNode({ tag: 'DIV', attrs: { role: 'banner' } });
    const list = [banner];
    list.push(mkNode({
      tag: 'BUTTON', parent: banner,
      attrs: { 'aria-label': `${weekLabel(state.monday)}, Jump to a specific date or date range.` }
    }));
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
    async goto(url) { currentUrl = url; return true; },
    async reload() { return true; },
    async waitForTimeout(ms) { state.fakeMs += Number(ms) || 0; },
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
              if (/Go to next week/i.test(label)) { state.monday = addDays(state.monday, 7); state.clicks.push('next'); }
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
export { buildCalendar };

const realEnsure = browserManager.ensureContext.bind(browserManager);
let activeCalendar = null;
const prevDoc = globalThis.document;
const prevWin = globalThis.window;

describe('READ-ONLY Kalender-Lese-Tool: OwaClient.listCalendarEvents', () => {
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
    return client.listCalendarEvents({ start, end });
  };

  test('listet ALLE Termine (busy, free, tentative, Ganztag) im Zeitraum', async () => {
    const cal = buildCalendar({
      displayedMonday: '2026-09-28',
      events: [
        'TeamSync, 10:00 to 11:00, Monday, September 28, 2026, Busy',
        'Urlaub, all day event, Tuesday, September 29, 2026, Free',
        'Review, 14:00 to 15:00, Wednesday, September 30, 2026, Tentative'
      ]
    });
    const res = await run(cal, '2026-09-28', '2026-10-04');
    assert.equal(res.status, 'ok', JSON.stringify(res.error || res.note).slice(0, 400));
    assert.equal(res.count, 3, `alle 3 Termine (war ${res.count})`);
    assert.equal(res.dataQuality, 'verified');
    assert.equal(res.weekVerified, true);
    const titles = res.events.map((e) => e.title).sort();
    assert.deepEqual(titles, ['Review', 'TeamSync', 'Urlaub'].sort(), 'Titel aller Events');
    const urlaub = res.events.find((e) => e.title === 'Urlaub');
    assert.equal(urlaub.allDay, true);
    assert.equal(urlaub.status, 'free');
    const teamSync = res.events.find((e) => e.title === 'TeamSync');
    assert.equal(teamSync.start, '10:00');
    assert.equal(teamSync.end, '11:00');
    assert.equal(teamSync.status, 'busy');
    // READ-ONLY: keine Schreibaktion ausgefuehrt (nur ggf. Navigation).
    assert.ok(cal.state.clicks.length === 0, 'Zielwoche stand bereits im Header — 0 Klicks');
  });

  test('navigiert die Woche zielgerichtet (1 Klick) und listet die Zielwoche', async () => {
    const eventsByWeek = (mondayIso) => (mondayIso === '2026-09-28'
      ? ['TeamSync, 10:00 to 11:00, Monday, September 28, 2026, Busy']
      : ['Alt, 11:00 to 12:00, Monday, September 21, 2026, Busy']);
    const cal = buildCalendar({ displayedMonday: '2026-09-21', eventsByWeek });
    const res = await run(cal, '2026-09-28', '2026-10-04');
    assert.equal(res.renderedWeek.startDate, '2026-09-28', 'Zielwoche angezeigt');
    assert.equal(res.count, 1, 'Events der Zielwoche');
    assert.equal(res.events[0].title, 'TeamSync');
    assert.deepEqual(cal.state.clicks, ['next']);
    assert.equal(res.weekVerified, true);
  });

  test('filtert Pseudo-Slots (Arbeitsort) und Duplikate heraus', async () => {
    const cal = buildCalendar({
      displayedMonday: '2026-09-28',
      events: [
        'TeamSync, 10:00 to 11:00, Monday, September 28, 2026, Busy',
        'TeamSync, 10:00 to 11:00, Monday, September 28, 2026, Busy',
        'Working location, all day event, Tuesday, September 29, 2026'
      ]
    });
    const res = await run(cal, '2026-09-28', '2026-10-04');
    assert.equal(res.status, 'ok');
    assert.equal(res.count, 1, 'Duplikat + Pseudo-Slot entfernt, 1 echter Termin bleibt');
    assert.equal(res.events[0].title, 'TeamSync');
  });

  test('ehrlich: keine Events im Zeitraum wird nicht als Fehler gemeldet', async () => {
    const cal = buildCalendar({
      displayedMonday: '2026-09-28',
      events: ['TeamSync, 10:00 to 11:00, Monday, September 28, 2026, Busy']
    });
    // Zeitraum NACH der gerenderten Woche -> kein Event im Range.
    const res = await run(cal, '2026-10-12', '2026-10-18');
    assert.equal(res.status, 'ok');
    assert.equal(res.count, 0);
    assert.match(res.note || '', /keine im angefragten Zeitraum/, 'ehrlicher Hinweis');
  });

  test('fehlende Parameter werfen einen Fehler', async () => {
    activeCalendar = buildCalendar({ displayedMonday: '2026-09-28' });
    const client = new OwaClient();
    await assert.rejects(() => client.listCalendarEvents({ start: '2026-09-28' }), /start.*end|erforderlich/i);
  });
});

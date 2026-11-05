import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { waitForEventSetStable, EVENT_SET_STABLE_ROUNDS, EVENT_SET_STABLE_POLL_MS } from '../src/owa.js';
import { OwaClient } from '../src/owa.js';
import { browserManager } from '../src/browser.js';
import { buildCalendar } from './calendar-list.test.mjs';

// ===========================================================================
// Event-Settlement (Runde 3, 05.11.2026 — Live-Befund KW47)
//
// Befund: get_freebusy lieferte 5 busySlots (u.a. Kellinghusen-Blocker
// Di 17.11. 10:00-10:30), list_calendar_events nur 4 Events — trotz
// verifiziertem Wochen-Header. Root-Cause: waitForWeekViewReady hat sein
// Ruhe-Fenster nur im notReady-Pfad; bei direkter URL-Navigation auf die
// Zielwoche kehrt es sofort mit targetWeek zurueck. Der finale Collect lief
// mitten in OWA's lazy Hydration -> partielles Event-Set mit verified Header.
//
// Fix: waitForEventSetStable — die Event-Signatur (sortierte Labels + Knoten-
// zahl) muss ueber stableRounds aufeinanderfolgende Polls identisch bleiben,
// bevor der finale Read laeuft. Spinner setzen das Ruhe-Fenster zurueck.
// ===========================================================================

const mkClock = () => {
  let t = 1_000_000;
  return {
    now: () => t,
    tick: (ms) => { t += ms; }
  };
};

const mkRead = (sets) => {
  // sets: Array von { labels, eventElementsFound, loading } — der n-te Call
  // liefert das min(n, len-1)-te Set.
  let calls = 0;
  return () => {
    const s = sets[Math.min(calls, sets.length - 1)];
    calls++;
    return s;
  };
};

describe('waitForEventSetStable: Signatur-Stabilitaet (Unit, DI)', () => {
  test('stabil, sobald die Signatur ueber stableRounds Polls identisch bleibt', async () => {
    const clock = mkClock();
    const read = mkRead([
      { labels: ['A, 10:00 to 11:00, Monday, November 16, 2026'], eventElementsFound: 1, loading: false },
      { labels: ['A, 10:00 to 11:00, Monday, November 16, 2026', 'B, 14:00 to 15:00, Monday, November 16, 2026'], eventElementsFound: 2, loading: false },
      { labels: ['A, 10:00 to 11:00, Monday, November 16, 2026', 'B, 14:00 to 15:00, Monday, November 16, 2026'], eventElementsFound: 2, loading: false }
    ]);
    const res = await waitForEventSetStable({
      readEvents: read, budgetMs: 10_000, pollMs: 500,
      stableRounds: EVENT_SET_STABLE_ROUNDS,
      sleep: async (ms) => clock.tick(ms), now: clock.now
    });
    assert.equal(res.stable, true);
    assert.equal(res.polls, 3, '1 Wechsel + 2 identische Polls');
    assert.equal(res.labelsCount, 2);
    assert.equal(res.exhausted, false);
  });

  test('wachsendes Event-Set setzt das Ruhe-Fenster zurueck (Lazy-Hydration)', async () => {
    const clock = mkClock();
    const read = mkRead([
      { labels: ['A'], eventElementsFound: 1, loading: false },
      { labels: ['A'], eventElementsFound: 1, loading: false },
      { labels: ['A', 'B'], eventElementsFound: 2, loading: false },
      { labels: ['A', 'B'], eventElementsFound: 2, loading: false },
      { labels: ['A', 'B'], eventElementsFound: 2, loading: false }
    ]);
    const res = await waitForEventSetStable({
      readEvents: read, budgetMs: 10_000, pollMs: 500,
      stableRounds: 3,
      sleep: async (ms) => clock.tick(ms), now: clock.now
    });
    assert.equal(res.stable, true);
    assert.equal(res.polls, 5, 'die 2x-identischen "A"-Polls zaehlen NICHT weiter — ' +
      'ohne Reset haette die Signatur "A" bei Poll 3 die 3. Runde erreicht');
    assert.equal(res.labelsCount, 2);
  });

  test('Lade-Indikator (Spinner) macht das Ruhe-Fenster ungueltig', async () => {
    const clock = mkClock();
    const read = mkRead([
      { labels: ['A'], eventElementsFound: 1, loading: false },
      { labels: ['A'], eventElementsFound: 1, loading: true },
      { labels: ['A'], eventElementsFound: 1, loading: false },
      { labels: ['A'], eventElementsFound: 1, loading: false }
    ]);
    const res = await waitForEventSetStable({
      readEvents: read, budgetMs: 10_000, pollMs: 500,
      stableRounds: EVENT_SET_STABLE_ROUNDS,
      sleep: async (ms) => clock.tick(ms), now: clock.now
    });
    assert.equal(res.stable, true);
    assert.equal(res.polls, 4, 'Spinner-Poll reset — Stabilitaet erst nach 2 weiteren identischen Polls');
  });

  test('Budget-Erschoepfung: ehrlich exhausted statt stable zu behaupten', async () => {
    const clock = mkClock();
    let calls = 0;
    const res = await waitForEventSetStable({
      readEvents: () => { calls++; return { labels: [`E${calls}`], eventElementsFound: 1, loading: false }; },
      budgetMs: 2_000, pollMs: 500,
      stableRounds: EVENT_SET_STABLE_ROUNDS,
      sleep: async (ms) => clock.tick(ms), now: clock.now
    });
    assert.equal(res.stable, false);
    assert.equal(res.exhausted, true);
    assert.ok(calls >= 2, 'mehrfach gepollt');
  });

  test('ohne readEvents/Budget wird ehrlich skipped gemeldet', async () => {
    const noRead = await waitForEventSetStable({ budgetMs: 1_000 });
    assert.equal(noRead.stable, false);
    assert.equal(noRead.skipped, true);
    const noBudget = await waitForEventSetStable({ readEvents: () => ({ labels: [] }), budgetMs: 0 });
    assert.equal(noBudget.stable, false);
    assert.equal(noBudget.skipped, true);
  });
});

// ---------------------------------------------------------------------------
// Echter Pfad: OwaClient.listCalendarEvents gegen ein lazily hydratisierendes
// Fake-DOM. Ohne Settlement las der finale Read das PARTIELLE Set (Live-Befund
// KW47: 4 von 7). Mit Settlement muss die VOLLSTAENDIGE Menge gelesen werden.
// ---------------------------------------------------------------------------
describe('listCalendarEvents: Event-Settlement auf dem Produktionspfad', () => {
  const realEnsure = browserManager.ensureContext.bind(browserManager);
  const prevDoc = globalThis.document;
  const prevWin = globalThis.window;
  let activeCalendar = null;

  test.before(() => {
    browserManager.ensureContext = async () => ({ page: activeCalendar.page, context: {} });
    browserManager.resetForTest();
    globalThis.window = { getComputedStyle: () => ({ display: 'block', visibility: 'visible', position: 'static' }) };
  });
  test.after(() => {
    browserManager.ensureContext = realEnsure;
    globalThis.document = prevDoc;
    globalThis.window = prevWin;
  });

  test('lazy Hydration: partielles Grid (2) -> vollstaendiges Grid (4) wird komplett gelesen', async () => {
    const partial = [
      'TeamSync, 10:00 to 11:00, Monday, November 16, 2026, Busy',
      'Blocker, 10:00 to 10:30, Tuesday, November 17, 2026, Busy'
    ];
    const full = [
      ...partial,
      'Review, 14:00 to 15:00, Wednesday, November 18, 2026, Busy',
      'JF, 09:00 to 09:30, Friday, November 20, 2026, Busy'
    ];
    activeCalendar = buildCalendar({
      displayedMonday: '2026-11-16',
      events: partial,
      hydrateTo: { afterMs: 300, labels: full }
    });
    const client = new OwaClient();
    const res = await client.listCalendarEvents({ start: '2026-11-16', end: '2026-11-22' });
    assert.equal(res.status, 'ok', JSON.stringify(res.error || res.note).slice(0, 400));
    assert.equal(res.count, 4, `volles Event-Set gelesen (war ${res.count})`);
    assert.equal(res.weekVerified, true);
    assert.equal(res.eventSettlement.stable, true, 'Settlement stabil');
    assert.ok(!/nicht stabil/.test(res.note), 'keine Instabilitaets-Warnung');
  });

  test('sofort vollstaendige Hydration (afterMs: 0): kein Settlement-Delay, volles Set', async () => {
    activeCalendar = buildCalendar({
      displayedMonday: '2026-11-16',
      events: ['TeamSync, 10:00 to 11:00, Monday, November 16, 2026, Busy'],
      hydrateTo: {
        afterMs: 0,
        labels: ['TeamSync, 10:00 to 11:00, Monday, November 16, 2026, Busy']
      }
    });
    const client = new OwaClient();
    const res = await client.listCalendarEvents({ start: '2026-11-16', end: '2026-11-22' });
    assert.equal(res.status, 'ok');
    assert.equal(res.count, 1, 'statisches Set bleibt stabil — keine Verzoegerung, keine Warnung');
    assert.equal(res.eventSettlement.stable, true);
  });
});

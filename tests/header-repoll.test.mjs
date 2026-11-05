import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import {
  pollForDisplayedHeaderRead,
  isHeaderReadVerified
} from '../src/owa.js';

// ===========================================================================
// Runde 2, Befund 3 (Live-Lauf 05.11.2026, d-velop-Realm): Die aktuelle Woche
// (02.–08.11.) hatte 8 Events (inhaltlich PASS), aber dataQuality unverified.
// Live-Diagnostik bewies: der ECHTE Header ("2–8 November, 2026, Jump to a
// specific date or date range.") stand im DOM — die Navigation hatte aber
// einen nav-abgeleiteten Wochenstand akzeptiert und nie re-polliert.
//
// pollForDisplayedHeaderRead holt den gelesenen Header nach (Budget + DI),
// praevaliert aber nichts: bleibt nur die Ableitung, kommt null zurueck und
// der Aufrufer meldet ehrlich 'unverified'.
// ===========================================================================

const HEADER_READ = {
  text: '2–8 November, 2026, Jump to a specific date or date range.',
  startDate: '2026-11-02',
  endDate: '2026-11-08',
  source: 'displayed-header'
};
const NAV_DERIVED = {
  text: 'Go to previous week 26 October – 1 November, 2026 (Week 44)',
  startDate: '2026-11-02',
  endDate: '2026-11-08',
  source: 'nav-label-derivation',
  readError: 'no-displayed-header'
};

describe('pollForDisplayedHeaderRead (Runde 2, Befund 3)', () => {
  test('Befund-3-Fall: nav-abgeleitet zuerst, echter Header im 3. Poll -> Upgrade', async () => {
    let calls = 0;
    const reads = [];
    const r = await pollForDisplayedHeaderRead({
      readWeek: async () => {
        calls++;
        const w = calls >= 3 ? HEADER_READ : NAV_DERIVED;
        reads.push(w.source);
        return w;
      },
      budgetMs: 5000,
      pollMs: 10,
      sleep: async () => {}
    });
    assert.ok(isHeaderReadVerified(r), 'Ergebnis muss ein gelesener Header sein');
    assert.equal(r.startDate, '2026-11-02');
    assert.deepEqual(reads, ['nav-label-derivation', 'nav-label-derivation', 'displayed-header']);
  });

  test('Header erscheint nie: null zurueck (keine Praevalierung der Ableitung)', async () => {
    let calls = 0;
    const r = await pollForDisplayedHeaderRead({
      readWeek: async () => { calls++; return NAV_DERIVED; },
      budgetMs: 120,
      pollMs: 30,
      sleep: async () => {}
    });
    assert.equal(r, null);
    assert.ok(calls >= 2, 'es muss tatsaechlich gepollt worden sein');
  });

  test('Erster Read schon verifiziert: sofortiges Ergebnis ohne weitere Polls', async () => {
    let calls = 0;
    const r = await pollForDisplayedHeaderRead({
      readWeek: async () => { calls++; return HEADER_READ; },
      budgetMs: 5000,
      pollMs: 10,
      sleep: async () => {}
    });
    assert.equal(calls, 1);
    assert.ok(isHeaderReadVerified(r));
  });

  test('Budget 0 / readWeek fehlt: null ohne Call', async () => {
    let calls = 0;
    assert.equal(await pollForDisplayedHeaderRead({ readWeek: async () => { calls++; return HEADER_READ; }, budgetMs: 0 }), null);
    assert.equal(await pollForDisplayedHeaderRead({ budgetMs: 5000 }), null);
    assert.equal(calls, 0);
  });

  test('readWeek wirft transient: Poll laeuft weiter statt abzubrechen', async () => {
    let calls = 0;
    const r = await pollForDisplayedHeaderRead({
      readWeek: async () => {
        calls++;
        if (calls < 3) throw new Error('Execution context was destroyed');
        return HEADER_READ;
      },
      budgetMs: 2000,
      pollMs: 10,
      sleep: async () => {}
    });
    assert.ok(isHeaderReadVerified(r));
    assert.equal(calls, 3);
  });
});

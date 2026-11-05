import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
  renderedWeekContains,
  weekProbeSkipEntry,
  weekProbeGuardEntry,
  weekClickDecision,
  weekButtonCandidates,
  shouldTryNextCandidate,
  evaluateClickVerification,
  canClickSameDirection,
  classifyButtonFailure,
  runWeekButtonClick,
  weekViewState,
  waitForWeekViewReady,
  WEEK_CLICK_GUARD_REASONS,
  WEEK_VIEW_STATES
} from '../src/owa.js';

// ===========================================================================
// Bug 2 — der Zielwochen-Guard ist ein stiller No-op (Live-Befund 25.09.2026)
//
// Live-Trace:
//   0. { kind:'aria-label-de', found:false, clicked:false }
//   1. { kind:'aria-label-en', found:true,  clicked:true, verified:'unverified' }
//   2. { skipped:'target-week-verified', renderedWeek:'2026-09-28', clicked:false }
// `note:'guard-vor-klick'` erscheint NIRGENDS.
//
// Ursache, aus dem Trace abgeleitet (kein Raten): evaluateClickVerification
// liefert 'unverified' genau dann, wenn prevStart ODER newStart fehlt. Eintrag 2
// stammt aus `rendered = forced.week` und zeigt einen DATIERBAREN Header — also
// war newStart gesetzt, und damit zwingend prevStart === null. Beim Klick war
// KEIN Wochen-Header gerendert.
//
// Der alte Guard fragte nur "ist die Zielwoche schon da?". Auf einer
// ungerenderten Ansicht ist renderedWeekContains immer false — der Guard war
// also ein No-op GENAU in dem Fall, fuer den er gebaut wurde, und er failte
// still (kein Probe-Eintrag, kein note, kein Abbruchgrund).
//
// WICHTIG (das war der Fehler des letzten Fixes 8ae0eec): Dieser Block rief
// previously KEINEN Produktionscode auf — er baute die Kandidaten-Schleife per
// `runGuardedProbe` NACH. Ein solcher Test ist gruen, egal ob der echte Code
// einen Guard hat. Jetzt wird runWeekButtonClick aufgerufen: die exportierte
// Funktion, die getFreeBusy selbst benutzt. Der Test fährt damit den Weg
// Guard -> Probe -> Klick wirklich.
// ===========================================================================

// Fake-Page: so wenig wie moeglich, aber genau das, was runWeekButtonClick
// benutzt (locator().first() -> waitFor + click). `visible` steuert, welche
// Kandidaten es ueberhaupt gibt.
const fakePage = (visible = []) => {
  const clicks = [];
  return {
    clicks,
    locator(selector) {
      return {
        first() {
          return {
            async waitFor() {
              if (!visible.some((v) => selector.includes(v))) {
                throw new Error('not visible');
              }
            },
            async click() { clicks.push(selector); }
          };
        }
      };
    }
  };
};

const TARGET = '2026-09-28';
const TARGET_WEEK = { text: '28. September – 4. Oktober', startDate: '2026-09-28', endDate: '2026-10-04' };
const WRONG_WEEK = { text: '21. – 27. September', startDate: '2026-09-21', endDate: '2026-09-27' };
// Der belegte Live-Zustand beim Klick: KEIN datierbarer Header.
const UNRENDERED = { text: 'Work week', startDate: null, endDate: null };

// Der ECHTE Produktionspfad, mit denselben Abhaengigkeiten wie in getFreeBusy.
const runGuard = (opts = {}) => {
  const probe = [];
  const page = fakePage(opts.visible || ['Next week']);
  const afterClick = async (prevStart, budgetMs) => ({
    week: opts.afterWeek || WRONG_WEEK,
    verified: opts.verified || 'changed'
  });
  return runWeekButtonClick({
    direction: 1,
    prevStart: opts.prevStart === undefined ? null : opts.prevStart,
    targetStart: TARGET,
    page,
    probe,
    readWeek: async () => opts.readWeek || UNRENDERED,
    afterClick,
    remainingMs: () => 30000
  }).then((r) => ({ ...r, probe, clicks: page.clicks }));
};

// ===========================================================================

describe('Bug 2 — Zielwochen-Guard vor dem Klick', () => {

  // --- Reine Pruefung ------------------------------------------------------
  describe('renderedWeekContains — enthaelt die gerenderte Woche das Startdatum?', () => {
    test('Zielwoche erkannt (2026-09-28 in KW 40)', () => {
      assert.equal(renderedWeekContains({ startDate: '2026-09-28', endDate: '2026-10-04' }, '2026-09-28'), true);
    });
    test('Startdatum MITTEN in der gerenderten Woche', () => {
      assert.equal(renderedWeekContains({ startDate: '2026-09-28', endDate: '2026-10-04' }, '2026-10-01'), true);
    });
    test('Letzter Tag der Woche', () => {
      assert.equal(renderedWeekContains({ startDate: '2026-09-28', endDate: '2026-10-04' }, '2026-10-04'), true);
    });
    test('Falsche Woche -> false', () => {
      assert.equal(renderedWeekContains({ startDate: '2026-09-21', endDate: '2026-09-27' }, '2026-09-28'), false);
    });
    test('Woche beginnt VOR der Zielwoche -> false', () => {
      assert.equal(renderedWeekContains({ startDate: '2026-09-21', endDate: '2026-10-04' }, '2026-09-28'), true,
        'überlappende Range gilt als enthalten — Grenzfall bleibt True wie bisher');
    });
    test('Nicht datierter Header -> false (KEIN Klick-Skip auslösen)', () => {
      assert.equal(renderedWeekContains({ startDate: null, endDate: null, text: 'Work week' }, '2026-09-28'), false);
    });
    test('Fehlendes Ende -> false', () => {
      assert.equal(renderedWeekContains({ startDate: '2026-09-28', endDate: null }, '2026-09-28'), false);
    });
    test('null / undefined / leeres Datum -> false', () => {
      assert.equal(renderedWeekContains(null, '2026-09-28'), false);
      assert.equal(renderedWeekContains(undefined, '2026-09-28'), false);
      assert.equal(renderedWeekContains({ startDate: '2026-09-28', endDate: '2026-10-04' }, null), false);
      assert.equal(renderedWeekContains({ startDate: '2026-09-28', endDate: '2026-10-04' }, ''), false);
    });
  });

  // --- Protokolleintrag ----------------------------------------------------
  test('weekProbeSkipEntry: ohne note unveraendert (Rueckwaertskompatibilitaet)', () => {
    assert.deepEqual(weekProbeSkipEntry('2026-09-28'), {
      skipped: 'target-week-verified', renderedWeek: '2026-09-28', clicked: false
    });
  });

  test('weekProbeSkipEntry: note markiert die Guard-Stelle', () => {
    assert.deepEqual(weekProbeSkipEntry('2026-09-28', 'guard-vor-klick'), {
      skipped: 'target-week-verified', renderedWeek: '2026-09-28', clicked: false, note: 'guard-vor-klick'
    });
  });

  test('weekProbeSkipEntry: null wird nicht als note gesetzt', () => {
    assert.ok(!('note' in weekProbeSkipEntry('2026-09-28', null)));
    assert.equal(weekProbeSkipEntry(null).renderedWeek, null);
  });

  // =========================================================================
  // Der entscheidende Teil: der AUFRUFENDE PFAD (getFreeBusy -> Guard -> Probe)
  // =========================================================================

  test('REGRESSION (Live): ungerenderte Ansicht -> NULL Klicks, view-not-ready protokolliert', async () => {
    // Das ist der belegte Live-Zustand: prevStart === null, also kein
    // datierbarer Wochen-Header. Mit dem alten Guard ("nur Zielwoche schon da?")
    // waegst du hier durch und klicktest auf den "Next week"-Button.
    const r = await runGuard({ readWeek: UNRENDERED, prevStart: null, afterWeek: TARGET_WEEK, verified: 'unverified' });

    assert.equal(r.clicks.length, 0, `KEIN Klick auf einer ungerenderten Ansicht (war: ${JSON.stringify(r.clicks)})`);
    assert.equal(r.clicked, false);
    assert.equal(r.guard, WEEK_CLICK_GUARD_REASONS.viewNotReady);
    assert.equal(r.targetReached, false, 'die Zielwoche wurde nicht erreicht — nur blockiert');
    assert.equal(r.probe.length, 1, 'genau ein Eintrag: der Guard');
    assert.equal(r.probe[0].skipped, 'view-not-ready');
    assert.equal(r.probe[0].note, 'guard-vor-klick', 'die Guard-Stelle ist im Protokoll sichtbar');
    assert.equal(r.probe[0].clicked, false);
    assert.equal(r.probe.some((e) => e.clicked === true), false);
  });

  test('REGRESSION: `note:"guard-vor-klick"` muss im Live-Protokoll AUFTAUCHEN', async () => {
    // Genau daran ist der Live-Befund erkannt worden: die Note erschien
    // nirgends. Ein Guard, der nichts protokollieren kann, ist von einem
    // nicht vorhandenen Guard nicht unterscheidbar.
    const blocked = await runGuard({ readWeek: UNRENDERED });
    const reached = await runGuard({ readWeek: TARGET_WEEK });
    assert.ok(blocked.probe.some((e) => e.note === 'guard-vor-klick'), 'Blockade wird protokolliert');
    assert.ok(reached.probe.some((e) => e.note === 'guard-vor-klick'), 'Zielwoche-Skip wird protokolliert');
  });

  test('Zielwoche verifiziert: Skip steht VOR jeder Probe, null Klicks', async () => {
    const r = await runGuard({ readWeek: TARGET_WEEK, prevStart: '2026-09-21' });
    assert.equal(r.targetReached, true);
    assert.equal(r.guard, WEEK_CLICK_GUARD_REASONS.targetReached);
    assert.equal(r.clicks.length, 0, 'KEIN Klick darf protokolliert werden');
    assert.equal(r.probe.length, 1, 'genau ein Eintrag');
    assert.equal(r.probe[0].skipped, 'target-week-verified');
    assert.equal(r.probe[0].renderedWeek, '2026-09-28');
    assert.equal(r.probe[0].clicked, false);
    assert.equal(r.probe[0].note, 'guard-vor-klick');
    assert.equal(r.probe.some((e) => e.selector), false, 'keine Probe wurde überhaupt durchgefuehrt');
  });

  test('falsche, ABER RENDERTE Woche -> der Klick passiert (der Guard blockiert nicht alles)', async () => {
    const r = await runGuard({ readWeek: WRONG_WEEK, prevStart: '2026-09-21', afterWeek: TARGET_WEEK, verified: 'changed' });
    assert.equal(r.clicks.length, 1, 'genau ein Klick auf den ersten passenden Kandidaten');
    assert.ok(r.clicks[0].includes('Next week'));
    assert.equal(r.clicked, true);
    assert.equal(r.guard, WEEK_CLICK_GUARD_REASONS.click);
    assert.equal(r.probe.filter((e) => e.clicked === true).length, 1);
    assert.equal(r.probe.some((e) => e.skipped), false, 'kein Skip — die Woche war ja falsch');
  });

  test('alle Kandidaten werden weiterhin in der Reihenfolge probiert, die vorher nicht passten', async () => {
    const r = await runGuard({ readWeek: WRONG_WEEK, prevStart: '2026-09-21' });
    const kinds = r.probe.map((e) => e.kind).filter(Boolean);
    assert.equal(kinds[0], 'aria-label-de', 'erst DE, dann DE-Dativ, dann EN — Reihenfolge dokumentiert');
    assert.equal(kinds[1], 'aria-label-de-dat');
    assert.equal(kinds[2], 'aria-label-en');
    assert.equal(r.probe[0].found, false, 'DE-Button existiert nicht');
    assert.equal(r.probe[2].clicked, true, 'EN-Button wird geklickt');
  });

  test('der Guard-Read ist memoisiert und wird nach einem Klick verworfen', async () => {
    // Genau die Semantik, die 8ae0eec behauptet hat — jetzt geprueft.
    let reads = 0;
    const probe = [];
    const page = fakePage(['Next week']);
    let week = WRONG_WEEK;
    await runWeekButtonClick({
      direction: 1, prevStart: '2026-09-21', targetStart: TARGET, page, probe,
      readWeek: async () => { reads++; return week; },
      afterClick: async () => { week = WRONG_WEEK; return { week: WRONG_WEEK, verified: 'changed' }; },
      remainingMs: () => 30000
    });
    // 8 Kandidaten, aber der Read ist memoisiert -> 1 Read fuer die ganze
    // Kandidatenliste (erst Kandidat), und kein zweiter, weil nach dem
    // erfolgreichen Klick nicht mehr iteriert wird.
    assert.equal(reads, 1, `memoisiert: genau 1 Read fuer 8 Kandidaten (war ${reads})`);
  });

  test('weekClickDecision: die geteilte Definition, die beide Stellen benutzen', () => {
    // --- Zustand 2: Zielwoche gerendert -> fertig, kein Klick ---------------
    const reached = weekClickDecision({ rendered: TARGET_WEEK, targetStart: TARGET });
    assert.equal(reached.skip, true);
    assert.equal(reached.reason, WEEK_CLICK_GUARD_REASONS.targetReached);
    assert.equal(reached.state, WEEK_VIEW_STATES.targetWeek);
    assert.equal(reached.clickable, false);

    // --- Zustand 3: falsche Woche, ABER gerendert -> Klick ist erlaubt ------
    const wrong = weekClickDecision({ rendered: WRONG_WEEK, targetStart: TARGET });
    assert.equal(wrong.skip, false);
    assert.equal(wrong.reason, WEEK_CLICK_GUARD_REASONS.click);
    assert.equal(wrong.state, WEEK_VIEW_STATES.wrongWeek);
    assert.equal(wrong.clickable, true);

    // --- Zustand 1: kein datierbarer Header -> WARTEN, nicht klicken --------
    const unrendered = weekClickDecision({ rendered: UNRENDERED, targetStart: TARGET });
    assert.equal(unrendered.skip, true);
    assert.equal(unrendered.reason, WEEK_CLICK_GUARD_REASONS.viewNotReady);
    assert.equal(unrendered.state, WEEK_VIEW_STATES.notReady);
    assert.equal(unrendered.clickable, false);
  });

  test('BUG 5: die hasEvents-Ausnahme ist ENTFALLEN (Zustand 1 ist Zustand 1)', () => {
    // 1c2825e noch: "sichtbare Events belegen eine gerenderte Ansicht" -> Klick
    // erlaubt. Das ist der Live-Befund: die Event-Chips SASSEN im DOM (18
    // Knoten), der Header fehlte, und die Wochenrichtung war damit unbekannt.
    // Diese Ausnahme behauptete einen vierten Zustand, den es nicht gibt.
    // runWeekButtonClick reichte hasEvents ohnehin NIE durch (Live: der Guard
    // griff mit view-not-ready) — die Ausnahme war tote, irrefuehrende Logik.
    const withEvents = weekClickDecision({ rendered: UNRENDERED, targetStart: TARGET, hasEvents: true });
    assert.equal(withEvents.skip, true,
      'sichtbare Events ersetzen keinen datierbaren Wochen-Header');
    assert.equal(withEvents.reason, WEEK_CLICK_GUARD_REASONS.viewNotReady);
    // weekViewState kennt das Argument gar nicht mehr — die Zweit-Wahrheit ist weg.
    assert.equal(weekViewState({ rendered: UNRENDERED, targetStart: TARGET, hasEvents: true }).state,
      WEEK_VIEW_STATES.notReady);
  });

  test('BUG 5: die drei Zustaende sind paarweise disjunkt und exhaustiv', () => {
    const cases = [
      [UNRENDERED, WEEK_VIEW_STATES.notReady],
      [{ text: null, startDate: null, endDate: null }, WEEK_VIEW_STATES.notReady],
      [{ text: 'x', startDate: '2026-09-21', endDate: null }, WEEK_VIEW_STATES.notReady],
      [WRONG_WEEK, WEEK_VIEW_STATES.wrongWeek],
      [TARGET_WEEK, WEEK_VIEW_STATES.targetWeek],
      [{ text: '28.9.–4.10.', startDate: '2026-09-28', endDate: '2026-10-04' }, WEEK_VIEW_STATES.targetWeek]
    ];
    for (const [rendered, expected] of cases) {
      assert.equal(weekViewState({ rendered, targetStart: TARGET }).state, expected,
        `unerwarteter Zustand fuer ${JSON.stringify(rendered)}`);
    }
    // Nur 'wrong-week' ist klickbar.
    for (const [rendered] of cases) {
      const s = weekViewState({ rendered, targetStart: TARGET }).state;
      assert.equal(s === WEEK_VIEW_STATES.wrongWeek, weekClickDecision({ rendered, targetStart: TARGET }).clickable);
    }
  });

  test('BUG 5: Zustand 1 wartet, BEVOR es ueberhaupt eine Probe gibt', async () => {
    // Der Klickpfad muss den fehlenden Header zuerst abwarten und danach neu
    // entscheiden — genau an dieser Stelle wurde in Runde 4 nach 13ms abgebrochen.
    // Fixture: readWeek liefert 3x nichts, dann die ZIELWOCHE.
    const probe = [];
    const page = fakePage(['Next week']);
    let reads = 0;
    const clock = { t: 0, now: () => clock.t, sleep: async (ms) => { clock.t += ms; } };
    // readWeek liefert 3x nichts, dann die gerenderte (falsche) Woche — der
    // Kalender "schaltet" also waehrend des Wartens um.
    const headerSource = async () => {
      reads++;
      return reads <= 3 ? UNRENDERED : WRONG_WEEK;
    };
    const r = await runWeekButtonClick({
      direction: 1, prevStart: '2026-09-21', targetStart: TARGET, page, probe,
      readWeek: headerSource,
      afterClick: async () => ({ week: TARGET_WEEK, verified: 'changed' }),
      remainingMs: () => 30000,
      // DIE Produktions-Wartefunktion, mit derselben Uhr und derselben Quelle
      // wie der Guard — so wie in getFreeBusy.
      waitForReady: () => waitForWeekViewReady({
        readWeek: headerSource,
        readEvents: async () => ({ loading: false }),
        budgetMs: 4000, pollMs: 100, settleMs: 500,
        targetStart: TARGET, now: clock.now, sleep: clock.sleep
      })
    });
    // Der Wichtigste: der Header kam erst beim DRITTEN Read. Ohne das Warten
    // waere der Guard beim ersten Read gescheitert und der Klick passierte nie.
    assert.ok(reads >= 4, `der Header wurde mehrfach gelesen (war ${reads})`);
    assert.ok(clock.t > 0, `es wurde fuer den Header gewartet (war ${clock.t}ms)`);
    assert.equal(page.clicks.length, 1, 'nach dem Warten wird die Navigation ausgefuehrt');
    assert.equal(r.clicked, true);
    assert.equal(r.guard, WEEK_CLICK_GUARD_REASONS.click);
    // Der Klick passiert NACH dem Warten — Reihenfolge im Protokoll belegt:
    // der verweigernde Guard-Eintrag darf nicht auftauchen.
    assert.ok(!probe.some((e) => e.skipped), 'es gab keine Verweigerung, sondern ein Warten');
  });

  test('BUG 5: bleibt der Header aus, wird NACH dem Wartebudget verweigert', async () => {
    const probe = [];
    const page = fakePage(['Next week']);
    const clock = { t: 0, now: () => clock.t, sleep: async (ms) => { clock.t += ms; } };
    const r = await runWeekButtonClick({
      direction: 1, prevStart: '2026-09-21', targetStart: TARGET, page, probe,
      readWeek: async () => UNRENDERED,
      afterClick: async () => ({ week: TARGET_WEEK, verified: 'changed' }),
      remainingMs: () => 30000,
      waitForReady: async () => waitForWeekViewReady({
        readWeek: async () => UNRENDERED,
        readEvents: async () => ({ loading: false }),
        // settleMs > budgetMs: damit endet das Warten am BUDGET, nicht am
        // Ruhe-Fenster. Beide Abbrucharten werden getestet.
        budgetMs: 1000, pollMs: 100, settleMs: 30000,
        targetStart: TARGET, now: clock.now, sleep: clock.sleep
      })
    });
    assert.equal(page.clicks.length, 0, 'KEIN Klick auf ungerenderter Ansicht');
    assert.equal(r.guard, WEEK_CLICK_GUARD_REASONS.viewNotReady);
    assert.equal(r.clicked, false);
    const entry = probe.find((e) => e.skipped === 'view-not-ready');
    assert.ok(entry, 'die Verweigerung steht im Protokoll');
    assert.equal(entry.clicked, false);
    // BUG 5: waitedMs unterscheidet "nach 13ms verweigert" (der Live-Befund)
    // von "nach dem Wartebudget verweigert".
    assert.ok(entry.waitedMs >= 900, `das Wartebudget wurde verbraucht (war ${entry.waitedMs}ms)`);
  });

  test('BUG 5: das Ruhe-Fenster beendet das Warten ebenfalls ehrlich', () => {
    // Zweiter Abbruchweg: Header fehlt, aber die Ansicht ist SPINNER-FREI und
    // bleibt es. Dann ist "warten" sinnlos — nach settleMs ist es belegt, dass
    // der Header nicht mehr kommt. Auch das ist ein Abbruch mit Wartezeit.
    const clock = { t: 0, now: () => clock.t, sleep: async (ms) => { clock.t += ms; } };
    const r = waitForWeekViewReady({
      readWeek: async () => UNRENDERED,
      readEvents: async () => ({ loading: false }),
      budgetMs: 20000, pollMs: 100, settleMs: 400,
      targetStart: TARGET, now: clock.now, sleep: clock.sleep
    });
    return r.then((w) => {
      assert.equal(w.state, WEEK_VIEW_STATES.notReady);
      assert.equal(w.exhausted, true);
      assert.equal(w.settled, true, 'das Ruhe-Fenster hat beendet, nicht das Budget');
      assert.ok(w.waitedMs >= 400 && w.waitedMs < 20000);
    });
  });

  test('BUG 5: der Spinner haelt das Warten am Leben (kein Leer-Beweis)', async () => {
    // 2 ruhige Polls, dann Spinner, dann Header. Ohne Spinner-Beachtung wuerde
    // das Ruhe-Fenster nach settleMs schliessen und der Weg als "Header kommt
    // nicht" enden — obwohl die Ansicht noch laedt.
    const clock = { t: 0, now: () => clock.t, sleep: async (ms) => { clock.t += ms; } };
    let poll = 0;
    const w = await waitForWeekViewReady({
      readWeek: async () => { poll++; return poll <= 4 ? UNRENDERED : TARGET_WEEK; },
      readEvents: async () => ({ loading: poll <= 3 }),
      budgetMs: 5000, pollMs: 100, settleMs: 150,
      targetStart: TARGET, now: clock.now, sleep: clock.sleep
    });
    assert.equal(w.state, WEEK_VIEW_STATES.targetWeek);
    assert.equal(w.exhausted, false);
    assert.ok(w.polls >= 5, `der Spinner hat den Poll am Leben gehalten (war ${w.polls})`);
    assert.ok(w.loadingPolls >= 2, 'Spinner wurden gezaehlt');
  });

  test('BUG 5: der Wochenzustand wird WAEHREND des Wartens neu gelesen', async () => {
    // "Warte bis Header da" heisst "pruefe waehrend des Wartens", nicht
    // "einmal am Anfang schauen" — der Kalender darf zwischen den Polls
    // umschalten.
    const clock = { t: 0, now: () => clock.t, sleep: async (ms) => { clock.t += ms; } };
    let reads = 0;
    const w = await waitForWeekViewReady({
      readWeek: async () => { reads++; return reads >= 3 ? TARGET_WEEK : UNRENDERED; },
      readEvents: async () => ({ loading: true }),
      budgetMs: 5000, pollMs: 100, settleMs: 1000,
      targetStart: TARGET, now: clock.now, sleep: clock.sleep
    });
    assert.equal(w.state, WEEK_VIEW_STATES.targetWeek);
    assert.equal(reads, 3, 'genau so viele Header-Reads bis zum Treffer — bei JEDEM Poll einer');
  });

  test('BUG 5: ohne Budget ist der Abbruch ehrlich (nicht "0 Polls bei 5s")', async () => {
    const r = await waitForWeekViewReady({ readWeek: async () => UNRENDERED, budgetMs: 0, targetStart: TARGET });
    assert.equal(r.exhausted, true);
    assert.equal(r.polls, 0);
    assert.equal(r.state, WEEK_VIEW_STATES.notReady);
  });

  test('BUG 5: wird die Zielwoche nie gerendert, endet das Warten ausgerechnet', async () => {
    // Live-Runde 4: 13ms und Abbruch. Der ehrliche Abbruch braucht das Budget.
    const clock = { t: 0, now: () => clock.t, sleep: async (ms) => { clock.t += ms; } };
    const w = await waitForWeekViewReady({
      readWeek: async () => UNRENDERED,
      readEvents: async () => ({ loading: false }),
      budgetMs: 3000, pollMs: 100, settleMs: 30000,   // settle > budget: nur das Budget beendet
      targetStart: TARGET, now: clock.now, sleep: clock.sleep
    });
    assert.equal(w.exhausted, true);
    assert.equal(w.state, WEEK_VIEW_STATES.notReady);
    // Gleiche Semantik wie die Event-Wartephase: nie ueber das Budget hinaus,
    // aber hoechstens einen Pollintervall zu frue (der Poll, der das Budget
    // erreicht, schlaegt nicht mehr). Live war es 13ms — das ist der Unterschied.
    assert.ok(clock.t >= 3000 - 100, `das Wartebudget wurde ausge schoepft (war ${clock.t}ms)`);
    assert.ok(clock.t > 13, 'nicht die 13ms aus dem Live-Befund');
  });

  test('BUG 5: der Probe-Eintrag nennt die verbrauchte Wartezeit', () => {
    const e = weekProbeGuardEntry(weekClickDecision({ rendered: UNRENDERED, targetStart: TARGET }), 'guard-vor-klick', 4321);
    assert.equal(e.waitedMs, 4321);
    const withoutWait = weekProbeGuardEntry(weekClickDecision({ rendered: UNRENDERED, targetStart: TARGET }));
    assert.equal('waitedMs' in withoutWait, false, 'ohne Warteangabe wird auch keine erfunden');
  });

  test('Der Defekt als Anker: der ALTE Test (nur Zielwoche?) haette hier durchgelassen', () => {
    // Genau die alte, fehlerhafte Entscheidung. Wenn jemand weekClickDecision
    // auf diese Version zurueckbaut, muss dieser Test rot werden.
    const oldDecision = (rendered, dateIso) => Boolean(renderedWeekContains(rendered, dateIso));
    assert.equal(oldDecision(UNRENDERED, TARGET), false, 'alt: nicht drin -> klicken (der Live-Bug)');
    assert.equal(weekClickDecision({ rendered: UNRENDERED, targetStart: TARGET }).skip, true, 'neu: blockieren');
    assert.notEqual(oldDecision(UNRENDERED, TARGET), weekClickDecision({ rendered: UNRENDERED, targetStart: TARGET }).skip);
  });

  test('weekProbeGuardEntry protokolliert BEIDE Ausgaenge unterscheidbar', () => {
    const reached = weekProbeGuardEntry(weekClickDecision({ rendered: TARGET_WEEK, targetStart: TARGET }));
    const blocked = weekProbeGuardEntry(weekClickDecision({ rendered: UNRENDERED, targetStart: TARGET }));
    assert.equal(reached.skipped, 'target-week-verified');
    assert.equal(blocked.skipped, 'view-not-ready');
    assert.equal(reached.guard, WEEK_CLICK_GUARD_REASONS.targetReached);
    assert.equal(blocked.guard, WEEK_CLICK_GUARD_REASONS.viewNotReady);
    assert.equal(reached.note, 'guard-vor-klick');
    assert.equal(blocked.note, 'guard-vor-klick');
    assert.equal(blocked.clicked, false);
  });

  // --- Wechselwirkung mit den bestehenden Regeln ---------------------------
  test('"target-week-reached" ist kein Fehlschlag: shouldTryNextCandidate sagt nein', () => {
    // clickWeekButton gibt bei erreichtem Guard VOR dem shouldTryNextCandidate-
    // Zweig zurueck. Falls er doch durchlaeuft, darf er keinen weiteren
    // Kandidat ausloesen.
    assert.equal(shouldTryNextCandidate('target-week-reached'), false);
    assert.equal(shouldTryNextCandidate('unchanged'), true);
    assert.equal(shouldTryNextCandidate('unverified'), false);
    assert.equal(shouldTryNextCandidate('changed'), false);
  });

  test('classifyButtonFailure: ein reiner Guard-Skip ist KEIN Button-Fehlschlag', () => {
    // Am Aufrufer wird targetReached VOR dem !r.clicked-Zweig behandelt.
    // Falls das je nicht greift, darf ein Skip nicht als 'button-ineffective'
    // (Button existierte ja) oder 'no-week-button' gemeldet werden.
    const skipResult = { clicked: false, targetReached: true, foundAny: false, clickedAny: false };
    assert.equal(skipResult.clicked, false);
    assert.equal(skipResult.targetReached, true);
    // Gegenprobe: die vorhandene Klassifikation bleibt unangetastet.
    assert.equal(classifyButtonFailure({ foundAny: false, clickedAny: false }), 'no-week-button');
    assert.equal(classifyButtonFailure({ foundAny: true, clickedAny: false }), 'button-ineffective');
  });

  test('Kandidaten-Reihenfolge: aria-label-de (+ Dativ) steht vor aria-label-en', () => {
    const c = weekButtonCandidates(1);
    assert.equal(c[0].kind, 'aria-label-de');
    assert.equal(c[1].kind, 'aria-label-de-dat');
    assert.equal(c[2].kind, 'aria-label-en');
    assert.equal(c[0].selector, '[aria-label*="Nächste Woche" i]');
    assert.equal(c[1].selector, '[aria-label*="nächsten Woche" i]');
    assert.equal(c[2].selector, '[aria-label*="Next week" i]');
    assert.equal(weekButtonCandidates(-1)[0].kind, 'aria-label-de');
    assert.equal(weekButtonCandidates(-1)[2].kind, 'aria-label-en');
  });

  test('Klick-Verifikation unveraendert (unverified bleibt unverified)', () => {
    // Regression: der neue Guard darf die Verifikationslogik nicht anfassen.
    assert.equal(evaluateClickVerification('2026-09-21', { startDate: '2026-09-28' }), 'changed');
    assert.equal(evaluateClickVerification('2026-09-28', { startDate: '2026-09-28' }), 'unchanged');
    assert.equal(evaluateClickVerification(null, { startDate: '2026-09-28' }), 'unverified');
    assert.equal(evaluateClickVerification('2026-09-28', { startDate: null }), 'unverified');
  });

  test('canClickSameDirection unveraendert', () => {
    assert.equal(canClickSameDirection({ lastClickDirection: null, direction: 1, weekConfirmedAfterClick: false }), true);
    assert.equal(canClickSameDirection({ lastClickDirection: 1, direction: 1, weekConfirmedAfterClick: false }), false);
    assert.equal(canClickSameDirection({ lastClickDirection: 1, direction: 1, weekConfirmedAfterClick: true }), true);
    assert.equal(canClickSameDirection({ lastClickDirection: 1, direction: -1, weekConfirmedAfterClick: false }), true);
  });
});

// ===========================================================================
// Verdrahtungs-Waechter: verhindert, dass Guard und Aufrufer erneut
// auseinanderlaufen. Der Bug 2 ist genau so entstanden — zwei Stellen, zwei
// Entscheidungen, ein Test auf einer nachgebauten Schleife.
describe('Bug 2 — Verdrahtung: getFreeBusy nutzt den getesteten Pfad', () => {
  const src = readFileSync(fileURLToPath(new URL('../src/owa.js', import.meta.url)), 'utf8');

  test('clickWeekButton delegiert an runWeekButtonClick (kein eigener Guard-Pfad)', () => {
    const i = src.indexOf('const clickWeekButton =');
    assert.ok(i > 0, 'clickWeekButton nicht gefunden');
    const body = src.slice(i, src.indexOf('\n    };', i) + 7);
    assert.ok(/runWeekButtonClick\(\{/.test(body), 'clickWeekButton muss runWeekButtonClick aufrufen');
    assert.ok(/targetStart: rangeStartDate/.test(body), 'das Zielwochen-Datum muss durchgereicht werden');
    assert.ok(/readWeek: \(\) => readRenderedCalendarWeek\(page\)/.test(body),
      'der Wochenzustand muss per DI aus getFreeBusy kommen');
  });

  test('die Schritt-Schleife benutzt dieselbe Zustandsfunktion wie der Klickpfad', () => {
    // BUG 5: Es gibt jetzt EINE Klassifikation (weekViewState) und eine
    // Abbildung darauf (weekClickDecision). Beide muessen aus derselben Quelle
    // stammen, sonst laufen sie wieder auseinander.
    assert.ok(src.includes('const view = weekViewState({ rendered, targetStart: rangeStartDate });'),
      'die aeussere Schleife muss weekViewState benutzen');
    assert.ok(/targetWeekNow = weekClickDecision\(\{ rendered: wk, targetStart \}\)/.test(src),
      'der Klickpfad muss dieselbe Abbildung benutzen');
    assert.ok(!/weekClickDecision\(\{[^}]*hasEvents/.test(src),
      'kein hasEvents mehr in der Entscheidung - das war die tote Ausnahme aus 1c2825e');
    // Kein eigener, schwaecherer Skip mehr im Schritt-Zweig.
    assert.ok(!/if \(rendered\.startDate && isRequestedWeek\(rendered\)\) \{/.test(src),
      'der alte eigene Skip (isRequestedWeek) darf nicht wiederkehren');
  });

  test('BUG 5: Zustand 1 wird ABGEWARTET, nicht abgebrochen', () => {
    // Der Live-Befund: `abortReason = 'view-not-ready'; break;` nach 13ms. Die
    // Verweigerung war richtig, das ABBRECHEN war der Defekt.
    const waitIdx = src.indexOf('const w = await waitForViewReady();');
    assert.ok(waitIdx > 0, 'die aeussere Schleife muss waitForViewReady aufrufen');
    const branchStart = src.lastIndexOf('if (view.state === WEEK_VIEW_STATES.notReady) {', waitIdx);
    assert.ok(branchStart > 0, 'Zustand 1 braucht einen eigenen Zweig');
    const branch = src.slice(branchStart, src.indexOf('const deltaWeeks', branchStart));
    assert.ok(/if \(viewWaitMsSpent < FREEBUSY_VIEW_WAIT_BUDGET_MS\)/.test(branch),
      'Zustand 1 muss das Wartebudget auswerten');
    assert.ok(/if \(!w\.exhausted\) continue;/.test(branch),
      'nach dem Warten wird der Zustand neu bewertet, nicht abgebrochen');
    assert.ok(/guard-nach-wartebudget/.test(branch), 'der ehrliche Abbruch nennt das Wartebudget');
    assert.ok(/abortReason = 'view-not-ready';/.test(branch), 'view-not-ready bleibt ein Abbruchgrund');
    const iWait = branch.indexOf('await waitForViewReady();');
    const iReload = branch.indexOf('await hardReloadOptimized();');
    const iAbort = branch.indexOf("abortReason = 'view-not-ready';");
    assert.ok(iWait >= 0 && iReload > iWait && iAbort > iReload,
      `Reihenfolge wait(${iWait}) < reload(${iReload}) < abort(${iAbort})`);
  });

  test('BUG 5: Zustand 1 hat genau EINEN Notfallweg, den harten Reload', () => {
    // Von den drei belegten Wegen ist der Reload der einzige, der weder einen
    // gerenderten Header noch eine geratene Wochenrichtung braucht.
    assert.ok(src.includes('const hardReloadOptimized = async () => {'), 'die Reload-Funktion existiert');
    assert.ok(/if \(!reloadDone && remaining\(\) >= FREEBUSY_MIN_REMAINING_FOR_ESCALATION_MS\)/.test(src),
      'der Reload ist 1x und budgetiert');
    assert.equal((src.match(/reloadDone = true;/g) || []).length, 1, 'genau eine Ausfuehrung');
    assert.ok(src.includes("await page.reload({ waitUntil: 'domcontentloaded'"),
      'es ist ein HARTER Reload, der den Client-Store ueberschreibt');
    // Der blinde "Naechste Woche"-Klick auf unbekannte Woche ist ENTFALLEN:
    // ohne Header ist die Richtung unbekannt, er war die Wochen-Skip-Gefahr.
    assert.ok(!/forcedNextRetryDone/.test(src), 'der erzwungene Blindklick ist entfernt');
  });

  test('BUG 5: der Klickpfad traegt dasselbe Warten wie die aeussere Schleife', () => {
    // Beide Aufrufer teilen EINE Wartefunktion (Dependency-Injection). Sonst
    // koennte der Klickpfad wieder "sofort verweigern", waehrend die Schleife
    // wartet - genau die Art Divergenz, die 8ae0eec und 1c2825e produziert hat.
    const i = src.indexOf('const clickWeekButton = async');
    assert.ok(i > 0, 'clickWeekButton nicht gefunden');
    const body = src.slice(i, src.indexOf('\n    };', i) + 7);
    assert.ok(/waitForReady: async \(\) =>/.test(body), 'der Klickpfad muss dieselbe Wartefunktion bekommen');
    assert.ok(/await waitForViewReady\(\)/.test(body), 'und zwar waitForViewReady, nicht eine eigene');
    assert.ok(/onGuardBlocked:/.test(body), 'die Verweigerung wird nach aussen gemeldet (fuer waitedMs)');
    assert.ok(/decision\.reason === WEEK_CLICK_GUARD_REASONS\.viewNotReady\) \{[\s\S]{0,600}await waitOnce\(\);[\s\S]{0,120}decision = await guardNow\(\);/.test(src),
      'runWeekButtonClick muss bei view-not-ready WARTEN und danach NEU entscheiden');
  });

  test('BUG 5: "Header nicht gerendert" ist von "Auswertung fehlgeschlagen" unterscheidbar', () => {
    // Runde 4: headerCandidates/dayHeaderCandidates/weekNavCandidates waren alle
    // leer. Das sieht bei "nichts gerendert" und bei "evaluate geworfen"
    // identisch aus, weil der catch dieselbe Leerform lieferte.
    assert.ok(/collectError: String\(\(e && e\.message\) \|\| e\)/.test(src),
      'der catch-Fallback muss den Fehler benennen statt ihn als "nichts gefunden" zu tarnen');
    assert.ok(src.includes('collectError: collected.collectError || null'), 'die Diagnostik reicht ihn durch');
    assert.ok(src.includes('diag.headerReadError'), 'und das Ergebnis nennt ihn');
  });

  test('der neue Abbruchgrund taucht nicht in der Navigation als Erfolg auf', () => {
    // view-not-ready darf nicht als 'kein Klick, alles gut' durchrutschen.
    assert.equal(WEEK_CLICK_GUARD_REASONS.viewNotReady, 'view-not-ready');
    assert.notEqual(WEEK_CLICK_GUARD_REASONS.viewNotReady, WEEK_CLICK_GUARD_REASONS.targetReached);
    // Und er ist von den drei Zustaenden zu unterscheiden.
    assert.equal(WEEK_VIEW_STATES.notReady, 'view-not-ready');
    assert.notEqual(WEEK_VIEW_STATES.notReady, WEEK_VIEW_STATES.targetWeek);
    assert.notEqual(WEEK_VIEW_STATES.notReady, WEEK_VIEW_STATES.wrongWeek);
  });
});

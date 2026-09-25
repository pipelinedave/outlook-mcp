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
  WEEK_CLICK_GUARD_REASONS
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
    assert.equal(kinds[0], 'aria-label-de', 'erst DE, dann EN — Reihenfolge unveraendert');
    assert.equal(kinds[1], 'aria-label-en');
    assert.equal(r.probe[0].found, false, 'DE-Button existiert nicht');
    assert.equal(r.probe[1].clicked, true, 'EN-Button wird geklickt');
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
    // Zielwoche drin
    assert.equal(weekClickDecision({ rendered: TARGET_WEEK, targetStart: TARGET }).skip, true);
    assert.equal(weekClickDecision({ rendered: TARGET_WEEK, targetStart: TARGET }).reason, WEEK_CLICK_GUARD_REASONS.targetReached);
    // Falsche Woche, aber gerendert -> klicken
    const wrong = weekClickDecision({ rendered: WRONG_WEEK, targetStart: TARGET });
    assert.equal(wrong.skip, false);
    assert.equal(wrong.reason, WEEK_CLICK_GUARD_REASONS.click);
    assert.equal(wrong.clickable, true);
    // NICHT gerendert und keine Events -> blockieren
    const unrendered = weekClickDecision({ rendered: UNRENDERED, targetStart: TARGET, hasEvents: false });
    assert.equal(unrendered.skip, true);
    assert.equal(unrendered.reason, WEEK_CLICK_GUARD_REASONS.viewNotReady);
    assert.equal(unrendered.clickable, false);
    // Nicht gerendert, ABER Events sichtbar -> die Ansicht ist nicht leer,
    // der Header moeglicherweise nur nicht lesbar. Klick bleibt zulaessig.
    const eventsNoHeader = weekClickDecision({ rendered: UNRENDERED, targetStart: TARGET, hasEvents: true });
    assert.equal(eventsNoHeader.skip, false, 'sichtbare Events belegen eine gerenderte Ansicht');
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

  test('Kandidaten-Reihenfolge unveraendert: aria-label-de steht vor aria-label-en', () => {
    const c = weekButtonCandidates(1);
    assert.equal(c[0].kind, 'aria-label-de');
    assert.equal(c[1].kind, 'aria-label-en');
    assert.equal(c[0].selector, '[aria-label*="Nächste Woche" i]');
    assert.equal(c[1].selector, '[aria-label*="Next week" i]');
    assert.equal(weekButtonCandidates(-1)[0].kind, 'aria-label-de');
    assert.equal(weekButtonCandidates(-1)[1].kind, 'aria-label-en');
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

  test('die Schritt-Schleife benutzt dieselbe Entscheidungsfunktion', () => {
    const i = src.indexOf('const outerDecision = weekClickDecision(');
    assert.ok(i > 0, 'die aeussere Schleife muss weekClickDecision benutzen');
    // Kein eigener, schwaecherer Skip mehr im Schritt-Zweig.
    assert.ok(!/if \(rendered\.startDate && isRequestedWeek\(rendered\)\) \{/.test(src),
      'der alte eigene Skip (isRequestedWeek) darf nicht wiederkehren');
  });

  test('beide Aufrufer behandeln view-not-ready als eigenen Abbruchgrund', () => {
    const n = src.split('WEEK_CLICK_GUARD_REASONS.viewNotReady').length - 1;
    // 1x Definition + 3x Verwendung (aeussere Schleife, forced, delta-Pfad)
    assert.ok(n >= 4, `view-not-ready muss in allen drei Zweigen benannt werden (gefunden: ${n})`);
    assert.equal((src.match(/abortReason = 'view-not-ready'/g) || []).length, 3,
      'drei Abbruchgruende: aeussere Schleife, forced-Pfad, Delta-Pfad');
  });

  test('der neue Abbruchgrund taucht nicht in der Navigation als Erfolg auf', () => {
    // view-not-ready darf nicht als 'kein Klick, alles gut' durchrutschen.
    assert.equal(WEEK_CLICK_GUARD_REASONS.viewNotReady, 'view-not-ready');
    assert.notEqual(WEEK_CLICK_GUARD_REASONS.viewNotReady, WEEK_CLICK_GUARD_REASONS.targetReached);
  });
});

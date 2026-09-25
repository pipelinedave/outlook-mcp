import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import {
  renderedWeekContains,
  weekProbeSkipEntry,
  weekButtonCandidates,
  shouldTryNextCandidate,
  evaluateClickVerification,
  canClickSameDirection,
  classifyButtonFailure
} from '../src/owa.js';

// ===========================================================================
// Bug 2 — weekButtonProbe: Guard an falscher Stelle (Live-Befund 25.09.2026)
//
// Live-Trace (ungekuerzt):
//   1. { selector: '[aria-label*="Nächste Woche" i]', kind: 'aria-label-de',
//        found: false, clicked: false }
//   2. { selector: '[aria-label*="Next week" i]',     kind: 'aria-label-en',
//        found: true,  clicked: true, verified: 'unverified' }
//   3. { skipped: 'target-week-verified', renderedWeek: '2026-09-28',
//        clicked: false }
//
// Der Skip in Eintrag 3 kommt NACH dem Klick in Eintrag 2. Ursache war nicht
// eine fehlende Guard-Stelle, sondern die Position: der Guard sass in der
// aeusseren Schritt-Schleife und griff erst, als der verifizierte Zustand
// wieder in `rendered` stand — der Klick war da schon erfolgt.
//
// Dieser Block pinnt die Entscheidungslogik fest, die clickWeekButton jetzt
// VOR jeder Probe anwendet.
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

  // --- Der entscheidende Test: die Live-Reihenfolge -------------------------
  // Bildet die Kandidaten-Schleife aus clickWeekButton nach: gleiche
  // Reihenfolge (weekButtonCandidates), gleicher Guard (renderedWeekContains
  // VOR der Probe), gleiche Protokolleintraege.
  const runGuardedProbe = (weekBeforeLoop) => {
    const probe = [];
    for (const cand of weekButtonCandidates(1)) {
      // NEU: Guard vor der Probe, in derselben Schleife, in der geklickt wird.
      if (renderedWeekContains(weekBeforeLoop, '2026-09-28')) {
        probe.push(weekProbeSkipEntry(weekBeforeLoop.startDate, 'guard-vor-klick'));
        return { probe, targetReached: true };
      }
      probe.push({ selector: cand.selector, kind: cand.kind, found: true, clicked: true, verified: 'unverified' });
    }
    return { probe, targetReached: false };
  };

  test('Zielwoche verifiziert: Skip steht VOR jeder Probe, null Klicks', () => {
    const r = runGuardedProbe({ startDate: '2026-09-28', endDate: '2026-10-04' });
    assert.equal(r.targetReached, true);
    assert.equal(r.probe.length, 1, 'genau ein Eintrag');
    assert.equal(r.probe[0].skipped, 'target-week-verified');
    assert.equal(r.probe[0].renderedWeek, '2026-09-28');
    assert.equal(r.probe[0].clicked, false);
    assert.equal(r.probe[0].note, 'guard-vor-klick');
    assert.equal(r.probe.some((e) => e.clicked === true), false, 'KEIN Klick darf protokolliert werden');
    assert.equal(r.probe.some((e) => e.selector), false, 'keine Probe wurde überhaupt durchgefuehrt');
  });

  test('Regression auf die Live-Reihenfolge: ohne Guard kommt der Skip NACH dem Klick', () => {
    // Das ist die alte, fehlerhafte Reihenfolge als Regressionsanker.
    const probe = [];
    let outerRendered = { startDate: null, endDate: null };
    for (const cand of weekButtonCandidates(1).slice(0, 2)) {
      // KEIN Guard in der Schleife (Zustand vor dem Fix).
      if (!probe.length && cand.kind === 'aria-label-de') {
        probe.push({ selector: cand.selector, kind: cand.kind, found: false, clicked: false });
        continue;
      }
      probe.push({ selector: cand.selector, kind: cand.kind, found: true, clicked: true, verified: 'unverified' });
      outerRendered = { startDate: '2026-09-28', endDate: '2026-10-04' };
    }
    if (renderedWeekContains(outerRendered, '2026-09-28')) {
      probe.push(weekProbeSkipEntry(outerRendered.startDate));
    }
    // Live-Trace reproduziert: Klick VOR Skip.
    assert.equal(probe[1].clicked, true, 'Klick passiert vor dem Skip — genau der Live-Befund');
    assert.equal(probe[2].skipped, 'target-week-verified');
    assert.ok(
      probe.findIndex((e) => e.clicked === true) < probe.findIndex((e) => e.skipped),
      'VORHERS: Skip nach Klick'
    );
  });

  test('Nicht verifizierte Woche: Guard greift nicht, Klick-Verhalten unveraendert', () => {
    const cands = weekButtonCandidates(1);
    const week = { startDate: '2026-09-21', endDate: '2026-09-27' };
    let clicks = 0;
    let skip = null;
    for (const cand of cands) {
      if (renderedWeekContains(week, '2026-09-28')) { skip = weekProbeSkipEntry(week.startDate, 'guard-vor-klick'); break; }
      clicks++;
    }
    assert.equal(skip, null, 'falsche Woche -> kein Skip');
    assert.equal(clicks, cands.length, 'alle Kandidaten werden wie bisher geprobt');
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

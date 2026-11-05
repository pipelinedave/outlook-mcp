import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import {
  parseEventLabel,
  isPickerControlLabel,
  CALENDAR_PICKER_LABEL_PATTERNS,
  resolveRenderedWeekFromCandidates
} from '../src/owa.js';

// ===========================================================================
// Runde 2, Befund 2 (Live-Lauf 05.11.2026, d-velop-Realm, KW47): Der
// Monats-Picker-BUTTON "November 2026, select to change the month" wurde als
// Datumsquelle missbraucht: parseEventLabel las "November 20(26)" als Tag 20
// und leitete daraus die Woche 16.–22.11. ab. Die "korrekte renderedWeek" war
// reiner Zufall (KW47 enthaelt den 20.11.) — gerendert war die aktuelle Woche
// (KW45), deshalb "8 Event-Knoten, aber keine im angefragten Zeitraum" und
// dataQuality unverified.
//
// Zweischichtiger Fix:
//   1. parseEventLabel: Tageserkennung mit (?<!\d)/(?!\d)-Grenzen — "2026"
//      kann nie mehr als "Tag 20(26)" gelesen werden.
//   2. isPickerControlLabel: Picker-/Flyout-Steuerelemente (Monats-/Datum-
//      Picker, "Go to today") sind nie angezeigte Header.
// ===========================================================================

describe('parseEventLabel — Tages-Boundary (Runde 2)', () => {
  test('LIVE-FALL: "November 2026, select to change the month" ist KEIN Datum', () => {
    // Vor dem Fix: { date: '2026-11-20' } — das "26" von 2026 als Rest des Tages.
    assert.equal(parseEventLabel('November 2026, select to change the month', 2026), null);
  });

  test('Jahreszahl im Text wird nicht als Tagesrest gelesen (DE/EN gemischt)', () => {
    assert.equal(parseEventLabel('Kickoff 2026 Planning', 2026), null);
    assert.equal(parseEventLabel('Review 2026: Status', 2026), null);
  });

  test('Legitime Datumsangaben funktionieren weiter', () => {
    assert.equal(parseEventLabel('Treffen 2. November 2026', 2026)?.date, '2026-11-02');
    assert.equal(parseEventLabel('Standup November 5, 2026', 2026)?.date, '2026-11-05');
    // Tag-zuerst MIT Jahr: Tag 3 + Jahr 2026 (vor dem Fix: "Tag 20" aus dem Jahr)
    assert.equal(parseEventLabel('3 November 2026 Besprechung', 2026)?.date, '2026-11-03');
    // Tag ohne Jahr: defaultYear greift ("Montag, 16 November" Spaltenkopf-Stil)
    assert.equal(parseEventLabel('Montag, 16 November', 2026)?.date, '2026-11-16');
  });

  test('Event mit Datum + Zeitfenster bleibt intakt', () => {
    const r = parseEventLabel('Sprint Review 2. November 2026, 14:00 bis 15:00', 2026);
    assert.equal(r.date, '2026-11-02');
    assert.equal(r.start, '14:00');
    assert.equal(r.end, '15:00');
    assert.match(r.title, /Sprint Review/);
    assert.equal(r.allDay, false);
  });
});

describe('isPickerControlLabel (Runde 2)', () => {
  test('LIVE-Belege: Picker-Buttons werden erkannt', () => {
    assert.equal(isPickerControlLabel('November 2026, select to change the month'), true);
    assert.equal(isPickerControlLabel('November 2026, Selected date November 5, 2026, Today\'s date November 5, 2026'), true);
    assert.equal(isPickerControlLabel('Go to today November 5, 2026'), true);
  });

  test('Der ECHTE Wochen-Header bleibt unberuehrt ("Jump to..."-Zusatz)', () => {
    assert.equal(isPickerControlLabel('2–8 November, 2026, Jump to a specific date or date range.'), false);
    assert.equal(isPickerControlLabel('16–22 November, 2026, Jump to a specific date or date range.'), false);
    assert.equal(isPickerControlLabel('21–27 September, 2026 (Week 39)'), false);
    assert.equal(isPickerControlLabel(''), false);
    assert.equal(isPickerControlLabel(null), false);
  });

  test('Pattern-Liste ist serialisierbar (evaluate-Grenze)', () => {
    for (const p of CALENDAR_PICKER_LABEL_PATTERNS) assert.equal(typeof p, 'string');
  });
});

describe('resolveRenderedWeekFromCandidates — Picker-Phantomwochen (Runde 2, KW47)', () => {
  const PHANTOM = {
    // Exakter Live-Snapshot des KW47-Fehllaufs: nur Picker-Labels + Wochenzahlen
    // des Picker-Flyouts, KEIN echter Wochen-Header, KEINE prev/next-Labels.
    headerTexts: [
      'November 2026, select to change the month',
      'November 2026, Selected date November 5, 2026, Today\'s date November 5, 2026'
    ],
    dayHeaderTexts: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
    navTexts: ['week number 44', 'week number 45', 'week number 46', 'week number 47', 'week number 48', 'week number 49']
  };

  test('LIVE-FALL KW47: Picker-only-Snapshot liefert KEINE Phantom-Woche 16.–22.11.', () => {
    const r = resolveRenderedWeekFromCandidates(PHANTOM, new Date('2026-11-05T12:00:00'));
    // Vor dem Fix: startDate 2026-11-16 (aus "November 20(26)") — Zufallstreffer.
    assert.equal(r.startDate, null);
    assert.equal(r.endDate, null);
    assert.equal(r.readError, 'no-displayed-header');
  });

  test('LIVE-FALL aktuelle Woche: echter Header gewinnt weiter (Regressionsschutz)', () => {
    const r = resolveRenderedWeekFromCandidates({
      headerTexts: [
        'Go to today November 5, 2026',
        '2–8 November, 2026, Jump to a specific date or date range.',
        'November 2026, select to change the month',
        'November 2026, Selected date November 5, 2026, Today\'s date November 5, 2026'
      ],
      dayHeaderTexts: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
      navTexts: [
        'week number 44', 'week number 45', 'week number 46',
        'Go to previous week 26 October – 1 November, 2026 (Week 44)',
        'Go to next week 9–15 November, 2026 (Week 46)'
      ]
    }, new Date('2026-11-05T12:00:00'));
    assert.equal(r.startDate, '2026-11-02');
    assert.equal(r.endDate, '2026-11-08');
    assert.equal(r.source, 'displayed-header');
    assert.match(r.text, /2–8 November, 2026/);
  });
});

import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import {
  parseConversationDateText,
  selectConversationTargetIndex
} from '../src/owa.js';

// ===========================================================================
// Defekt A (05.11.2026): outlook_get_email renderte bei Konversations-IDs die
// AELTESTE statt die getroffene (neueste) Thread-Mail. Der Lesebereich stapelt
// die Konversation als Karten (aelteste zuerst, inkl. Draft-Karten); der
// Renderer muss die neueste sendende Karte auswaehlen. Hier: die reinen
// Hilfsfunktionen (Datum-Parsing + Zielauswahl).
// Live-Beleg (d-velop-Realm, Thread "Rueckfragen Aenderungen
// DokuFIS-Schnittstelle"): 12 Karten, 6.10.-19.10.2026, 3 Drafts dazwischen.
// ===========================================================================

describe('parseConversationDateText', () => {
  test('DE voll: "Gesendet: Dienstag, 6. Oktober 2026 11:22"', () => {
    assert.equal(parseConversationDateText('Gesendet: Dienstag, 6. Oktober 2026 11:22'), '2026-10-06 11:22');
  });

  test('EN kompakt: "Mon 19 Oct 2026 11:51" (Live-Header der neuesten Karte)', () => {
    assert.equal(parseConversationDateText('Mon 19 Oct 2026 11:51'), '2026-10-19 11:51');
  });

  test('EN kompakt: "Tue 6 Oct 2026 15:39"', () => {
    assert.equal(parseConversationDateText('Tue 6 Oct 2026 15:39'), '2026-10-06 15:39');
  });

  test('EN voll mit PM: "Sent: Tuesday, October 6, 2026 3:39 PM"', () => {
    assert.equal(parseConversationDateText('Sent: Tuesday, October 6, 2026 3:39 PM'), '2026-10-06 15:39');
  });

  test('DE kompakt: "Di 6. Okt. 2026 15:39"', () => {
    assert.equal(parseConversationDateText('Di 6. Okt. 2026 15:39'), '2026-10-06 15:39');
  });

  test('Draft: "Saved: Tue 13 Oct 2026 13:38"', () => {
    assert.equal(parseConversationDateText('Saved: Tue 13 Oct 2026 13:38'), '2026-10-13 13:38');
  });

  test('Kein Datum: Leerstring/Bodyspeak ohne Datumszeile', () => {
    assert.equal(parseConversationDateText(''), null);
    assert.equal(parseConversationDateText('Hallo Herr Mlynarek'), null);
    assert.equal(parseConversationDateText(null), null);
  });
});

describe('selectConversationTargetIndex', () => {
  const LIVE_THREAD = [
    // Reihenfolge = DOM-Reihenfolge (aelteste zuerst), gekuerzt aus dem Live-Befund
    { dateText: 'Gesendet: Dienstag, 6. Oktober 2026 11:22' },
    { dateText: 'Tue 6 Oct 2026 15:39' },
    { dateText: 'Tue 6 Oct 2026 16:34' },
    { dateText: 'Thu 8 Oct 2026 11:15' },
    { dateText: 'Mon 12 Oct 2026 16:31' },
    { dateText: 'Saved: Tue 13 Oct 2026 13:38', isDraft: true },
    { dateText: 'Saved: Tue 13 Oct 2026 13:47', isDraft: true },
    { dateText: 'Tue 13 Oct 2026 13:47' },
    { dateText: 'Tue 13 Oct 2026 14:19' },
    { dateText: 'Saved: Thu 15 Oct 2026 13:08', isDraft: true },
    { dateText: 'Thu 15 Oct 2026 13:08' },
    { dateText: 'Mon 19 Oct 2026 11:51' }
  ];

  test('Live-Thread: neueste SENDENDE Karte (19 Oct), Drafts und aelteste ignoriert', () => {
    assert.equal(selectConversationTargetIndex(LIVE_THREAD), 11);
  });

  test('Draft mit spaetestem Zeitstempel gewinnt NIE (Drafts ausgeschlossen)', () => {
    const cards = [
      { dateText: 'Mon 19 Oct 2026 11:51' },
      { dateText: 'Saved: Fri 30 Oct 2026 09:00', isDraft: true }
    ];
    assert.equal(selectConversationTargetIndex(cards), 0);
  });

  test('Nur Drafts -> letzte Karte (Fallback, kein Crash)', () => {
    const cards = [
      { dateText: 'Saved: Tue 13 Oct 2026 13:38', isDraft: true },
      { dateText: 'Saved: Thu 15 Oct 2026 13:08', isDraft: true }
    ];
    assert.equal(selectConversationTargetIndex(cards), 1);
  });

  test('Ohne parsebares Datum: letzte Karte in DOM-Reihenfolge', () => {
    const cards = [{ dateText: null }, { dateText: null }];
    assert.equal(selectConversationTargetIndex(cards), 1);
  });

  test('Leere Eingabe -> -1', () => {
    assert.equal(selectConversationTargetIndex([]), -1);
    assert.equal(selectConversationTargetIndex(null), -1);
  });
});

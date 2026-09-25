import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { resolveFolder, buildComposeUrl, buildCalendarComposeUrl, buildCalendarWeekUrl, normalizeCalendarDate, FOLDER_MAP, parseWeekHeaderText, parseEventLabel, classifyEventStatus, parseMailRowTimestamp, parseNumericDayHeaderText, resolveRenderedWeekFromCandidates, isWorkLocationPseudoSlot, dedupeBusySlots, partitionSlotsByRange, formatCalendarDiagnostics, FREEBUSY_TOTAL_BUDGET_MS, FREEBUSY_MAX_BUTTON_CLICKS, FREEBUSY_CLICK_RENDER_WAIT_MS, FREEBUSY_MIN_REMAINING_FOR_ESCALATION_MS, phaseBudgetMs, affordableClicks, weekButtonCandidates, evaluateClickVerification, shouldTryNextCandidate, classifyButtonFailure, assembleFreeBusyDiagnostics, slotExtractionWarning, collectCalendarEventLabels, isCalendarEventLabel, filterCalendarHeaderCandidates, isVisibleCalendarElement, parseWeekNumberFromNav, parseNavWeekRange, isoWeekMonday, resolveWeekFromNavButtons, MAILBOX_VERIFICATION_SOURCES, resolveMailboxVerification, isEventGridLabel, EVENT_LABEL_PATTERNS, eventWaitBudgetMs, pollUntilEventGrid, FREEBUSY_EVENT_WAIT_MAX_MS, FREEBUSY_EVENT_WAIT_RESERVE_MS, FREEBUSY_EVENT_WAIT_POLL_MS, EVENT_WAIT_OUTCOMES, canClickSameDirection, weekProbeSkipEntry, eventWaitWarning, CALENDAR_EVENT_SOURCES, eventSourcePlan, normalizeClockTime, mapXToDayColumn, mapYToClockTime, estimateHourPx, buildGeometryEventLabel, deriveEventLabelFromNode, resolveDataQuality, DATA_QUALITY_VALUES, resolveEventWaitOutcome } from '../src/owa.js';
import { config } from '../src/config.js';

describe('resolveFolder - Outlook Web Ordner-Mapping', () => {
  test('Standardordner Deutsch & Englisch auflösen', () => {
    assert.equal(resolveFolder('inbox'), 'inbox');
    assert.equal(resolveFolder('posteingang'), 'inbox');
    assert.equal(resolveFolder('Posteingang'), 'inbox');
    assert.equal(resolveFolder('sent'), 'sentitems');
    assert.equal(resolveFolder('sentitems'), 'sentitems');
    assert.equal(resolveFolder('gesendet'), 'sentitems');
    assert.equal(resolveFolder('Gesendet'), 'sentitems');
    assert.equal(resolveFolder('archive'), 'archive');
    assert.equal(resolveFolder('archiv'), 'archive');
    assert.equal(resolveFolder('drafts'), 'drafts');
    assert.equal(resolveFolder('entwürfe'), 'drafts');
    assert.equal(resolveFolder('Entwürfe'), 'drafts');
    assert.equal(resolveFolder('deleteditems'), 'deleteditems');
    assert.equal(resolveFolder('papierkorb'), 'deleteditems');
    assert.equal(resolveFolder('newsletter'), 'newsletter');
  });

  test('Fallback bei unbekanntem Ordner', () => {
    assert.equal(resolveFolder('custom-folder'), 'custom-folder');
  });

  test('Leerer Wert fällt auf inbox zurück', () => {
    assert.equal(resolveFolder(''), 'inbox');
    assert.equal(resolveFolder(null), 'inbox');
    assert.equal(resolveFolder(undefined), 'inbox');
  });
});

describe('buildComposeUrl - DeepLink Generierung', () => {
  test('Erzeugt korrekte Compose-URL mit Parametern und Encodierung', () => {
    const url = buildComposeUrl('d-velop.de', {
      to: 'kollege@firma.de',
      subject: 'Test Betreff',
      body: 'Hallo Welt!\nDies ist ein Test.',
      cc: 'chef@firma.de'
    });

    assert.ok(url.startsWith('https://outlook.office.com/mail/deeplink/compose?'));
    assert.ok(url.includes('realm=d-velop.de'));
    assert.ok(url.includes('to=kollege%40firma.de'));
    assert.ok(url.includes('subject=Test%20Betreff'));
    assert.ok(url.includes('cc=chef%40firma.de'));
    assert.ok(url.includes('body=Hallo%20Welt!%0ADies%20ist%20ein%20Test.'));
  });

  test('Mehrere Empfänger als Array oder String', () => {
    const url = buildComposeUrl('adesso.de', {
      to: ['a@adesso.de', 'b@adesso.de'],
      subject: 'Sync',
      body: 'Hallo'
    });
    assert.ok(url.includes('realm=adesso.de'));
    assert.ok(url.includes('to=a%40adesso.de%3B%20b%40adesso.de'));
  });
});

describe('buildCalendarComposeUrl - Kalender DeepLink Generierung', () => {
  test('Erzeugt korrekte Kalender-Compose-URL mit Titel, Start/Ende und Body', () => {
    const url = buildCalendarComposeUrl('d-velop.de', {
      subject: 'Blocker für Vorbereitung UKO',
      start: '2026-09-15T13:00:00',
      end: '2026-09-15T17:00:00',
      body: 'PoC-Architektur testen'
    });
    assert.ok(url.startsWith('https://outlook.office.com/calendar/deeplink/compose?'));
    assert.ok(url.includes('realm=d-velop.de'));
    assert.ok(url.includes('subject=Blocker%20f%C3%BCr%20Vorbereitung%20UKO'));
    assert.ok(url.includes('startdt=2026-09-15T13%3A00%3A00'));
    assert.ok(url.includes('enddt=2026-09-15T17%3A00%3A00'));
    assert.ok(url.includes('body=PoC-Architektur%20testen'));
  });

  test('Ganztägig-Flag wird gesetzt', () => {
    const url = buildCalendarComposeUrl('d-velop.de', {
      subject: 'Urlaub',
      start: '2026-09-15',
      end: '2026-09-16',
      is_all_day: true
    });
    assert.ok(url.includes('allday=true'));
  });

  test('Ohne Body und Realm keine leeren Parameter', () => {
    const url = buildCalendarComposeUrl('d-velop.de', {
      subject: 'Test',
      start: '2026-09-15T10:00:00',
      end: '2026-09-15T11:00:00'
    });
    assert.ok(!url.includes('body='));
    assert.ok(url.includes('realm=d-velop.de'));
  });
});

describe('buildCalendarWeekUrl - Wochenansicht-Navigation (Live-Fail-Fix 24.09.)', () => {
  test('Realm ZUERST, startdate als zweiter Parameter (bewährte Reihenfolge)', () => {
    const url = buildCalendarWeekUrl('d-velop.de', '2026-09-28');
    assert.equal(url, 'https://outlook.office.com/calendar/view/week?realm=d-velop.de&startdate=2026-09-28');
  });

  test('Ohne Realm nur startdate', () => {
    const url = buildCalendarWeekUrl(null, '2026-09-28');
    assert.equal(url, 'https://outlook.office.com/calendar/view/week?startdate=2026-09-28');
  });

  test('Realm wird URL-encodiert', () => {
    const url = buildCalendarWeekUrl('adesso.de', '2026-10-05');
    assert.ok(url.startsWith('https://outlook.office.com/calendar/view/week?realm=adesso.de&startdate=2026-10-05'));
  });
});

describe('normalizeCalendarDate - Start/Ende Normalisierung', () => {
  test('ISO-Datum ohne Offset bleibt als Berlin-Lokalzeit (kein Z - OWA-Deep-Link-Falle)', () => {
    // OWA interpretiert startdt ohne Offset als lokale Browserzeit (Europe/Berlin).
    // Eine UTC-Konvertierung (Z) würde 08:00 gewollt-Berlin auf 10:00 verschieben.
    const out = normalizeCalendarDate('2026-09-15T13:00:00');
    assert.equal(out, '2026-09-15T13:00:00');
  });

  test('ISO-Datum mit Offset wird zu UTC-ISO normalisiert', () => {
    const out = normalizeCalendarDate('2026-09-15T13:00:00+02:00');
    assert.equal(out, '2026-09-15T11:00:00.000Z');
  });

  test('ISO-Datum mit explizitem Z bleibt unverändert', () => {
    const out = normalizeCalendarDate('2026-09-15T13:00:00Z');
    assert.equal(out, '2026-09-15T13:00:00.000Z');
  });

  test('Sekunden werden bei fehlenden Werten ergänzt', () => {
    const out = normalizeCalendarDate('2026-09-15T08:00');
    assert.equal(out, '2026-09-15T08:00:00');
  });

  test('Leerer/undefinierter Wert gibt undefined zurück', () => {
    assert.equal(normalizeCalendarDate(undefined), undefined);
    assert.equal(normalizeCalendarDate(''), undefined);
  });
});

describe('config - Defaults & Umgebungsvariablen', () => {
  test('ProfileDir ist gesetzt', () => {
    assert.ok(config.profileDir.length > 0);
  });

  test('DefaultRealm ist ein String', () => {
    assert.equal(typeof config.defaultRealm, 'string');
    assert.ok(config.defaultRealm.length > 0);
  });
});

describe('owaClient - replyDraft Validierung', () => {
  test('replyDraft Methode existiert auf owaClient', async () => {
    const { owaClient } = await import('../src/owa.js');
    assert.equal(typeof owaClient.replyDraft, 'function');
  });

  test('replyDraft wirft Fehler bei fehlendem oder leerem body', async () => {
    const { owaClient } = await import('../src/owa.js');
    await assert.rejects(
      async () => {
        await owaClient.replyDraft({});
      },
      /Parameter "body" ist erforderlich/
    );

    await assert.rejects(
      async () => {
        await owaClient.replyDraft({ body: '   ' });
      },
      /Parameter "body" ist erforderlich/
    );
  });
});

describe('parseWeekHeaderText - Wochen-Header der Kalenderansicht', () => {
  test('DE-Header mit Jahr nur am Ende ("28. September – 4. Oktober 2026")', () => {
    const out = parseWeekHeaderText('28. September – 4. Oktober 2026', new Date(2026, 8, 24));
    assert.deepEqual(out, { startDate: '2026-09-28', endDate: '2026-10-04' });
  });

  test('DE-Header mit Jahr an beiden Daten', () => {
    const out = parseWeekHeaderText('28. September 2026 – 4. Oktober 2026', new Date(2026, 8, 24));
    assert.deepEqual(out, { startDate: '2026-09-28', endDate: '2026-10-04' });
  });

  test('EN-Header ("September 28 - October 4, 2026")', () => {
    const out = parseWeekHeaderText('September 28 - October 4, 2026', new Date(2026, 8, 24));
    assert.deepEqual(out, { startDate: '2026-09-28', endDate: '2026-10-04' });
  });

  test('EN-Header gleicher Monat ("September 21 - September 27, 2026")', () => {
    const out = parseWeekHeaderText('September 21 - September 27, 2026', new Date(2026, 8, 24));
    assert.deepEqual(out, { startDate: '2026-09-21', endDate: '2026-09-27' });
  });

  test('Jahreswechsel ("29. Dezember – 4. Januar 2026") hebt das Enddatum ins Folgejahr', () => {
    const out = parseWeekHeaderText('29. Dezember – 4. Januar 2026', new Date(2026, 8, 24));
    assert.deepEqual(out, { startDate: '2026-12-29', endDate: '2027-01-04' });
  });

  test('Falsche Woche (KW39-Bugfall) wird korrekt geparst und vom Verifikationsvergleich erkannt', () => {
    const out = parseWeekHeaderText('21. September – 27. September 2026', new Date(2026, 8, 24));
    assert.deepEqual(out, { startDate: '2026-09-21', endDate: '2026-09-27' });
    // Anfrage ab 28.09. liegt NICHT in der gerenderten Woche 21.-27.09.
    assert.ok(!(out.startDate <= '2026-09-28' && '2026-09-28' <= out.endDate));
  });

  test('Nicht parsebarer Text liefert null', () => {
    assert.equal(parseWeekHeaderText('Posteingang', new Date(2026, 8, 24)), null);
    assert.equal(parseWeekHeaderText('', new Date(2026, 8, 24)), null);
    assert.equal(parseWeekHeaderText(null), null);
  });

  test('Live-Fail 24.09.: Ansichtsumschalter "Work week" ist KEIN Wochen-Header (null)', () => {
    // Der alte Code gab candidates[0]="Work week" mit null-Daten zurück — der
    // Falschgreifer muss durch die Datums-Muster-Validierung fallen.
    assert.equal(parseWeekHeaderText('Work week', new Date(2026, 8, 24)), null);
    assert.equal(parseWeekHeaderText('Arbeitswoche', new Date(2026, 8, 24)), null);
    assert.equal(parseWeekHeaderText('Month', new Date(2026, 8, 24)), null);
    assert.equal(parseWeekHeaderText('Monat', new Date(2026, 8, 24)), null);
  });

  test('DE-Header ohne Jahr nutzt das aktuelle Jahr', () => {
    const out = parseWeekHeaderText('28. September – 2. Oktober', new Date(2026, 8, 24));
    assert.deepEqual(out, { startDate: '2026-09-28', endDate: '2026-10-02' });
  });

  test('Numerischer Wochen-Header ("28.9. – 4.10.2026") wird geparst (Runde-2-Lücke)', () => {
    const out = parseWeekHeaderText('28.9. – 4.10.2026', new Date(2026, 8, 24));
    assert.deepEqual(out, { startDate: '2026-09-28', endDate: '2026-10-04' });
  });

  test('Numerischer Wochen-Header ohne Jahr ("28.9. - 4.10.") nutzt das aktuelle Jahr', () => {
    const out = parseWeekHeaderText('28.9. - 4.10.', new Date(2026, 8, 24));
    assert.deepEqual(out, { startDate: '2026-09-28', endDate: '2026-10-04' });
  });

  test('Live-Format Runde 4: Tages-zuerst EN ohne Punkt ("28 September – 4 October, 2026")', () => {
    // Der ECHTE sichtbare OWA-Wochen-Header. Vorher griff keines der Muster ->
    // die Erkennung fiel auf ein Event-Label zurück (Root Cause Runde 4).
    const out = parseWeekHeaderText('28 September – 4 October, 2026', new Date(2026, 8, 24));
    assert.deepEqual(out, { startDate: '2026-09-28', endDate: '2026-10-04' });
  });

  test('Live-Format Runde 4 mit Suffix ("…, Jump to a specific date or date range.")', () => {
    const out = parseWeekHeaderText(
      '28 September – 4 October, 2026, Jump to a specific date or date range.',
      new Date(2026, 8, 24)
    );
    assert.deepEqual(out, { startDate: '2026-09-28', endDate: '2026-10-04' });
  });

  test('Tages-zuerst EN ohne Jahr nutzt das aktuelle Jahr', () => {
    const out = parseWeekHeaderText('28 September – 4 October', new Date(2026, 8, 24));
    assert.deepEqual(out, { startDate: '2026-09-28', endDate: '2026-10-04' });
  });
});

describe('parseEventLabel - strukturierte Busy-Slots aus ARIA-Labels', () => {
  test('DE-Label mit Zeitfenster ("bis")', () => {
    const out = parseEventLabel('Team Meeting, 14:00 bis 15:00 Montag, 28. September', 2026);
    assert.equal(out.date, '2026-09-28');
    assert.equal(out.start, '14:00');
    assert.equal(out.end, '15:00');
    assert.equal(out.title, 'Team Meeting');
    assert.equal(out.allDay, false);
  });

  test('DE-Label mit "von ... bis ..." entfernt die Präposition aus dem Titel', () => {
    const out = parseEventLabel('Besetzt von 09:00 bis 10:00 Montag, 28. September', 2026);
    assert.equal(out.date, '2026-09-28');
    assert.equal(out.start, '09:00');
    assert.equal(out.end, '10:00');
    assert.equal(out.title, 'Besetzt');
  });

  test('EN-Label mit AM/PM ("Busy 9:00 AM to 10:00 AM Monday, September 28, 2026")', () => {
    const out = parseEventLabel('Busy 9:00 AM to 10:00 AM Monday, September 28, 2026', 2026);
    assert.equal(out.date, '2026-09-28');
    assert.equal(out.start, '09:00');
    assert.equal(out.end, '10:00');
  });

  test('PM-Zeiten werden korrekt konvertiert', () => {
    const out = parseEventLabel('Review, 2:00 PM to 3:30 PM Tuesday, September 29, 2026', 2026);
    assert.equal(out.date, '2026-09-29');
    assert.equal(out.start, '14:00');
    assert.equal(out.end, '15:30');
  });

  test('Ganztag-Event: start/end null, allDay-Flag gesetzt', () => {
    const out = parseEventLabel('Urlaub, Ganztägig, Montag, 28. September 2026', 2026);
    assert.equal(out.date, '2026-09-28');
    assert.equal(out.start, null);
    assert.equal(out.end, null);
    assert.equal(out.allDay, true);
    assert.equal(out.title, 'Urlaub');
  });

  test('Slot über Mitternacht rollt die Endzeit auf den Folgetag-Anteil', () => {
    const out = parseEventLabel('Nachtbereitschaft, 23:30 bis 00:30 Montag, 28. September', 2026);
    assert.equal(out.start, '23:30');
    assert.equal(out.end, '00:30');
  });

  test('Label ohne erkennbares Datum liefert null', () => {
    assert.equal(parseEventLabel('Besetzt von 09:00 bis 10:00', 2026), null);
    assert.equal(parseEventLabel('', 2026), null);
  });

  test('DE-Label "von 08:00 bis 10:00" (Live-Fail-Variante) wird geparst', () => {
    const out = parseEventLabel('MKK-Termin von 08:00 bis 10:00 Montag, 28. September', 2026);
    assert.equal(out.date, '2026-09-28');
    assert.equal(out.start, '08:00');
    assert.equal(out.end, '10:00');
    assert.equal(out.title, 'MKK-Termin');
  });

  test('DE-Label mit Bindestrich-Zeitfenster ("14:00 - 15:00") wird geparst', () => {
    const out = parseEventLabel('Review, 14:00 - 15:00, Montag, 28. September', 2026);
    assert.equal(out.date, '2026-09-28');
    assert.equal(out.start, '14:00');
    assert.equal(out.end, '15:00');
  });

  test('EN-Label ohne Statuswort ("Team Sync 9:00 to 10:00 Monday, September 28") wird geparst', () => {
    const out = parseEventLabel('Team Sync 9:00 to 10:00 Monday, September 28', 2026);
    assert.equal(out.date, '2026-09-28');
    assert.equal(out.start, '09:00');
    assert.equal(out.end, '10:00');
    assert.equal(out.title, 'Team Sync');
  });

  test('EN-Ganztag-Label ("All day event") wird erkannt', () => {
    const out = parseEventLabel('Offsite, All day event, Monday, September 28, 2026', 2026);
    assert.equal(out.date, '2026-09-28');
    assert.equal(out.start, null);
    assert.equal(out.end, null);
    assert.equal(out.allDay, true);
  });

  test('Tages-Spaltenkopf ("Montag, 28. September") liefert das Datum für den Header-Fallback', () => {
    const out = parseEventLabel('Montag, 28. September', 2026);
    assert.equal(out.date, '2026-09-28');
    assert.equal(out.allDay, true);
  });

  test('EN-Tages-Spaltenkopf ("Monday, September 28") liefert das Datum', () => {
    const out = parseEventLabel('Monday, September 28', 2026);
    assert.equal(out.date, '2026-09-28');
  });
});

describe('classifyEventStatus - Wortlisten-Klassifikation der Event-Labels', () => {
  test('Busy/Besetzt (DE+EN) -> busy', () => {
    assert.equal(classifyEventStatus('Besetzt von 09:00 bis 10:00 Montag, 28. September'), 'busy');
    assert.equal(classifyEventStatus('Busy 9:00 AM to 10:00 AM Monday, September 28'), 'busy');
  });

  test('Tentative/Mit Vorbehalt (DE+EN) -> tentative', () => {
    assert.equal(classifyEventStatus('Mit Vorbehalt, 14:00 bis 15:00, Montag, 28. September'), 'tentative');
    assert.equal(classifyEventStatus('Tentative 14:00 to 15:00 Monday, September 28'), 'tentative');
  });

  test('Free/Frei (DE+EN) -> free', () => {
    assert.equal(classifyEventStatus('Free 14:00 to 15:00 Monday, September 28'), 'free');
    assert.equal(classifyEventStatus('Frei, 14:00 bis 15:00, Montag, 28. September'), 'free');
  });

  test('"Freitag" wird NICHT als free klassifiziert (Wortgrenzen-Pflicht)', () => {
    assert.equal(classifyEventStatus('Freitag, 28. September'), 'busy');
    assert.equal(classifyEventStatus('Termin am Freitag, 14:00 bis 15:00'), 'busy');
  });

  test('Arbeitsort-Banner sind keine Events (non_event)', () => {
    assert.equal(classifyEventStatus('Arbeitsort: Homeoffice, Ganztägig, Montag, 28. September'), 'non_event');
    assert.equal(classifyEventStatus('Working location: Home, All day, Monday, September 28'), 'non_event');
  });

  test('Label ohne Statuswort -> busy (Default für gerenderte Event-Blöcke)', () => {
    assert.equal(classifyEventStatus('MKK-Termin, 14:00 bis 15:00, Montag, 28. September'), 'busy');
    assert.equal(classifyEventStatus(''), 'busy');
  });
});

describe('parseMailRowTimestamp - Datum aus Nachrichtenlisten-Zeilen', () => {
  const now = new Date(2026, 8, 24, 16, 0); // Do, 24.09.2026 16:00

  test('Vollständiges Datum "24.09.2026"', () => {
    const ts = parseMailRowTimestamp('Markus\nAW: Vertrag\n24.09.2026', now);
    assert.equal(new Date(ts).getDate(), 24);
    assert.equal(new Date(ts).getMonth(), 8);
    assert.equal(new Date(ts).getFullYear(), 2026);
  });

  test('Kurzes Datum "21.09." (aktuelles Jahr)', () => {
    const ts = parseMailRowTimestamp('Gabriel\nFrage zum Budget\n21.09.', now);
    assert.equal(new Date(ts).getDate(), 21);
    assert.equal(new Date(ts).getMonth(), 8);
    assert.equal(new Date(ts).getFullYear(), 2026);
  });

  test('"Heute 14:32" liefert heute mit Uhrzeit', () => {
    const ts = parseMailRowTimestamp('David\nEntwurf\nHeute 14:32', now);
    const d = new Date(ts);
    assert.equal(d.getDate(), 24);
    assert.equal(d.getHours(), 14);
    assert.equal(d.getMinutes(), 32);
  });

  test('"Gestern" liefert den Vortag', () => {
    const ts = parseMailRowTimestamp('Kollege\nRückfrage\nGestern', now);
    const d = new Date(ts);
    assert.equal(d.getDate(), 23);
    assert.equal(d.getMonth(), 8);
  });

  test('Wochentag-Kürzel "Mo." am Donnerstag = letzter Montag (21.09.)', () => {
    const ts = parseMailRowTimestamp('Team\nStandup-Ergebnisse\nMo.', now);
    const d = new Date(ts);
    assert.equal(d.getDate(), 21);
    assert.equal(d.getDay(), 1);
  });

  test('Freistehende Uhrzeit in eigener Zeile = heute', () => {
    const ts = parseMailRowTimestamp('David\nKurzlink\n14:05', now);
    const d = new Date(ts);
    assert.equal(d.getDate(), 24);
    assert.equal(d.getHours(), 14);
  });

  test('Uhrzeit im Fließtext ("Meeting um 14:00") wird NICHT als Datum interpretiert', () => {
    assert.equal(parseMailRowTimestamp('Einladung: Meeting um 14:00 im Raum 4', now), null);
  });

  test('Text ohne Datum liefert null', () => {
    assert.equal(parseMailRowTimestamp('Nur ein Betreff ohne Datum', now), null);
    assert.equal(parseMailRowTimestamp('', now), null);
  });
});

describe('parseNumericDayHeaderText - numerische Tages-Spaltenköpfe (Runde-2-Lücke)', () => {
  const now = new Date(2026, 8, 24);

  test('DE-Kurzformat "Montag, 28.9." -> 2026-09-28', () => {
    assert.equal(parseNumericDayHeaderText('Montag, 28.9.', now), '2026-09-28');
  });

  test('Ohne Wochentag "28.9." -> 2026-09-28', () => {
    assert.equal(parseNumericDayHeaderText('28.9.', now), '2026-09-28');
  });

  test('Mit Jahr "28.9.2026" -> 2026-09-28', () => {
    assert.equal(parseNumericDayHeaderText('28.9.2026', now), '2026-09-28');
  });

  test('Wochentag-Kürzel "Di 29.9." -> 2026-09-29', () => {
    assert.equal(parseNumericDayHeaderText('Di 29.9.', now), '2026-09-29');
  });

  test('Umschalter-Text ohne Datum -> null', () => {
    assert.equal(parseNumericDayHeaderText('Work week', now), null);
    assert.equal(parseNumericDayHeaderText('Arbeitswoche', now), null);
  });

  test('Uhrzeit ist kein Datum -> null', () => {
    assert.equal(parseNumericDayHeaderText('14:00', now), null);
  });

  test('Ungültiges Datum (32.13.) -> null', () => {
    assert.equal(parseNumericDayHeaderText('32.13.', now), null);
  });

  test('Leerer/undefinierter Text -> null', () => {
    assert.equal(parseNumericDayHeaderText('', now), null);
    assert.equal(parseNumericDayHeaderText(null, now), null);
  });
});

describe('resolveRenderedWeekFromCandidates - Wochen-Erkennung aus Kandidaten', () => {
  const now = new Date(2026, 8, 24);

  test('Live-Fail Runde 2: Nur Umschalter-Texte ("Work week") -> renderedWeek komplett null', () => {
    // Der Umschalter darf NIE als renderedWeek.text mit null-Daten durchkommen.
    const out = resolveRenderedWeekFromCandidates(
      { headerTexts: ['Work week', 'Posteingang'], dayHeaderTexts: [] }, now
    );
    assert.equal(out.startDate, null);
    assert.equal(out.endDate, null);
    assert.equal(out.text, null);
  });

  test('Wochen-Range-Header mit Monatsnamen wird erkannt', () => {
    const out = resolveRenderedWeekFromCandidates(
      { headerTexts: ['Work week', '28. September – 4. Oktober 2026'], dayHeaderTexts: [] }, now
    );
    assert.equal(out.startDate, '2026-09-28');
    assert.equal(out.endDate, '2026-10-04');
  });

  test('Tages-Spaltenkopf mit Monatsname ("Montag, 28. September") leitet die Woche ab', () => {
    const out = resolveRenderedWeekFromCandidates(
      { headerTexts: ['Work week'], dayHeaderTexts: ['Montag, 28. September'] }, now
    );
    assert.equal(out.startDate, '2026-09-28');
    assert.equal(out.endDate, '2026-10-04');
  });

  test('Numerischer Tages-Spaltenkopf ("Montag 28.9.") leitet die Woche ab (Runde-2-Lücke)', () => {
    const out = resolveRenderedWeekFromCandidates(
      { headerTexts: ['Work week'], dayHeaderTexts: ['Montag 28.9.'] }, now
    );
    assert.equal(out.startDate, '2026-09-28');
    assert.equal(out.endDate, '2026-10-04');
  });

  test('Body-Text-Heuristik-Zeile ("Montag, 28. September") wird als Tageskopf-Kandidat gewertet', () => {
    const out = resolveRenderedWeekFromCandidates(
      { headerTexts: [], dayHeaderTexts: ['Montag, 28. September'] }, now
    );
    assert.equal(out.startDate, '2026-09-28');
    assert.equal(out.endDate, '2026-10-04');
  });

  test('Numerischer Range im Tageskopf-Kandidaten ("28.9. – 4.10.2026") wird erkannt', () => {
    const out = resolveRenderedWeekFromCandidates(
      { headerTexts: [], dayHeaderTexts: ['28.9. – 4.10.2026'] }, now
    );
    assert.equal(out.startDate, '2026-09-28');
    assert.equal(out.endDate, '2026-10-04');
  });

  test('Live-Fail Runde 4: stale All-Day-Event-Label wird NICHT als Header akzeptiert', () => {
    // Das DOM enthielt neben dem echten Range-Header ein stale/hidden Event-Label
    // ("Urlaub … September 23", Woche 39). Dieses darf die Erkennung NICHT kapern.
    const out = resolveRenderedWeekFromCandidates({
      headerTexts: [
        'Urlaub, all day event, Wednesday, September 23, 2026, Busy',
        '28 September – 4 October, 2026, Jump to a specific date or date range.'
      ],
      dayHeaderTexts: []
    }, now);
    assert.equal(out.startDate, '2026-09-28');
    assert.equal(out.endDate, '2026-10-04');
    assert.ok(!out.text.includes('Urlaub'));
  });

  test('Live-Fail Runde 4: Event-Label ALLEIN -> unknown (keine Woche 39 aus Event)', () => {
    const out = resolveRenderedWeekFromCandidates({
      headerTexts: ['Urlaub, all day event, Wednesday, September 23, 2026, Busy'],
      dayHeaderTexts: []
    }, now);
    assert.equal(out.startDate, null);
    assert.equal(out.endDate, null);
    assert.equal(out.text, null);
  });

  test('Live-Fail Runde 4: Zeitfenster-Event-Label wird nicht als Header akzeptiert', () => {
    const out = resolveRenderedWeekFromCandidates({
      headerTexts: ['Stadt Osnabrück, 09:00 to 12:00, Monday, September 28, 2026'],
      dayHeaderTexts: []
    }, now);
    assert.equal(out.startDate, null);
  });

  test('Prev/Next-Button-Ableitung: Week 39/41 -> angezeigte Woche 40', () => {
    const out = resolveRenderedWeekFromCandidates({
      headerTexts: [],
      dayHeaderTexts: [],
      navTexts: [
        'Go to previous week 21–27 September, 2026 (Week 39)',
        'Go to next week 5–11 October, 2026 (Week 41)'
      ]
    }, now);
    assert.equal(out.startDate, '2026-09-28');
    assert.equal(out.endDate, '2026-10-04');
    assert.equal(out.weekNumber, 40);
  });

  test('Prev/Next-Ableitung greift nur als Fallback (Range-Header hat Vorrang)', () => {
    const out = resolveRenderedWeekFromCandidates({
      headerTexts: ['28 September – 4 October, 2026'],
      dayHeaderTexts: [],
      navTexts: ['Go to previous week 21–27 September, 2026 (Week 39)']
    }, now);
    assert.equal(out.startDate, '2026-09-28');
    assert.equal(out.weekNumber, null);
  });

  test('Leere Kandidaten -> unknown (alles null)', () => {
    const out = resolveRenderedWeekFromCandidates({ headerTexts: [], dayHeaderTexts: [] }, now);
    assert.equal(out.startDate, null);
    assert.equal(out.endDate, null);
    assert.equal(out.text, null);
  });

  test('Fehlendes Kandidaten-Objekt -> unknown (defensiver Default)', () => {
    const out = resolveRenderedWeekFromCandidates(undefined, now);
    assert.equal(out.startDate, null);
    assert.equal(out.text, null);
  });
});

describe('isCalendarEventLabel / filterCalendarHeaderCandidates - strikte Header-/Event-Trennung (Runde 4)', () => {
  test('Zeitfenster-Labels sind Events (to/bis/Bindestrich)', () => {
    assert.equal(isCalendarEventLabel('Stadt Osnabrück, 09:00 to 12:00, Monday, September 28, 2026'), true);
    assert.equal(isCalendarEventLabel('Teammeeting, 08:00 bis 09:00, Dienstag, 29. September 2026'), true);
    assert.equal(isCalendarEventLabel('Block, 13:00 - 18:30, Wednesday, September 30, 2026'), true);
  });

  test('Ganztag-Labels sind Events (all day / ganztägig)', () => {
    assert.equal(isCalendarEventLabel('Urlaub, all day event, Wednesday, September 23, 2026, Busy'), true);
    assert.equal(isCalendarEventLabel('Urlaub, Ganztägig, Montag, 28. September 2026'), true);
    assert.equal(isCalendarEventLabel('Feiertag, all-day, 3 October 2026'), true);
  });

  test('Statuswort + Datum ist ein Event (Busy/Tentative/Free/Besetzt/Vorbehalt)', () => {
    assert.equal(isCalendarEventLabel('Review, Busy, September 28, 2026'), true);
    assert.equal(isCalendarEventLabel('Termin, Besetzt, 28. September 2026'), true);
    assert.equal(isCalendarEventLabel('Optional, Tentative, September 28, 2026'), true);
  });

  test('Statuswort OHNE Datum ist KEIN Event (kein Fehlfilter)', () => {
    assert.equal(isCalendarEventLabel('Busy'), false);
    assert.equal(isCalendarEventLabel('Free/Busy settings'), false);
  });

  test('Echter Wochen-Header ist KEIN Event', () => {
    assert.equal(isCalendarEventLabel('28 September – 4 October, 2026, Jump to a specific date or date range.'), false);
    assert.equal(isCalendarEventLabel('28. September – 4. Oktober 2026'), false);
    assert.equal(isCalendarEventLabel('Go to previous week 21–27 September, 2026 (Week 39)'), false);
  });

  test('filterCalendarHeaderCandidates entfernt Event-Labels, behält Header', () => {
    const out = filterCalendarHeaderCandidates([
      'Urlaub, all day event, Wednesday, September 23, 2026, Busy',
      '28 September – 4 October, 2026',
      'Stadt Hörstel, 13:00 to 15:00, Tuesday, September 29, 2026'
    ]);
    assert.deepEqual(out, ['28 September – 4 October, 2026']);
  });

  test('Leere/undefined Eingabe -> leeres Array', () => {
    assert.deepEqual(filterCalendarHeaderCandidates([]), []);
    assert.deepEqual(filterCalendarHeaderCandidates(undefined), []);
  });
});

describe('isVisibleCalendarElement - hidden-Element-Filter (Runde 4, Fixture)', () => {
  test('offsetParent null (hidden) -> unsichtbar', () => {
    assert.equal(isVisibleCalendarElement({ offsetParent: null, position: 'static', rect: { width: 100, height: 20 } }), false);
  });

  test('position:fixed mit offsetParent null ist sichtbar (Spezifikations-Ausnahme)', () => {
    assert.equal(isVisibleCalendarElement({ offsetParent: null, position: 'fixed', rect: { width: 100, height: 20 } }), true);
  });

  test('Null-Fläche (Breite/Höhe 0) -> unsichtbar (virtualisiertes Element)', () => {
    assert.equal(isVisibleCalendarElement({ offsetParent: {}, position: 'static', rect: { width: 0, height: 20 } }), false);
    assert.equal(isVisibleCalendarElement({ offsetParent: {}, position: 'static', rect: { width: 100, height: 0 } }), false);
    assert.equal(isVisibleCalendarElement({ offsetParent: {}, position: 'static', rect: null }), false);
  });

  test('Sichtbares Element mit Fläche -> sichtbar', () => {
    assert.equal(isVisibleCalendarElement({ offsetParent: {}, position: 'static', rect: { width: 100, height: 20 } }), true);
  });

  test('Defensiver Default (kein Argument) -> unsichtbar', () => {
    assert.equal(isVisibleCalendarElement(), false);
  });
});

describe('Prev/Next-Wochen-Ableitung - parseWeekNumberFromNav / parseNavWeekRange / resolveWeekFromNavButtons', () => {
  const now = new Date(2026, 8, 24);

  test('parseWeekNumberFromNav liest "(Week NN)" und "(KW NN)"', () => {
    assert.equal(parseWeekNumberFromNav('Go to previous week 21–27 September, 2026 (Week 39)'), 39);
    assert.equal(parseWeekNumberFromNav('Vorherige Woche (KW 41)'), 41);
    assert.equal(parseWeekNumberFromNav('Go to next week'), null);
    assert.equal(parseWeekNumberFromNav('(Week 99)'), null);
  });

  test('parseNavWeekRange liest "21–27 September, 2026"', () => {
    assert.deepEqual(parseNavWeekRange('Go to previous week 21–27 September, 2026 (Week 39)', now), {
      startDate: '2026-09-21', endDate: '2026-09-27'
    });
    assert.deepEqual(parseNavWeekRange('Go to next week 5–11 October, 2026 (Week 41)', now), {
      startDate: '2026-10-05', endDate: '2026-10-11'
    });
  });

  test('resolveWeekFromNavButtons: Week 39/41 -> Woche 40 (28.09.–04.10.)', () => {
    const out = resolveWeekFromNavButtons([
      'Go to previous week 21–27 September, 2026 (Week 39)',
      'Go to next week 5–11 October, 2026 (Week 41)'
    ], now);
    assert.equal(out.startDate, '2026-09-28');
    assert.equal(out.endDate, '2026-10-04');
    assert.equal(out.weekNumber, 40);
  });

  test('resolveWeekFromNavButtons: nur prev -> Montag = Tag nach prev.endDate', () => {
    const out = resolveWeekFromNavButtons(['Go to previous week 21–27 September, 2026 (Week 39)'], now);
    assert.equal(out.startDate, '2026-09-28');
    assert.equal(out.endDate, '2026-10-04');
    assert.equal(out.weekNumber, 40);
  });

  test('resolveWeekFromNavButtons: nur next -> Sonntag = Tag vor next.startDate', () => {
    const out = resolveWeekFromNavButtons(['Go to next week 5–11 October, 2026 (Week 41)'], now);
    assert.equal(out.startDate, '2026-09-28');
    assert.equal(out.endDate, '2026-10-04');
    assert.equal(out.weekNumber, 40);
  });

  test('resolveWeekFromNavButtons: nur "(Week NN)" ohne Datumsbereich -> ISO-Woche', () => {
    const out = resolveWeekFromNavButtons(['Go to previous week (Week 39)'], now);
    assert.equal(out.weekNumber, 40);
    assert.equal(out.startDate, '2026-09-28');
    assert.equal(out.endDate, '2026-10-04');
  });

  test('resolveWeekFromNavButtons: keine Nav-Texte -> null', () => {
    assert.equal(resolveWeekFromNavButtons([], now), null);
    assert.equal(resolveWeekFromNavButtons(undefined, now), null);
  });

  test('isoWeekMonday: KW 40/2026 beginnt am 28.09.2026', () => {
    const d = isoWeekMonday(40, 2026);
    const iso = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    assert.equal(iso, '2026-09-28');
  });
});

describe('classifyButtonFailure - präzise Abbruch-Ursache (Runde-4-Forderung 3)', () => {
  test('Button gefunden/geklickt, aber wirkungslos -> button-ineffective', () => {
    assert.equal(classifyButtonFailure({ foundAny: true, clickedAny: true }), 'button-ineffective');
    assert.equal(classifyButtonFailure({ foundAny: true, clickedAny: false }), 'button-ineffective');
  });

  test('Gar kein Button gefunden -> no-week-button', () => {
    assert.equal(classifyButtonFailure({ foundAny: false, clickedAny: false }), 'no-week-button');
    assert.equal(classifyButtonFailure(), 'no-week-button');
  });
});

describe('isWorkLocationPseudoSlot - Work-Location/Pseudo-Slot-Filter (Beauty-Fix Runde 2)', () => {
  test('"(ohne Titel)"-Schnipsel werden ausgeschlossen', () => {
    assert.equal(isWorkLocationPseudoSlot('(ohne Titel)'), true);
    assert.equal(isWorkLocationPseudoSlot('(No title)'), true);
    assert.equal(isWorkLocationPseudoSlot('Ohne Titel'), true);
    assert.equal(isWorkLocationPseudoSlot(''), true);
    assert.equal(isWorkLocationPseudoSlot(null), true);
    assert.equal(isWorkLocationPseudoSlot(undefined), true);
  });

  test('Arbeitsort-/Working-Location-Begriffe werden ausgeschlossen', () => {
    assert.equal(isWorkLocationPseudoSlot('Arbeitsort: Homeoffice'), true);
    assert.equal(isWorkLocationPseudoSlot('Working location: Office'), true);
    assert.equal(isWorkLocationPseudoSlot('Homeoffice'), true);
    assert.equal(isWorkLocationPseudoSlot('Home Office'), true);
    assert.equal(isWorkLocationPseudoSlot('Büro'), true);
  });

  test('URL-Schnipsel als Titel werden ausgeschlossen', () => {
    assert.equal(isWorkLocationPseudoSlot('https://teams.microsoft.com/l/meetup-join/xyz'), true);
    assert.equal(isWorkLocationPseudoSlot('http://example.com'), true);
  });

  test('Echte Termine bleiben erhalten', () => {
    assert.equal(isWorkLocationPseudoSlot('Team Meeting'), false);
    assert.equal(isWorkLocationPseudoSlot('FIT-Connect'), false);
    assert.equal(isWorkLocationPseudoSlot('Urlaub'), false);
    assert.equal(isWorkLocationPseudoSlot('Besprechung Bürogebäude 4. OG'), false);
  });
});

describe('dedupeBusySlots - Zeit-Duplikate zusammenführen (Beauty-Fix Runde 2)', () => {
  const mk = (date, start, end, title, allDay = false) => ({ date, start, end, title, allDay, raw: '' });

  test('Identisches Datum+Zeit+Titel nur 1x', () => {
    const out = dedupeBusySlots([
      mk('2026-09-25', '09:00', '10:00', 'Sync'),
      mk('2026-09-25', '09:00', '10:00', 'Sync')
    ]);
    assert.equal(out.length, 1);
    assert.equal(out[0].title, 'Sync');
  });

  test('Überlappende Duplikate gleicher Zeit werden zusammengeführt (echter Titel gewinnt)', () => {
    const out = dedupeBusySlots([
      mk('2026-09-25', '09:00', '09:30', '(ohne Titel)'),
      mk('2026-09-25', '09:00', '09:30', 'FIT-Connect')
    ]);
    assert.equal(out.length, 1);
    assert.equal(out[0].title, 'FIT-Connect');
  });

  test('Verschiedene Zeiten bleiben erhalten', () => {
    const out = dedupeBusySlots([
      mk('2026-09-25', '09:00', '10:00', 'A'),
      mk('2026-09-25', '14:00', '15:00', 'B'),
      mk('2026-09-28', '09:00', '10:00', 'A')
    ]);
    assert.equal(out.length, 3);
  });

  test('Ganztag-Slots am gleichen Tag werden zusammengeführt', () => {
    const out = dedupeBusySlots([
      mk('2026-09-28', null, null, 'Urlaub', true),
      mk('2026-09-28', null, null, 'Urlaub', true)
    ]);
    assert.equal(out.length, 1);
  });

  test('Leere/defekte Slots werden ignoriert', () => {
    const out = dedupeBusySlots([null, mk('2026-09-25', '09:00', '10:00', 'A'), {}]);
    assert.equal(out.length, 1);
  });
});

describe('formatCalendarDiagnostics - Diagnostik-Formatierung für Guards (Runde-3-Forderung)', () => {
  test('Enthält URL, Kandidaten, ARIA-Anzahl und Sample', () => {
    const out = formatCalendarDiagnostics({
      url: 'https://outlook.cloud.microsoft/calendar/view/week?startdate=2026-09-28',
      headerCandidates: ['Work week'],
      dayHeaderCandidates: ['Montag 28.9.'],
      ariaButtonCount: 42,
      ariaLabelSample: ['Team Meeting, 14:00 bis 15:00, Montag, 28. September']
    });
    assert.ok(out.includes('outlook.cloud.microsoft'));
    assert.ok(out.includes('Work week'));
    assert.ok(out.includes('Montag 28.9.'));
    assert.ok(out.includes('42'));
    assert.ok(out.includes('Team Meeting'));
  });

  test('Leere Kandidatenlisten werden als "(keine)" ausgewiesen', () => {
    const out = formatCalendarDiagnostics({
      url: 'https://outlook.office.com/calendar',
      headerCandidates: [],
      dayHeaderCandidates: [],
      ariaButtonCount: 0,
      ariaLabelSample: []
    });
    assert.ok(out.includes('(keine)'));
    assert.ok(out.includes('ARIA-Buttons gesamt: 0'));
  });

  test('Ohne Diagnostik -> Platzhalter', () => {
    assert.equal(formatCalendarDiagnostics(null), '(keine Diagnostik verfügbar)');
    assert.equal(formatCalendarDiagnostics(undefined), '(keine Diagnostik verfügbar)');
  });
});


// ---------------------------------------------------------------------------
// Iteration 4 (Runde-3-Live-Befunde 25.09.): MCP-Timeout -32001 ohne Ergebnis,
// Diagnostik vom Timeout verschluckt, Button-Klick wirkungslos, outlook_status
// falsches activeMailboxMatchesRealm:false im Kalender-Tab.
// ---------------------------------------------------------------------------

describe('FreeBusy-Budget-Konstanten (Runde-4-Forderung A)', () => {
  test('Gesamtbudget sicher unter dem MCP-Client-Timeout, Rationierung konsistent', () => {
    assert.ok(FREEBUSY_TOTAL_BUDGET_MS <= 45000);
    assert.equal(FREEBUSY_MAX_BUTTON_CLICKS, 6);
    assert.equal(FREEBUSY_CLICK_RENDER_WAIT_MS, 4000);
    assert.ok(FREEBUSY_MIN_REMAINING_FOR_ESCALATION_MS < FREEBUSY_TOTAL_BUDGET_MS);
  });
});

describe('phaseBudgetMs - Budget-Rationierung der Phasen (Runde-4-Forderung A)', () => {
  test('Cap greift, wenn Restzeit größer ist', () => {
    assert.equal(phaseBudgetMs(45000, 12000), 12000);
    assert.equal(phaseBudgetMs(45000, 4000), 4000);
  });

  test('Restzeit begrenzt, wenn sie unter dem Cap liegt', () => {
    assert.equal(phaseBudgetMs(5000, 12000), 5000);
  });

  test('Unter der Mindestscheibe wird die Phase geskippt (0) — kein Floor mehr', () => {
    // Runde-3-Root-Cause: Math.max(1000, ...) ließ Phasen nach Ablauf der
    // Deadline weiterlaufen und ließ die Phasensumme das Budget sprengen.
    assert.equal(phaseBudgetMs(1000, 12000), 0);
    assert.equal(phaseBudgetMs(0, 12000), 0);
    assert.equal(phaseBudgetMs(-500, 12000), 0);
  });

  test('Phasen teilen sich EIN Budget: Summe überschreitet das Gesamtbudget nie', () => {
    // Simulation der getFreeBusy-Phaskette (Navigation, URL-Wait, Render-Wait,
    // Retry-Stufen, Klicks) gegen das 45s-Gesamtbudget.
    let remaining = FREEBUSY_TOTAL_BUDGET_MS;
    const caps = [12000, 4000, 12000, 4000, 8000, 6000, 6000, 6000, 6000, 6000, 6000, 6000];
    let allocated = 0;
    for (const cap of caps) {
      const slice = phaseBudgetMs(remaining, cap);
      allocated += slice;
      remaining -= slice;
    }
    assert.ok(allocated <= FREEBUSY_TOTAL_BUDGET_MS);
    assert.equal(remaining, FREEBUSY_TOTAL_BUDGET_MS - allocated);
  });
});

describe('affordableClicks - Klick-Rationierung (Runde-4-Forderung A)', () => {
  const clickCost = 1500 + FREEBUSY_CLICK_RENDER_WAIT_MS; // 5500ms

  test('Volles Budget trägt die maximale Klickanzahl (6)', () => {
    assert.equal(affordableClicks(FREEBUSY_TOTAL_BUDGET_MS, clickCost, FREEBUSY_MAX_BUTTON_CLICKS), 6);
  });

  test('Knappes Budget rationiert die Klicks', () => {
    assert.equal(affordableClicks(20000, clickCost, 6), 3);
    assert.equal(affordableClicks(5499, clickCost, 6), 0);
  });

  test('Kein Restbudget -> 0 Klicks', () => {
    assert.equal(affordableClicks(0, clickCost, 6), 0);
    assert.equal(affordableClicks(-100, clickCost, 6), 0);
  });
});

describe('weekButtonCandidates - Multi-Selektor-Kandidatenliste (Runde-4-Forderung C)', () => {
  test('Nächste Woche: DE/EN aria-label, title und Chevron-Symbol-Kandidaten', () => {
    const cands = weekButtonCandidates(1);
    const sels = cands.map(c => c.selector).join(' | ');
    assert.ok(sels.includes('Nächste Woche'));
    assert.ok(sels.includes('Next week'));
    assert.ok(cands.some(c => c.kind === 'title-de'));
    assert.ok(cands.some(c => c.kind === 'title-en'));
    assert.ok(cands.some(c => c.kind.startsWith('symbol-')));
    assert.ok(cands.every(c => c.selector && c.kind));
  });

  test('Vorherige Woche: DE/EN-Varianten vorhanden', () => {
    const cands = weekButtonCandidates(-1);
    const sels = cands.map(c => c.selector).join(' | ');
    assert.ok(sels.includes('Vorherige Woche'));
    assert.ok(sels.includes('Previous week'));
    assert.ok(cands.some(c => c.kind.startsWith('symbol-')));
  });

  test('Richtungen überschneiden sich nicht (kein Cross-Matching)', () => {
    const next = weekButtonCandidates(1).map(c => c.selector);
    const prev = weekButtonCandidates(-1).map(c => c.selector);
    for (const s of next) assert.ok(!prev.includes(s));
  });
});

describe('evaluateClickVerification - Klick-Verifikation (Runde-4-Forderung C)', () => {
  test('Header-Datum geändert -> changed', () => {
    assert.equal(evaluateClickVerification('2026-09-21', { startDate: '2026-09-28' }), 'changed');
  });

  test('Header-Datum identisch -> unchanged (Klick wirkungslos)', () => {
    assert.equal(evaluateClickVerification('2026-09-28', { startDate: '2026-09-28' }), 'unchanged');
  });

  test('Header vor oder nach dem Klick unlesbar -> unverified', () => {
    assert.equal(evaluateClickVerification(null, { startDate: '2026-09-28' }), 'unverified');
    assert.equal(evaluateClickVerification('2026-09-28', { startDate: null }), 'unverified');
    assert.equal(evaluateClickVerification('2026-09-28', null), 'unverified');
  });

  test('Nur nachweislich wirkungslose Klicks führen zum nächsten Kandidaten', () => {
    assert.equal(shouldTryNextCandidate('unchanged'), true);
    assert.equal(shouldTryNextCandidate('changed'), false);
    assert.equal(shouldTryNextCandidate('unverified'), false);
  });
});

describe('assembleFreeBusyDiagnostics - Diagnostik überlebt den Transport (Runde-4-Forderung B)', () => {
  const early = { url: 'https://outlook.cloud.microsoft/calendar/view/week?startdate=2026-09-28', headerCandidates: ['Work week'], dayHeaderCandidates: [], ariaButtonCount: 3, ariaLabelSample: ['a'] };
  const fresh = { url: 'https://outlook.office.com/calendar/view/week', headerCandidates: ['28. September – 4. Oktober 2026'], dayHeaderCandidates: ['Montag 28.9.'], ariaButtonCount: 42, ariaLabelSample: ['Team Meeting, 14:00 bis 15:00'] };

  test('Frische Diagnostik gewinnt, früher Snapshot bleibt als urlAfterNavigation erhalten', () => {
    const out = assembleFreeBusyDiagnostics({ fresh, early, abortReason: 'week-mismatch' });
    assert.equal(out.url, fresh.url);
    assert.equal(out.urlAfterNavigation, early.url);
    assert.equal(out.abortReason, 'week-mismatch');
    assert.deepEqual(out.headerCandidates, fresh.headerCandidates);
  });

  test('Ohne frische Diagnostik (Budget) trägt der frühe Snapshot die Diagnostik', () => {
    const out = assembleFreeBusyDiagnostics({ fresh: null, early, abortReason: 'budget' });
    assert.equal(out.url, early.url);
    assert.equal(out.urlAfterNavigation, early.url);
    assert.equal(out.abortReason, 'budget');
  });

  test('Ohne beides: leere, aber strukturierte Diagnostik (kein Crash)', () => {
    const out = assembleFreeBusyDiagnostics({});
    assert.equal(out.url, '');
    assert.deepEqual(out.headerCandidates, []);
    assert.deepEqual(out.weekNavCandidates, []);
    assert.deepEqual(out.weekButtonProbe, []);
    assert.equal(out.abortReason, 'none');
  });

  test('Klick-Protokoll und Budget-Report werden eingebettet', () => {
    const probe = [{ selector: '[aria-label*="Next week" i]', kind: 'aria-label-en', found: true, clicked: true, verified: 'changed' }];
    const budget = { totalMs: 45000, usedMs: 12000, remainingMs: 33000 };
    const out = assembleFreeBusyDiagnostics({ fresh, early, weekButtonProbe: probe, abortReason: 'none', budget });
    assert.deepEqual(out.weekButtonProbe, probe);
    assert.deepEqual(out.budget, budget);
  });
});

describe('resolveMailboxVerification - Verifikations-Quellen (Runde-4-Forderung D)', () => {
  test('Enum enthält genau die vier dokumentierten Quellen', () => {
    assert.deepEqual(MAILBOX_VERIFICATION_SOURCES, ['mail-view', 'calendar-header', 'tab-title', 'unknown']);
  });

  test('Passender Hint -> true', () => {
    const v = resolveMailboxVerification(['David Hallmann (david@adesso.de)'], 'adesso.de');
    assert.equal(v.activeMailboxMatchesRealm, true);
    assert.ok(v.matchingHint.includes('adesso.de'));
  });

  test('Hints ohne Passung -> false (nachweislich falsche Mailbox)', () => {
    const v = resolveMailboxVerification(['david@d-velop.de'], 'adesso.de');
    assert.equal(v.activeMailboxMatchesRealm, false);
  });

  test('Keine Hints (Kalender-Tab-Fall) -> null statt false (Kollateral-Bug Runde 3)', () => {
    const v = resolveMailboxVerification([], 'adesso.de');
    assert.equal(v.activeMailboxMatchesRealm, null);
    assert.equal(v.matchingHint, null);
  });

  test('Ohne Realm -> null (nichts zu matchen)', () => {
    assert.equal(resolveMailboxVerification(['david@adesso.de'], null).activeMailboxMatchesRealm, null);
  });
});

// ---------------------------------------------------------------------------
// Iteration 6 (Live-Befund Runde 5 vom 25.09.): Slot-Extraktion lieferte 0,
// obwohl dieselben Events in Runde 4 als ARIA-Labels im DOM nachweisbar waren.
// ---------------------------------------------------------------------------

describe('parseEventLabel - Live-Labels Runde 5 (OWA-Realformate)', () => {
  const LIVE_LABELS = [
    'Stadt Osnabrück I Portalintegration OpenRathaus, 09:00 to 12:00, Monday, September 28, 2026, Microsoft Teams-Besprechung',
    'Austausch Mathis:David:Alex, 14:00 to 15:00, Monday, September 28, 2026, Microsoft Teams-Besprechung, By Schnieders, Ale',
    'Teammeeting Team 1 - Local Administration, 08:00 to 09:00, Tuesday, September 29, 2026, Microsoft Teams Meeting, By Schn…',
    "Stadt Hörstel: CMIS Übertragungsfehler 'Field repo[355]', 13:00 to 15:00, Tuesday, September 29, 2026, https://video.hoe…",
    'Block, 13:00 to 18:30, Wednesday, September 30, 2026, Busy, Recurring event',
    'Urlaub, all day event, Wednesday, September 30, 2026, Busy'
  ];

  test('alle 6 Live-Labels parsen zu korrekten Slots', () => {
    const slots = LIVE_LABELS.map((l) => parseEventLabel(l, 2026));
    assert.equal(slots.length, 6);
    for (const s of slots) assert.ok(s, 'jedes Live-Label muss parsen');

    assert.deepEqual(
      { date: slots[0].date, start: slots[0].start, end: slots[0].end, title: slots[0].title, allDay: slots[0].allDay },
      { date: '2026-09-28', start: '09:00', end: '12:00', title: 'Stadt Osnabrück I Portalintegration OpenRathaus', allDay: false }
    );
    // Titel mit Doppelpunkten bleibt intakt; "By …"-Abschnitt wird abgeschnitten
    assert.equal(slots[1].title, 'Austausch Mathis:David:Alex');
    assert.equal(slots[1].start, '14:00');
    assert.equal(slots[2].title, 'Teammeeting Team 1 - Local Administration');
    // Titel mit Komma + Doppelpunkt + Klammern bleibt intakt
    assert.equal(slots[3].title, "Stadt Hörstel: CMIS Übertragungsfehler 'Field repo[355]'");
    assert.equal(slots[3].date, '2026-09-29');
    // "Busy, Recurring event" wird als Folge-Attribut abgeschnitten
    assert.equal(slots[4].title, 'Block');
    assert.equal(slots[4].end, '18:30');
    // Ganztag: allDay, start/end null, Titel vor dem Marker
    assert.equal(slots[5].allDay, true);
    assert.equal(slots[5].start, null);
    assert.equal(slots[5].end, null);
    assert.equal(slots[5].title, 'Urlaub');
    assert.equal(slots[5].date, '2026-09-30');
  });
});

describe('partitionSlotsByRange - Range-Filter [start, end) (Runde-5-Forderung)', () => {
  const rangeStart = new Date('2026-09-28T00:00:00').getTime();
  const rangeEnd = new Date('2026-10-02T23:59:59').getTime();

  test('allDay am 30.09. liegt im Range -> behalten', () => {
    const allDay = { date: '2026-09-30', start: null, end: null, allDay: true, title: 'Urlaub', status: 'busy' };
    const { kept, rejected } = partitionSlotsByRange([allDay], rangeStart, rangeEnd);
    assert.equal(kept.length, 1);
    assert.equal(rejected.length, 0);
  });

  test('Slot vor dem Range -> date-out-of-range', () => {
    const before = { date: '2026-09-25', start: '09:00', end: '10:00', allDay: false, title: 'Alt', status: 'busy' };
    const { kept, rejected } = partitionSlotsByRange([before], rangeStart, rangeEnd);
    assert.equal(kept.length, 0);
    assert.equal(rejected[0].reason, 'date-out-of-range');
  });

  test('freier Slot -> free (keine Belegung)', () => {
    const free = { date: '2026-09-29', start: '09:00', end: '10:00', allDay: false, title: 'Frei', status: 'free' };
    const { kept, rejected } = partitionSlotsByRange([free], rangeStart, rangeEnd);
    assert.equal(kept.length, 0);
    assert.equal(rejected[0].reason, 'free');
  });

  test('überlappender Slot am Rand -> behalten', () => {
    const edge = { date: '2026-10-02', start: '22:00', end: '23:00', allDay: false, title: 'Spät', status: 'busy' };
    const { kept } = partitionSlotsByRange([edge], rangeStart, rangeEnd);
    assert.equal(kept.length, 1);
  });
});

describe('slotExtractionWarning + Slot-Extraktions-Diagnostik (Runde-5-Kernlücke)', () => {
  test('0 Slots trotz Event-Elementen -> Widerspruchs-Warnung', () => {
    const w = slotExtractionWarning(6, 0);
    assert.ok(w && /Widerspruch/.test(w));
    assert.ok(/6 Event-Element/.test(w));
  });

  test('kein Widerspruch bei 0 Events oder >0 Slots', () => {
    assert.equal(slotExtractionWarning(0, 0), null);
    assert.equal(slotExtractionWarning(6, 3), null);
  });

  test('assembleFreeBusyDiagnostics übernimmt die Slot-Extraktions-Felder', () => {
    const d = assembleFreeBusyDiagnostics({
      eventElementsFound: 6,
      selectorUsed: '[role="button"][aria-label]',
      rawEventLabels: ['Block, 13:00 to 18:30, Wednesday, September 30, 2026, Busy'],
      rejectionReasons: [{ label: 'x', reason: 'date-out-of-range' }]
    });
    assert.equal(d.eventElementsFound, 6);
    assert.equal(d.selectorUsed, '[role="button"][aria-label]');
    assert.equal(d.rawEventLabels.length, 1);
    assert.equal(d.rejectionReasons[0].reason, 'date-out-of-range');
  });

  test('ohne Slot-Felder bleiben sie weg (Rückwärtskompatibilität)', () => {
    const d = assembleFreeBusyDiagnostics({ abortReason: 'none' });
    assert.equal('eventElementsFound' in d, false);
    assert.equal('selectorUsed' in d, false);
    assert.equal('rejectionReasons' in d, false);
  });
});

describe('collectCalendarEventLabels - gemeinsamer Kollektionspfad & Sichtbarkeits-Toleranz (Runde 5)', () => {
  const fakeEl = (label, { width = 120, height = 24, offsetParent = null } = {}) => ({
    getAttribute: (n) => (n === 'aria-label' ? label : null),
    getBoundingClientRect: () => ({ width, height }),
    offsetParent
  });

  const withFakeDocument = async (doc, fn) => {
    const prev = globalThis.document;
    globalThis.document = doc;
    try { return await fn(); } finally { globalThis.document = prev; }
  };

  // Iteration 7: Der Kollektor reicht seinen Pattern-Satz + Selektor-Memory als
  // evaluate-Argument in die Browser-Callback — der Fake muss es durchreichen.
  const fakePage = { evaluate: async (fn, arg) => fn(arg) };

  test('offsetParent-null Event mit rect>0 wird toleriert (kein Hard-K.O.)', async () => {
    const doc = {
      querySelectorAll: (sel) => (sel === '[role="button"][aria-label]'
        ? [fakeEl('Block, 13:00 to 18:30, Wednesday, September 30, 2026, Busy', { offsetParent: null, width: 120, height: 24 })]
        : [])
    };
    const res = await withFakeDocument(doc, () => collectCalendarEventLabels(fakePage));
    assert.equal(res.eventElementsFound, 1);
    assert.equal(res.hiddenCount, 0); // rect>0 -> sichtbar, obwohl offsetParent null
    assert.equal(res.selectorUsed, '[role="button"][aria-label]');
    assert.equal(res.labels.length, 1);
  });

  test('Fallback auf breiteren aria-label-Scan, wenn der Button-Selektor leer ist', async () => {
    const doc = {
      querySelectorAll: (sel) => (sel === '[aria-label]'
        ? [fakeEl('Urlaub, all day event, Wednesday, September 30, 2026, Busy')]
        : [])
    };
    const res = await withFakeDocument(doc, () => collectCalendarEventLabels(fakePage));
    assert.equal(res.eventElementsFound, 1);
    assert.equal(res.selectorUsed, '[aria-label]');
  });

  test('Header-artiges Label (Datum ohne Zeit) -> headerLikeLabels, kein Event', async () => {
    const doc = {
      querySelectorAll: (sel) => (sel === '[role="button"][aria-label]'
        ? [fakeEl('28 September – 4 October, 2026, Jump to a specific date or date range.')]
        : [])
    };
    const res = await withFakeDocument(doc, () => collectCalendarEventLabels(fakePage));
    assert.equal(res.eventElementsFound, 0);
    assert.equal(res.headerLikeLabels.length, 1);
  });

  test('Arbeitsort-Banner wird als Nicht-Event verworfen', async () => {
    const doc = {
      querySelectorAll: (sel) => (sel === '[role="button"][aria-label]'
        ? [fakeEl('Arbeitsort, 09:00 to 17:00, Monday, September 28, 2026')]
        : [])
    };
    const res = await withFakeDocument(doc, () => collectCalendarEventLabels(fakePage));
    assert.equal(res.eventElementsFound, 0);
  });

  // ---------------------------------------------------------------------
  // Iteration 7 (Live-Befund Runde 6): der Uhrzeit-Indikator war der Grund
  // fuer eventElementsFound: 1 bei real leerem Grid.
  // ---------------------------------------------------------------------
  test('Uhrzeit-Indikator "current time" ist KEIN Event (Runde-6-Fehltreffer)', async () => {
    const doc = {
      querySelectorAll: () => [fakeEl('calendar view, current time: Fri 22:31')]
    };
    const res = await withFakeDocument(doc, () => collectCalendarEventLabels(fakePage));
    assert.equal(res.eventElementsFound, 0);
    assert.equal(res.labels.length, 0);
    assert.equal(res.nonEventLabels.length, 1);
    assert.equal(res.nonEventLabels[0].reason, 'clock-indicator');
  });

  test('Doppelte Event-Labels werden dedupliziert (eventElementsFound bleibt ehrlich)', async () => {
    const label = 'Block, 13:00 to 18:30, Wednesday, September 30, 2026, Busy';
    const doc = {
      querySelectorAll: (sel) => (sel === '[role="button"][aria-label]'
        ? [fakeEl(label), fakeEl(label), fakeEl(label)]
        : [])
    };
    const res = await withFakeDocument(doc, () => collectCalendarEventLabels(fakePage));
    assert.equal(res.eventElementsFound, 1);
    assert.equal(res.labels.length, 1);
  });

  test('Selektor-Memory: gemerkter Selektor wird zuerst probiert, Fallback danach', async () => {
    const calls = [];
    const doc = {
      querySelectorAll: (sel) => {
        calls.push(sel);
        return sel === '[aria-label]'
          ? [fakeEl('Urlaub, all day event, Wednesday, September 30, 2026, Busy')]
          : [];
      }
    };
    const res = await withFakeDocument(doc, () => collectCalendarEventLabels(fakePage, { preferSelector: '[role="button"][aria-label]' }));
    assert.equal(res.eventElementsFound, 1);
    assert.equal(res.selectorUsed, '[aria-label]');
    // Versuchsreihenfolge der LEITER (der Census-Vorlauf davor ist kein
    // Quellen-Versuch): gemerkter Selektor vor dem breiten Fallback.
    assert.deepEqual(res.sourceTried.slice(0, 2).map((s) => s.selector), ['[role="button"][aria-label]', '[aria-label]']);
  });
});

// ---------------------------------------------------------------------------
// Iteration 7 (Live-Befund Runde 6 vom 28.09.2026): Defekt 1 — die
// Event-Wartephase nahm den Zeit-Indikator "calendar view, current time: Fri
// 22:31" als gerendertes Grid und brach nach ~4s "erfolgreich" ab (usedMs 4141
// bei remainingMs 40859), waehrend das Grid erst nach ~17s rendert.
// ---------------------------------------------------------------------------

describe('isEventGridLabel - strenger Grid-Fertig-Beweis (Defekt 1)', () => {
  test('echte Zeitfenster-Labels gelten als Grid-Beweis', () => {
    assert.equal(isEventGridLabel('Stadt Osnabrück I Portalintegration OpenRathaus, 09:00 to 12:00, Monday, September 28, 2026'), true);
    assert.equal(isEventGridLabel('Besetzt von 09:00 bis 10:00 Montag, 28. September'), true);
    assert.equal(isEventGridLabel('Block, 13:00 - 18:30, Wednesday, September 30, 2026'), true);
    assert.equal(isEventGridLabel('Block, 9:00 AM - 10:30 AM, Wednesday'), true);
  });

  test('Ganztag-Labels gelten als Grid-Beweis', () => {
    assert.equal(isEventGridLabel('Urlaub, all day event, Wednesday, September 30, 2026, Busy'), true);
    assert.equal(isEventGridLabel('Urlaub, Ganztägig, Montag, 28. September 2026'), true);
    assert.equal(isEventGridLabel('Homeoffice, all-day, Tuesday'), true);
  });

  test('der Runde-6-Fehltreffer wird abgelehnt', () => {
    // 22:31 matcht /\\d{1,2}:\\d{2}/ — ist aber KEIN Zeitfenster.
    assert.equal(isEventGridLabel('calendar view, current time: Fri 22:31'), false);
  });

  test('Work-Location bleibt ausgeschlossen', () => {
    assert.equal(isEventGridLabel('Arbeitsort, 09:00 to 17:00, Monday, September 28, 2026'), false);
    assert.equal(isEventGridLabel('Working location, 08:00 to 16:00, Monday'), false);
  });

  test('Header/Datum ohne Zeitfenster ist kein Beweis, leere Eingaben ebenso', () => {
    assert.equal(isEventGridLabel('28 September – 4 October, 2026, Jump to a specific date or date range.'), false);
    assert.equal(isEventGridLabel('Work week'), false);
    assert.equal(isEventGridLabel(''), false);
    assert.equal(isEventGridLabel(null), false);
    assert.equal(isEventGridLabel(undefined), false);
  });

  test('alle 6 Live-Labels aus Runde 5/6 gelten als Grid-Beweis', () => {
    for (const l of [
      'Stadt Osnabrück I Portalintegration OpenRathaus, 09:00 to 12:00, Monday, September 28, 2026',
      'Austausch Mathis:David:Alex, 14:00 to 15:00, Monday, September 28, 2026',
      'Teammeeting Team 1 - Local Administration, 08:00 to 09:00, Tuesday, September 29, 2026',
      "Stadt Hörstel: CMIS Übertragungsfehler, 13:00 to 15:00, Tuesday, September 29, 2026",
      'Block, 13:00 to 18:30, Wednesday, September 30, 2026, Busy',
      'Urlaub, all day event, Wednesday, September 30, 2026, Busy'
    ]) {
      assert.equal(isEventGridLabel(l), true, l);
    }
  });
});

describe('EVENT_LABEL_PATTERNS - serialisierbarer Pattern-Satz', () => {
  test('alle Patterns sind Strings (die evaluate-Grenze ist serialisierbar)', () => {
    for (const [key, value] of Object.entries(EVENT_LABEL_PATTERNS)) {
      assert.equal(typeof value, 'string', `${key} muss ein String sein`);
    }
    assert.doesNotThrow(() => JSON.parse(JSON.stringify(EVENT_LABEL_PATTERNS)));
  });

  test('jedes Pattern laesst sich mit Flag i kompilieren', () => {
    for (const [key, value] of Object.entries(EVENT_LABEL_PATTERNS)) {
      assert.doesNotThrow(() => new RegExp(value, 'i'), `${key} kompiliert nicht`);
    }
  });
});

describe('eventWaitBudgetMs - Budget der Event-Wartephase (Defekt 1)', () => {
  test('Konstanten: Cap 20s, Reserve 5s, Poll 500ms', () => {
    assert.equal(FREEBUSY_EVENT_WAIT_MAX_MS, 20000);
    assert.equal(FREEBUSY_EVENT_WAIT_RESERVE_MS, 5000);
    assert.equal(FREEBUSY_EVENT_WAIT_POLL_MS, 500);
    // Die Phase muss in das Gesamtbudget passen (Cap + Reserve <= 45s).
    assert.ok(FREEBUSY_EVENT_WAIT_MAX_MS + FREEBUSY_EVENT_WAIT_RESERVE_MS <= FREEBUSY_TOTAL_BUDGET_MS);
  });

  test('viel Restzeit -> Cap greift (Runde 6: remainingMs 40859 -> 20000)', () => {
    assert.equal(eventWaitBudgetMs(40859), 20000);
    assert.equal(eventWaitBudgetMs(45000), 20000);
    assert.equal(eventWaitBudgetMs(100000), 20000);
  });

  test('wenig Restzeit -> Reserve wird abgezogen', () => {
    assert.equal(eventWaitBudgetMs(24000), 19000);
    assert.equal(eventWaitBudgetMs(6000), 1000);
  });

  test('zu wenig Restzeit -> 0 (Wartephase wird uebersprungen, nie negativ)', () => {
    assert.equal(eventWaitBudgetMs(5000), 0);
    assert.equal(eventWaitBudgetMs(3000), 0);
    assert.equal(eventWaitBudgetMs(0), 0);
    assert.equal(eventWaitBudgetMs(-100), 0);
    assert.equal(eventWaitBudgetMs(undefined), 0);
    assert.equal(eventWaitBudgetMs(null), 0);
  });

  test('Overrides werden respektiert', () => {
    assert.equal(eventWaitBudgetMs(40000, { maxMs: 5000 }), 5000);
    assert.equal(eventWaitBudgetMs(9000, { reserveMs: 1000 }), 8000);
  });
});

describe('pollUntilEventGrid - Wartephase pollt bis echte Events da sind (Defekt 1)', () => {
  const fakeClock = () => {
    let t = 0;
    return { now: () => t, sleep: async (ms) => { t += ms; } };
  };

  test('Fixture: Events erscheinen erst im 3. Poll -> found nach 1000ms', async () => {
    const clock = fakeClock();
    let polls = 0;
    const res = await pollUntilEventGrid({
      hasEvents: async () => { polls++; return polls >= 3; },
      budgetMs: 20000,
      pollMs: 500,
      now: clock.now,
      sleep: clock.sleep
    });
    assert.equal(res.found, true);
    assert.equal(res.outcome, 'found');
    assert.equal(res.polls, 3);
    assert.equal(res.waitedMs, 1000);
  });

  test('sofort vorhandene Events -> 1 Poll, 0ms Wartezeit', async () => {
    const clock = fakeClock();
    const res = await pollUntilEventGrid({
      hasEvents: async () => true,
      budgetMs: 20000, pollMs: 500, now: clock.now, sleep: clock.sleep
    });
    assert.equal(res.found, true);
    assert.equal(res.outcome, 'found');
    assert.equal(res.polls, 1);
    assert.equal(res.waitedMs, 0);
  });

  test('Budget-Cap: nie laenger als budgetMs, dann budget-exhausted', async () => {
    const clock = fakeClock();
    const res = await pollUntilEventGrid({
      hasEvents: async () => false,
      budgetMs: eventWaitBudgetMs(40859), // 20000
      pollMs: 500, now: clock.now, sleep: clock.sleep
    });
    assert.equal(res.found, false);
    assert.equal(res.outcome, 'budget-exhausted');
    assert.ok(res.waitedMs <= 20000, `waitedMs ${res.waitedMs} darf das Budget nicht ueberziehen`);
    assert.ok(res.waitedMs > 19000);
    assert.equal(res.polls, 40); // 20000/500 Polls, letzter Poll ohne Ueberziehung
  });

  test('Budget 0 -> sofort budget-exhausted, kein Poll', async () => {
    const clock = fakeClock();
    let polls = 0;
    const res = await pollUntilEventGrid({
      hasEvents: async () => { polls++; return true; },
      budgetMs: 0, pollMs: 500, now: clock.now, sleep: clock.sleep
    });
    assert.equal(res.found, false);
    assert.equal(res.outcome, 'budget-exhausted');
    assert.equal(res.polls, 0);
    assert.equal(polls, 0);
  });

  test('fehlendes hasEvents crasht nicht (defensiv)', async () => {
    const res = await pollUntilEventGrid({ budgetMs: 20000 });
    assert.equal(res.found, false);
    assert.equal(res.outcome, 'budget-exhausted');
  });

  test('Ausgaenge sind genau die dokumentierten Werte', () => {
    // Defekt 3: 'not-found' kam hinzu — 'found' ist nur noch bei tatsaechlich
    // gelesenen Event-Knoten erlaubt.
    assert.deepEqual(EVENT_WAIT_OUTCOMES, ['found', 'not-found', 'budget-exhausted', 'not-rendered']);
  });
});

// ---------------------------------------------------------------------------
// Iteration 7, Defekt 2: latenter Wochen-Skip. Runde 6 protokollierte einen
// "Next week"-Klick, obwohl der Header die Zielwoche bereits bestaetigt hatte.
// ---------------------------------------------------------------------------

describe('canClickSameDirection - Klick-Sperre pro Richtung (Defekt 2)', () => {
  test('vor dem ersten Klick ist jede Richtung erlaubt', () => {
    assert.equal(canClickSameDirection({ lastClickDirection: null, direction: 1, weekConfirmedAfterClick: false }), true);
    assert.equal(canClickSameDirection({ lastClickDirection: -1, direction: 1, weekConfirmedAfterClick: false }), true);
  });

  test('gleiche Richtung OHNE erneute Wochen-Bestaetigung -> blockiert (verhindert KW-Skip)', () => {
    assert.equal(canClickSameDirection({ lastClickDirection: 1, direction: 1, weekConfirmedAfterClick: false }), false);
    assert.equal(canClickSameDirection({ lastClickDirection: -1, direction: -1, weekConfirmedAfterClick: false }), false);
  });

  test('gleiche Richtung NACH erneuter Wochen-Bestaetigung -> erlaubt', () => {
    assert.equal(canClickSameDirection({ lastClickDirection: 1, direction: 1, weekConfirmedAfterClick: true }), true);
  });

  test('Gegenrichtung bleibt erlaubt — der Rueckweg ist vorbereitet', () => {
    assert.equal(canClickSameDirection({ lastClickDirection: 1, direction: -1, weekConfirmedAfterClick: false }), true);
    assert.equal(canClickSameDirection({ lastClickDirection: -1, direction: 1, weekConfirmedAfterClick: false }), true);
  });

  test('Vorzeichen wird normalisiert (3 und 1 sind dieselbe Richtung, 2 und -1 nicht)', () => {
    // Gleiches Vorzeichen trotz anderer Betraege -> gleiche Richtung -> ohne
    // Wochen-Bestaetigung blockiert.
    assert.equal(canClickSameDirection({ lastClickDirection: 3, direction: 1, weekConfirmedAfterClick: false }), false);
    // Umgekehrtes Vorzeichen -> Rueckweg, erlaubt.
    assert.equal(canClickSameDirection({ lastClickDirection: 2, direction: -1, weekConfirmedAfterClick: false }), true);
  });
});

describe('weekProbeSkipEntry - Skip-Protokoll bei verifizierter Zielwoche (Defekt 2)', () => {
  test('protokolliert explizit den Verzicht auf Probe und Klick', () => {
    const e = weekProbeSkipEntry('2026-09-28');
    assert.equal(e.skipped, 'target-week-verified');
    assert.equal(e.renderedWeek, '2026-09-28');
    assert.equal(e.clicked, false);
    // Kein Selektor, kein found -> es wurde tatsaechlich nicht geprobt/geklickt.
    assert.equal('selector' in e, false);
    assert.equal('found' in e, false);
    assert.equal('verified' in e, false);
  });

  test('toleriert fehlendes renderndes Wochen-Datum', () => {
    const e = weekProbeSkipEntry(null);
    assert.equal(e.skipped, 'target-week-verified');
    assert.equal(e.renderedWeek, null);
    assert.equal(e.clicked, false);
  });

  test('der Skip-Eintrag beendet die Klick-Sequenz (kein Folgeeintrag noetig)', () => {
    const probe = [weekProbeSkipEntry('2026-09-28')];
    assert.equal(probe.length, 1);
    assert.equal(probe.filter((p) => p.clicked === true).length, 0);
  });
});

describe('eventWaitWarning - Warnung bei erschoepfter Wartephase', () => {
  test('found -> keine Warnung', () => {
    assert.equal(eventWaitWarning({ outcome: 'found', waitedMs: 1200, budgetMs: 20000 }), null);
  });

  test('budget-exhausted -> erklaert, dass 0 Slots ein noch nicht gerendertes Grid sein koennen', () => {
    const w = eventWaitWarning({ outcome: 'budget-exhausted', waitedMs: 19500, budgetMs: 20000 });
    assert.ok(w);
    assert.ok(/19500ms/.test(w));
    assert.ok(/20000ms/.test(w));
    assert.ok(/noch nicht gerendertes Grid/.test(w));
  });

  test('not-rendered -> eigener Hinweistext', () => {
    const w = eventWaitWarning({ outcome: 'not-rendered', waitedMs: 0, budgetMs: 0 });
    assert.ok(/nicht gerendert/.test(w));
  });

  test('unbekannter Ausgang -> keine Warnung', () => {
    assert.equal(eventWaitWarning({ outcome: 'irgendwas' }), null);
  });
});

describe('assembleFreeBusyDiagnostics - Event-Wartephase in der Diagnostik (Iteration 7)', () => {
  test('waitedForEventsMs und eventWaitOutcome werden uebernommen', () => {
    const d = assembleFreeBusyDiagnostics({
      abortReason: 'none',
      eventElementsFound: 0,
      selectorUsed: '[aria-label]',
      rawEventLabels: [],
      rejectionReasons: [{ label: 'calendar view, current time: Fri 22:31', reason: 'clock-indicator' }],
      waitedForEventsMs: 19500,
      eventWaitOutcome: 'budget-exhausted'
    });
    assert.equal(d.waitedForEventsMs, 19500);
    assert.equal(d.eventWaitOutcome, 'budget-exhausted');
    assert.equal(d.rejectionReasons[0].reason, 'clock-indicator');
  });

  test('ohne die neuen Felder bleiben sie weg (Rueckwaertskompatibilitaet)', () => {
    const d = assembleFreeBusyDiagnostics({ abortReason: 'none' });
    assert.equal('waitedForEventsMs' in d, false);
    assert.equal('eventWaitOutcome' in d, false);
  });
});

// ===========================================================================
// Defekt-Fix KW 40 (Live-Smoke-Test 2026-09-25, Commit 0c4cfee):
//   Defekt 1  Selektor `[aria-label]` trifft NUR Chrome-/Nav-Knoten
//             (ariaButtonCount=9, 12 Labels, alle header/nav/clock).
//             Es wird eine echte Grid-/Struktur-Suche noetig.
//   Defekt 2  dataQuality:'verified' bei eventElementsFound=0 = Falschaussage.
//   Defekt 3  eventWaitOutcome:'found' bei waitedForEventsMs:0 = Widerspruch.
//
// Messbare Ziele dieser Tests:
//   - eine LEITER unabhaengiger Event-Quellen statt eines geratenen Selektors
//   - Label-Ableitung auch OHNE aria-label (title / data-* / Text+Geometrie)
//   - offsetParent===null ist kein Hard-K.O., Sichtbarkeit ueber boundingRect
//   - ehrliche dataQuality-Werte und an die BEobachtUNG gekoppelter Wait-Ausgang
// ===========================================================================

describe('CALENDAR_EVENT_SOURCES - Quellen-Leiter statt geratenem Selektor (Defekt 1)', () => {
  test('die Leiter enthaelt mindestens eine Quelle OHNE aria-label', () => {
    // Runde 7 hat bewiesen: [aria-label] trifft im Tenant KEINE Event-Chips.
    // Eine Leiter, die nur aria-label-Varianten kennt, waere blind.
    const nonAria = CALENDAR_EVENT_SOURCES.filter((s) => s.source !== 'aria-label');
    assert.ok(nonAria.length >= 2, 'es braucht mindestens 2 Quellen jenseits von aria-label');
  });

  test('die historisch bewaesste Quelle steht weiterhin vorn (Rueckwaertskompatibilitaet)', () => {
    assert.equal(CALENDAR_EVENT_SOURCES[0].selector, '[role="button"][aria-label]');
    assert.equal(CALENDAR_EVENT_SOURCES[0].source, 'aria-label');
  });

  test('Struktur-Quellen (selector null) sind als solche markiert', () => {
    for (const s of CALENDAR_EVENT_SOURCES) {
      if (s.selector === null) assert.equal(typeof s.scope, 'string', `${s.id} braucht einen scope`);
      else assert.equal(typeof s.selector, 'string');
      assert.equal(typeof s.id, 'string');
      assert.equal(typeof s.source, 'string');
    }
  });

  test('jede Quellen-ID ist eindeutig (Diagnostik muss zuordenbar bleiben)', () => {
    const ids = CALENDAR_EVENT_SOURCES.map((s) => s.id);
    assert.equal(new Set(ids).size, ids.length);
  });

  test('eventSourcePlan stellt den gemerkten Selektor voran (Selektor-Memory)', () => {
    const plan = eventSourcePlan('[role="gridcell"][aria-label]');
    assert.equal(plan[0].id, 'remembered');
    assert.equal(plan[0].selector, '[role="gridcell"][aria-label]');
    // danach die regulaere Leiter, ohne Duplikate
    const rest = plan.slice(1).map((s) => s.selector);
    assert.equal(rest.includes('[role="gridcell"][aria-label]'), false);
    assert.equal(rest[0], '[role="button"][aria-label]');
  });

  test('ohne Memory ist die Leiter unveraendert', () => {
    assert.equal(eventSourcePlan(null)[0].id, 'aria-role-button');
    assert.equal(eventSourcePlan(undefined).length, CALENDAR_EVENT_SOURCES.length);
  });
});

describe('normalizeClockTime - Uhrzeit-Raster lesbar machen (Defekt 1, Geometrie)', () => {
  test('24h-Format', () => {
    assert.equal(normalizeClockTime('08:00'), '08:00');
    assert.equal(normalizeClockTime('8:00'), '08:00');
    assert.equal(normalizeClockTime('17:30'), '17:30');
  });

  test('12h-Format mit AM/PM', () => {
    assert.equal(normalizeClockTime('8 AM'), '08:00');
    assert.equal(normalizeClockTime('8:30 am'), '08:30');
    assert.equal(normalizeClockTime('1 PM'), '13:00');
    assert.equal(normalizeClockTime('12 PM'), '12:00');
    assert.equal(normalizeClockTime('12 AM'), '00:00');
  });

  test('Text drumherum wird toleriert', () => {
    assert.equal(normalizeClockTime('  09:00  '), '09:00');
    assert.equal(normalizeClockTime('Uhrzeit 14:00'), '14:00');
  });

  test('MEHRDEUTIGES wird verworfen statt geraten (keine erfundenen Zeiten)', () => {
    // Eine nackte "8" koennte 08:00 ODER 20:00 sein -> nicht ableitbar.
    assert.equal(normalizeClockTime('8'), null);
    assert.equal(normalizeClockTime(''), null);
    assert.equal(normalizeClockTime(null), null);
    assert.equal(normalizeClockTime('Work week'), null);
    assert.equal(normalizeClockTime('25:99'), null);
  });
});

describe('mapXToDayColumn - Chip x-Position -> Wochentag (Defekt 1, Geometrie)', () => {
  // 5 Spalten a 100px: Mo 100-200, Di 200-300, Mi 300-400, Do 400-500, Fr 500-600
  const columns = [100, 200, 300, 400, 500].map((left) => ({ left, width: 100 }));

  test('Treffer in der Spalte', () => {
    assert.equal(mapXToDayColumn(columns, 150), 0);
    assert.equal(mapXToDayColumn(columns, 250), 1);
    assert.equal(mapXToDayColumn(columns, 550), 4);
  });

  test('Randfall exakt auf der Kante', () => {
    assert.equal(mapXToDayColumn(columns, 200), 0);
  });

  test('schmaler Spaltenkopf -> naechste Spalte, solange der Abstand passt', () => {
    // OWA rendert den Tageskopf oft SCHMALER als die Spalte (nur die Zahl).
    assert.equal(mapXToDayColumn([{ left: 145, width: 10 }].concat(columns.slice(1)), 160), 0);
  });

  test('weit ausserhalb -> -1 (KEINE Zuordnung raten)', () => {
    assert.equal(mapXToDayColumn(columns, 5000), -1);
    assert.equal(mapXToDayColumn(columns, 10), -1);
  });

  test('ohne brauchbare Spalten -> -1', () => {
    assert.equal(mapXToDayColumn([], 150), -1);
    assert.equal(mapXToDayColumn(null, 150), -1);
    assert.equal(mapXToDayColumn([{ left: 100, width: 0 }], 150), -1);
  });
});

describe('mapYToClockTime - Chip y-Position -> Startzeit (Defekt 1, Geometrie)', () => {
  // Stundenraster 8-17 Uhr, 60px je Stunde, Oberkante bei top=200
  const ruler = [8, 9, 10, 11, 12, 13, 14, 15, 16, 17].map((h, i) => ({
    time: `${String(h).padStart(2, '0')}:00`,
    top: 200 + i * 60
  }));

  test('Chip kurz nach einer Rasterlinie -> diese Stunde', () => {
    assert.equal(mapYToClockTime(ruler, 205), '08:00');
    assert.equal(mapYToClockTime(ruler, 320), '10:00');
  });

  test('Chip auf halber Stunde -> die Stunde darunter (Start, nicht Mitte)', () => {
    assert.equal(mapYToClockTime(ruler, 250), '09:00');
  });

  test('oberhalb des Rasters -> keine Zuordnung', () => {
    assert.equal(mapYToClockTime(ruler, 100), null);
  });

  test('unterhalb des letzten Rasterpunkts -> linear fortgeschrieben (Stundenhöhe belegt)', () => {
    // Raster endet 17:00 (top 740). Ein Chip auf top 860 liegt 2h weiter unten.
    assert.equal(mapYToClockTime(ruler, 860), '19:00');
    // 30 Minuten Toleranz: knapp unterhalb -> noch keine Verschiebung
    assert.equal(mapYToClockTime(ruler, 760), '17:00');
  });

  test('ohne belegte Stundenhöhe wird unterhalb NICHT geraten', () => {
    const oneEntry = [{ time: '08:00', top: 200 }];
    assert.equal(mapYToClockTime(oneEntry, 500), null);
  });

  test('leeres/ungueltiges Raster -> null statt Rateversuch', () => {
    assert.equal(mapYToClockTime([], 300), null);
    assert.equal(mapYToClockTime(null, 300), null);
    assert.equal(mapYToClockTime([{ time: 'kaputt', top: 200 }], 300), null);
  });
});

describe('estimateHourPx - Stundenhöhe aus dem Raster (Endezeit)', () => {
  test('regelmaessiges Raster -> 60px', () => {
    const ruler = [8, 9, 10, 11].map((h, i) => ({ time: `${String(h).padStart(2, '0')}:00`, top: i * 60 }));
    assert.equal(estimateHourPx(ruler), 60);
  });

  test('Rauschen wird ueber Mehrheitsabstand geglaettet', () => {
    const ruler = [
      { time: '08:00', top: 0 }, { time: '09:00', top: 60 },
      { time: '10:00', top: 121 }, { time: '11:00', top: 180 }
    ];
    const px = estimateHourPx(ruler);
    assert.ok(px >= 55 && px <= 65, `erwartet ~60, war ${px}`);
  });

  test('zu wenige Stufen -> 0 (keine Endezeit erfinden)', () => {
    assert.equal(estimateHourPx([{ time: '08:00', top: 0 }, { time: '09:00', top: 60 }]), 60);
    assert.equal(estimateHourPx([{ time: '08:00', top: 0 }]), 0);
    assert.equal(estimateHourPx([]), 0);
    assert.equal(estimateHourPx(null), 0);
  });
});

describe('buildGeometryEventLabel - Label im Parser-Format aus Geometrie', () => {
  test('zeitbasiertes Event mit Datum aus der Spalte', () => {
    const l = buildGeometryEventLabel({ title: 'Block', start: '13:00', end: '18:30', dateIso: '2026-09-30' });
    assert.equal(l, 'Block, 13:00 to 18:30, Wednesday, September 30, 2026');
  });

  test('Ganztag', () => {
    const l = buildGeometryEventLabel({ title: 'Urlaub', allDay: true, dateIso: '2026-09-30' });
    assert.equal(l, 'Urlaub, all day event, Wednesday, September 30, 2026');
  });

  test('deutsches Datum wird ebenfalls gebaut (Locale robust)', () => {
    const l = buildGeometryEventLabel({ title: 'Termin', start: '09:00', end: '10:00', dateIso: '2026-09-28' });
    assert.ok(l.includes('Monday, September 28, 2026'), l);
  });

  test('ohne brauchbare Angaben -> null (KEIN Phantom-Label)', () => {
    assert.equal(buildGeometryEventLabel({ title: 'Block', start: '13:00', end: '18:30' }), null);
    assert.equal(buildGeometryEventLabel({ title: 'Block', dateIso: '2026-09-30' }), null);
    assert.equal(buildGeometryEventLabel({ start: '13:00', end: '14:00', dateIso: '2026-09-30' }), null);
    assert.equal(buildGeometryEventLabel({ title: 'Block', start: '13:00', end: '14:00', dateIso: 'kein-datum' }), null);
  });

  test('der gebaute Label ist vom echten Parser lesbar (End-to-End)', () => {
    const l = buildGeometryEventLabel({ title: 'Block', start: '13:00', end: '18:30', dateIso: '2026-09-30' });
    const slot = parseEventLabel(l, 2026);
    assert.ok(slot);
    assert.equal(slot.date, '2026-09-30');
    assert.equal(slot.start, '13:00');
    assert.equal(slot.end, '18:30');
    assert.equal(slot.title, 'Block');
  });
});

describe('deriveEventLabelFromNode - Label aus einem beliebigen Grid-Knoten', () => {
  const pats = EVENT_LABEL_PATTERNS;
  const columns = [100, 200, 300].map((left) => ({ left, width: 100 }));
  const ruler = [8, 9, 10, 11, 12, 13].map((h, i) => ({ time: `${String(h).padStart(2, '0')}:00`, top: 200 + i * 60 }));
  const geo = { columns, ruler, weekStartIso: '2026-09-28' };
  const node = (over = {}) => ({
    tag: 'DIV', role: null, cls: 'ms-Cal-event',
    ariaLabel: '', title: '', text: '', data: null, inGrid: true,
    rect: { left: 250, top: 620, width: 90, height: 150, x: 250, y: 620 },
    area: true, offsetParentNull: true, ...over
  });

  test('1) vollstaendiges aria-Label bleibt erste Wahl (timeSource: label)', () => {
    const r = deriveEventLabelFromNode(node({ ariaLabel: 'Block, 13:00 to 18:30, Wednesday, September 30, 2026, Busy' }), { pats });
    assert.equal(r.label, 'Block, 13:00 to 18:30, Wednesday, September 30, 2026, Busy');
    assert.equal(r.reason, null);
    assert.equal(r.source, 'aria-label');
    assert.equal(r.timeSource, 'label');
  });

  test('2) title-Attribut traegt das Label, wenn aria-label fehlt', () => {
    const r = deriveEventLabelFromNode(node({ title: 'Urlaub, all day event, Wednesday, September 30, 2026, Busy' }), { pats });
    assert.equal(r.source, 'title');
    assert.ok(r.label);
    assert.equal(r.reason, null);
  });

  test('3) data-Attribut traegt das Label (OWA-typisch)', () => {
    const r = deriveEventLabelFromNode(node({ data: { 'data-event-aria-label': 'Block, 13:00 to 18:30, Wednesday, September 30, 2026, Busy' } }), { pats });
    assert.equal(r.source, 'data-attr');
    assert.ok(r.label);
  });

  test('4) reiner Text MIT Zeit UND Datum ist ein vollwertiges Label', () => {
    const r = deriveEventLabelFromNode(node({ text: 'Austausch Mathis:David:Alex\n14:00 to 15:00\nMonday, September 28, 2026' }), { pats });
    assert.equal(r.source, 'text');
    assert.equal(r.reason, null);
    assert.equal(r.timeSource, 'label');
  });

  test('5) NUR Titel im Text -> Zeit+Datum werden aus der Geometrie abgeleitet', () => {
    // Chip ohne JEDE Beschriftung ausser dem Titel — der Defekt-1-Fall.
    // Raster 8-13 Uhr (60px/Stunde ab top 200), Chip top 620 -> 15:00
    // (Fortschreibung ueber das Rasterende hinaus), Hoehe 150px -> 2,5h.
    const r = deriveEventLabelFromNode(node({ text: 'Block' }), { pats, ...geo });
    assert.equal(r.reason, null);
    assert.equal(r.source, 'text');
    assert.equal(r.timeSource, 'geometry');
    // Spalte 250px -> Index 1 = Dienstag 29.09.
    assert.equal(r.label, 'Block, 15:00 to 17:30, Tuesday, September 29, 2026');
  });

  test('5b) Toolbar-Text AUSSERHALB des Grids wird NIE zum Event', () => {
    // "Work week" ist der Ansichtswechsler, KEIN Termin. Ohne inGrid-Guard
    // wuerde die Geometrie daraus einen Phantom-Slot bauen.
    const r = deriveEventLabelFromNode(node({ text: 'Work week', inGrid: false }), { pats, ...geo });
    assert.equal(r.label, null);
    assert.equal(r.reason, 'no-event-signature');
  });

  test('6) OHNE Geometrie wird NICHT geraten -> reason no-time-source', () => {
    const r = deriveEventLabelFromNode(node({ text: 'Block' }), { pats });
    assert.equal(r.label, null);
    assert.equal(r.reason, 'no-time-source');
  });

  test('7) Uhrzeit-Indikator bleibt ausgeschlossen', () => {
    const r = deriveEventLabelFromNode(node({ ariaLabel: 'calendar view, current time: Fri 22:31' }), { pats, ...geo });
    assert.equal(r.reason, 'clock-indicator');
    assert.equal(r.label, null);
  });

  test('8) Arbeitsort-Banner bleibt ausgeschlossen', () => {
    const r = deriveEventLabelFromNode(node({ text: 'Arbeitsort\n09:00 to 17:00' }), { pats, ...geo });
    assert.equal(r.reason, 'work-location');
  });

  test('9) reines Rauschen wird verworfen', () => {
    assert.equal(deriveEventLabelFromNode(node({ text: 'Work week', inGrid: false }), { pats, ...geo }).reason, 'no-event-signature');
    assert.equal(deriveEventLabelFromNode(node({}), { pats, ...geo }).reason, 'no-event-signature');
  });

  test('10) Element OHNE Flaeche wird nie zum Event (rect>0 als Sichtbarkeits-Beweis)', () => {
    const r = deriveEventLabelFromNode(
      node({ ariaLabel: 'Block, 13:00 to 18:30, Wednesday, September 30, 2026', area: false, rect: { left: 0, top: 0, width: 0, height: 0 } }),
      { pats, ...geo }
    );
    assert.equal(r.reason, 'no-area');
  });

  test('11) offsetParent===null ist KEIN Hard-K.O. (Virtualisierung)', () => {
    const r = deriveEventLabelFromNode(
      node({ ariaLabel: 'Block, 13:00 to 18:30, Wednesday, September 30, 2026', offsetParentNull: true, area: true }),
      { pats }
    );
    assert.equal(r.reason, null);
    assert.ok(r.label);
  });
});

describe('resolveDataQuality - ehrliche Verifikationsangabe (Defekt 2)', () => {
  test('verified NUR wenn Woche UND echte Event-Knoten bestaetigt sind', () => {
    assert.equal(resolveDataQuality({ weekVerified: true, eventElementsFound: 6 }), 'verified');
  });

  test('Woche ok, aber 0 Event-Knoten -> unverified-empty (KEIN "frei")', () => {
    // Genau der Smoke-Test-Befund: dataQuality:'verified' bei
    // eventElementsFound:0 war eine Falschaussage.
    assert.equal(resolveDataQuality({ weekVerified: true, eventElementsFound: 0 }), 'unverified-empty');
  });

  test('Woche nicht lesbar -> unverified (unabhaengig von den Events)', () => {
    assert.equal(resolveDataQuality({ weekVerified: false, eventElementsFound: 6 }), 'unverified');
    assert.equal(resolveDataQuality({ weekVerified: false, eventElementsFound: 0 }), 'unverified');
  });

  test('negativ/fehlende Event-Zahl gilt als 0 (nicht als "hat Events")', () => {
    assert.equal(resolveDataQuality({ weekVerified: true, eventElementsFound: -1 }), 'unverified-empty');
    assert.equal(resolveDataQuality({ weekVerified: true, eventElementsFound: null }), 'unverified-empty');
    assert.equal(resolveDataQuality({ weekVerified: true }), 'unverified-empty');
  });

  test('das dokumentierte Werte-Trio ist geschlossen', () => {
    assert.deepEqual(DATA_QUALITY_VALUES, ['verified', 'unverified-empty', 'unverified']);
  });

  test('nur "verified" gilt als belastbare Belegungsaussage', () => {
    const reliable = (dq) => dq === 'verified';
    assert.equal(reliable(resolveDataQuality({ weekVerified: true, eventElementsFound: 3 })), true);
    assert.equal(reliable(resolveDataQuality({ weekVerified: true, eventElementsFound: 0 })), false);
    assert.equal(reliable(resolveDataQuality({ weekVerified: false, eventElementsFound: 3 })), false);
  });
});

describe('resolveEventWaitOutcome - Ausgang an die BEOBACHTUNG koppeln (Defekt 3)', () => {
  test('Event-Knoten gesehen -> found ( unabhaengig von der Wartezeit)', () => {
    assert.equal(resolveEventWaitOutcome({ polls: 1, eventElementsFound: 6, waitBudgetMs: 20000 }), 'found');
  });

  test('0 Event-Knoten NIE "found" — auch nicht bei polls 0 (Smoke-Test-Befund)', () => {
    // Live-Befund: eventWaitOutcome:"found" bei waitedForEventsMs:0 und
    // eventElementsFound:0. Genau das darf nicht mehr passieren.
    assert.notEqual(resolveEventWaitOutcome({ polls: 0, eventElementsFound: 0, waitBudgetMs: 20000 }), 'found');
    assert.equal(resolveEventWaitOutcome({ polls: 0, eventElementsFound: 0, waitBudgetMs: 20000 }), 'not-found');
  });

  test('0 Polls + 0 Events -> not-found (Wartephase lief gar nicht)', () => {
    assert.equal(resolveEventWaitOutcome({ polls: 0, eventElementsFound: 0, waitBudgetMs: 0 }), 'not-found');
    assert.equal(resolveEventWaitOutcome({ polls: 0, eventElementsFound: 0, waitBudgetMs: null }), 'not-found');
  });

  test('Polls ohne Treffer -> not-found, das Budget wird getrennt berichtet', () => {
    assert.equal(resolveEventWaitOutcome({ polls: 40, eventElementsFound: 0, waitBudgetMs: 20000 }), 'not-found');
  });

  test('Ausgaenge sind genau die dokumentierten Werte', () => {
    assert.deepEqual(EVENT_WAIT_OUTCOMES, ['found', 'not-found', 'budget-exhausted', 'not-rendered']);
  });

  test('die Warte-Warnung erklaert not-found genauso deutlich wie budget-exhausted', () => {
    const w = eventWaitWarning({ outcome: 'not-found', waitedMs: 20000, budgetMs: 20000 });
    assert.ok(w);
    assert.ok(/20000ms/.test(w), w);
    assert.ok(/leeren Kalender/.test(w), w);
    assert.ok(/census/.test(w), w);
  });
});

describe('collectCalendarEventLabels - Quellen-Leiter + Struktur-Census (Defekt 1)', () => {
  const fakePage = { evaluate: async (fn, arg) => fn(arg) };
  const withFakeDocument = async (doc, fn) => {
    const prev = globalThis.document;
    globalThis.document = doc;
    try { return await fn(); } finally { globalThis.document = prev; }
  };
  // Minimaler Knoten: der Kollektor darf NICHTS voraussetzen, was ein
  // unit-test-Fake nicht liefert (tagName/className/innerText/attributes...).
  const chip = (over = {}) => ({
    tagName: 'DIV',
    getAttribute: (n) => (over.attrs && over.attrs[n] !== undefined ? over.attrs[n] : null),
    getBoundingClientRect: () => over.rect || { left: 0, top: 0, width: 120, height: 24, x: 0, y: 0 },
    offsetParent: null,
    innerText: over.text || '',
    className: over.className || '',
    attributes: Object.entries(over.attrs || {}).map(([name, value]) => ({ name, value })),
    childElementCount: 0,
    children: [],
    ...over.extra
  });
  const emptyDoc = (chips = []) => ({
    querySelectorAll: (sel) => (sel === null || sel === '*' ? chips : []),
    querySelector: () => null,
    body: null
  });

  test('Event im TITLE statt aria-label wird jetzt gefunden (vorher: 0 Slots)', () => {
    const doc = emptyDoc([chip({ attrs: { title: 'Urlaub, all day event, Wednesday, September 30, 2026, Busy' } })]);
    return withFakeDocument(doc, async () => {
      const res = await collectCalendarEventLabels(fakePage);
      assert.equal(res.eventElementsFound, 1);
      assert.equal(res.sourceUsed, 'title');
      assert.ok(res.labels[0].includes('Urlaub'));
    });
  });

  test('Struktur-Census wird IMMER geliefert (auch bei 0 Events) — das ist die Bodenwahrheit', () => {
    const doc = emptyDoc([]);
    return withFakeDocument(doc, async () => {
      const res = await collectCalendarEventLabels(fakePage);
      assert.ok(Array.isArray(res.sourceTried), 'sourceTried fehlt');
      assert.ok(res.sourceTried.length >= 2, 'die Leiter muss protokolliert sein');
      assert.ok(res.census, 'census fehlt');
      assert.ok(Array.isArray(res.census.scanned), 'census.scanned fehlt');
    });
  });

  test('alle Quellen der Leiter werden protokolliert, nicht nur der Gewinner', () => {
    const doc = emptyDoc([]);
    return withFakeDocument(doc, async () => {
      const res = await collectCalendarEventLabels(fakePage);
      const ids = res.sourceTried.map((s) => s.source || s.id);
      assert.ok(ids.includes('title'), 'title-Quelle wurde nicht probiert');
      assert.ok(ids.includes('data-attr'), 'data-attr-Quelle wurde nicht probiert');
      assert.ok(ids.includes('text'), 'text-Quelle wurde nicht probiert');
    });
  });

  test('0 Events -> eventElementsFound 0 und KEIN phantom-found', () => {
    const doc = emptyDoc([chip({ attrs: { 'aria-label': 'Work week' } })]);
    return withFakeDocument(doc, async () => {
      const res = await collectCalendarEventLabels(fakePage);
      assert.equal(res.eventElementsFound, 0);
      assert.equal(res.labels.length, 0);
    });
  });

  test('evaluate-Fehlschlag liefert eine ehrliche 0-Struktur (kein Absturz, kein false found)', async () => {
    const brokenPage = { evaluate: async () => { throw new Error('Execution context was destroyed'); } };
    const res = await collectCalendarEventLabels(brokenPage);
    assert.equal(res.eventElementsFound, 0);
    assert.equal(res.labels.length, 0);
    assert.equal(res.collectError, 'Execution context was destroyed');
  });

  // -------------------------------------------------------------------
  // Der Geometrie-Weg muss auch WIRKLICH durch den Kollektor laufen —
  // eine nur unit-getestete Hilfsfunktion waere toter Code.
  // -------------------------------------------------------------------
  // Fake-DOM einer OWA-Wochenansicht: 3 Tageskoepfe, Stundenraster 8-18 Uhr
  // (60px je Stunde, Oberkante 200) und ein Chip "Block" ohne JEDE Beschriftung
  // ausser Titel — genau der Fall, an dem [aria-label] scheitert.
  const gridDoc = () => {
    const col = (left) => ({ tagName: 'TH', getAttribute: (n) => (n === 'role' ? 'columnheader' : null), getBoundingClientRect: () => ({ left, top: 120, width: 100, height: 30, x: left, y: 120 }), innerText: '', attributes: [] });
    const hour = (top) => ({ tagName: 'DIV', getAttribute: () => null, getBoundingClientRect: () => ({ left: 10, top, width: 40, height: 20, x: 10, y: top }), innerText: `${String((top - 200) / 60 + 8).padStart(2, '0')}:00`, attributes: [] });
    const chip = () => ({ tagName: 'DIV', className: 'ms-Cal-event', getAttribute: () => null, getBoundingClientRect: () => ({ left: 250, top: 560, width: 90, height: 180, x: 250, y: 560 }), innerText: 'Block', offsetParent: null, children: [], attributes: [] });
    const els = [col(100), col(200), col(300)].concat([200, 260, 320, 380, 440, 500, 560, 620, 680, 740, 800].map(hour), [chip()]);
    return {
      querySelectorAll: (sel) => {
        if (sel === '[role="columnheader"], [role="grid"] th, th') return [col(100), col(200), col(300)];
        if (sel === '*') return els;
        return [];
      },
      querySelector: () => null,
      body: null
    };
  };

  test('Geometrie-Weg laeuft durch den Kollektor: Chip ohne aria-label wird zum Label', async () => {
    await withFakeDocument(gridDoc(), async () => {
      const res = await collectCalendarEventLabels(fakePage, { weekStartIso: '2026-09-28' });
      assert.equal(res.eventElementsFound, 1, `erwartet 1 Event, Labels: ${JSON.stringify(res.labels)}`);
      assert.equal(res.sourceUsed, 'text+geometry');
      // Spalte 250px -> Index 1 = Dienstag; top 560px -> 14:00; Hoehe 180px -> 3h
      assert.equal(res.labels[0], 'Block, 14:00 to 17:00, Tuesday, September 29, 2026');
      // und der echte Parser kann es lesen
      const slot = parseEventLabel(res.labels[0], 2026);
      assert.equal(slot.date, '2026-09-29');
      assert.equal(slot.start, '14:00');
      assert.equal(slot.end, '17:00');
      assert.equal(slot.title, 'Block');
    });
  });

  test('Geometrie braucht Tageskoepfe+WOCHENSTART: ohne weekStartIso wird nichts geraten', async () => {
    await withFakeDocument(gridDoc(), async () => {
      const res = await collectCalendarEventLabels(fakePage); // KEIN weekStartIso
      assert.equal(res.eventElementsFound, 0);
      assert.equal(res.labels.length, 0);
    });
  });

  test('fast-Mode ueberspringt den Struktur-Scan (Poll-Budget) und liefert trotzdem Attribute', async () => {
    await withFakeDocument(gridDoc(), async () => {
      const res = await collectCalendarEventLabels(fakePage, { weekStartIso: '2026-09-28', fast: true });
      assert.equal(res.eventElementsFound, 0, 'im fast-Mode wird die Geometrie nicht ausgewertet');
      assert.deepEqual(res.census, { skipped: 'fast-mode' });
      assert.equal(res.structureNodes.length, 0);
    });
  });

  test('Census liefert Tages-Spalten und Zeitraster (Beweis fuer die Geometrie)', async () => {
    await withFakeDocument(gridDoc(), async () => {
      const res = await collectCalendarEventLabels(fakePage, { weekStartIso: '2026-09-28' });
      assert.equal(res.census.dayColumns.length, 3);
      assert.ok(res.census.clockRuler.length >= 3, `Raster: ${JSON.stringify(res.census.clockRuler)}`);
      assert.ok(res.census.structureNodes >= 1);
      assert.ok(res.census.scanned.length >= 1);
      assert.ok('tag' in res.census.scanned[0] && 'rect' in res.census.scanned[0]);
    });
  });
});

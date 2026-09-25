import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { resolveFolder, buildComposeUrl, buildCalendarComposeUrl, buildCalendarWeekUrl, normalizeCalendarDate, FOLDER_MAP, parseWeekHeaderText, parseEventLabel, classifyEventStatus, parseMailRowTimestamp, parseNumericDayHeaderText, resolveRenderedWeekFromCandidates, isWorkLocationPseudoSlot, dedupeBusySlots, partitionSlotsByRange, formatCalendarDiagnostics, FREEBUSY_TOTAL_BUDGET_MS, FREEBUSY_MAX_BUTTON_CLICKS, FREEBUSY_CLICK_RENDER_WAIT_MS, FREEBUSY_MIN_REMAINING_FOR_ESCALATION_MS, phaseBudgetMs, affordableClicks, weekButtonCandidates, evaluateClickVerification, shouldTryNextCandidate, classifyButtonFailure, assembleFreeBusyDiagnostics, slotExtractionWarning, collectCalendarEventLabels, isCalendarEventLabel, filterCalendarHeaderCandidates, isVisibleCalendarElement, parseWeekNumberFromNav, parseNavWeekRange, isoWeekMonday, resolveWeekFromNavButtons, MAILBOX_VERIFICATION_SOURCES, resolveMailboxVerification } from '../src/owa.js';
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

  const fakePage = { evaluate: async (fn) => fn() };

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
});

import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { resolveFolder, buildComposeUrl, buildCalendarComposeUrl, normalizeCalendarDate, FOLDER_MAP, parseWeekHeaderText, parseEventLabel, parseMailRowTimestamp } from '../src/owa.js';
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


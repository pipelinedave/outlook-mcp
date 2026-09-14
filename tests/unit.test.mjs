import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { resolveFolder, buildComposeUrl, buildCalendarComposeUrl, normalizeCalendarDate, FOLDER_MAP } from '../src/owa.js';
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


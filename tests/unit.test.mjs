import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { resolveFolder, buildComposeUrl, FOLDER_MAP } from '../src/owa.js';
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

describe('config - Defaults & Umgebungsvariablen', () => {
  test('ProfileDir ist gesetzt', () => {
    assert.ok(config.profileDir.length > 0);
  });

  test('DefaultRealm ist ein String', () => {
    assert.equal(typeof config.defaultRealm, 'string');
    assert.ok(config.defaultRealm.length > 0);
  });
});

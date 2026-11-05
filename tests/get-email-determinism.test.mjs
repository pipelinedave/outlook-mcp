import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { OwaClient } from '../src/owa.js';

// ===========================================================================
// Runde 2, Befund 1 (Live-Lauf 05.11.2026, d-velop-Realm): Auf 3 IDENTISCHE
// outlook_get_email-Calls mit derselben Conversation-ID renderte das Tool
// jedes Mal eine ANDERE, fremde Mail ("Beschaffungsauftrag Stand",
// "Gesprächszusammenfassung 15.10.2026", "dms - BA (Beschaffungsantrag UHK)").
//
// Root Cause: War die ID in der gerade gerenderten Liste nicht enthalten
// (geteilter Daemon-Browser in fremdem Zustand — alte Suchliste, Inbox,
// Kalender-Tab), klickte _getEmailDetailsInner stillschweigend rows[0].
// Der ID→Mail-Mapping war damit nicht deterministisch.
//
// Fix: ID-Pfad findet NUR die exakt treffende Zeile (data-convid/id),
// fällt bei Nichttreffer auf begrenzte Scroll-Enumeration zurück und wirft
// dann einen ehrlichen Fehler — NIEMALS rows[0]. Out-of-Range-Indizes
// werfen ebenfalls, statt still Zeile 0 zu rendern.
// Hier mit gemockter Playwright-Page (ohne echten Browser).
// ===========================================================================

function makeMockPage({ renderedIds = [], locatorCount = null } = {}) {
  const clicks = [];
  const rows = renderedIds.map((id) => ({
    getAttribute: async (attr) => (attr === 'data-convid' ? id : `elem-${id}`),
    click: async () => clicks.push(`row:${id}`)
  }));
  const locatorClicks = [];
  return {
    clicks,
    locatorClicks,
    waitForSelector: async () => true,
    waitForTimeout: async () => {},
    $$: async () => rows,
    evaluate: async () => {},
    locator: (selector) => {
      const first = {
        count: async () => (typeof locatorCount === 'function' ? locatorCount(selector) : (locatorCount ?? 0)),
        click: async () => locatorClicks.push(selector)
      };
      return { first: () => first };
    }
  };
}

function makeClient(page) {
  const client = new OwaClient('test.local');
  client.getPage = async () => page;
  client._renderReadingPane = async (_page, opts) => ({
    renderStrategy: 'mock', messageIndex: opts?.messageIndex ?? null
  });
  return client;
}

describe('get_email ID-Determinismus (Runde 2, Befund 1)', () => {
  test('String-ID in der Liste: exakt diese Zeile wird geklickt', async () => {
    const page = makeMockPage({ locatorCount: 1 });
    const client = makeClient(page);
    const r = await client.getEmailDetails('conv-42');
    assert.equal(r.renderStrategy, 'mock');
    assert.equal(page.locatorClicks.length, 1);
    assert.match(page.locatorClicks[0], /data-convid="conv-42"/);
  });

  test('String-ID NICHT in der Liste: ehrlicher Fehler, KEIN rows[0]-Klick', async () => {
    const page = makeMockPage({ renderedIds: ['fremd-1', 'fremd-2'], locatorCount: 0 });
    const client = makeClient(page);
    await assert.rejects(
      () => client.getEmailDetails('doku-fis-convid'),
      (err) => {
        assert.match(err.message, /nicht vorhanden/);
        assert.match(err.message, /KEINE andere E-Mail gerendert/);
        return true;
      }
    );
    // Beweis: Es wurde weder eine fremde Zeile noch irgendein Locator geklickt.
    assert.equal(page.clicks.length, 0);
    assert.equal(page.locatorClicks.length, 0);
  });

  test('String-ID erst nach Scroll-Enumeration gefunden: Treffer-Zeile geklickt', async () => {
    // locatorCount als Funktion: Treffer erst beim 3. Versuch (nach 2 Scroll-Polls)
    let attempts = 0;
    const page = makeMockPage({
      locatorCount: () => (++attempts >= 3 ? 1 : 0)
    });
    const client = makeClient(page);
    const r = await client.getEmailDetails('deep-conv');
    assert.equal(r.renderStrategy, 'mock');
    assert.equal(page.locatorClicks.length, 1);
    assert.match(page.locatorClicks[0], /data-convid="deep-conv"/);
  });

  test('Out-of-Range-Index: ehrlicher Fehler, kein stiller rows[0]-Render', async () => {
    const page = makeMockPage({ renderedIds: ['a', 'b'] });
    const client = makeClient(page);
    await assert.rejects(
      () => client.getEmailDetails(7),
      (err) => {
        assert.match(err.message, /ausserhalb der aktuellen Liste/);
        return true;
      }
    );
    assert.equal(page.clicks.length, 0);
  });

  test('Gültiger Index: genau diese Zeile wird geklickt', async () => {
    const page = makeMockPage({ renderedIds: ['a', 'b', 'c'] });
    const client = makeClient(page);
    const r = await client.getEmailDetails(2);
    assert.equal(r.renderStrategy, 'mock');
    assert.deepEqual(page.clicks, ['row:c']);
  });

  test('Negativer Index: ehrlicher Fehler', async () => {
    const page = makeMockPage({ renderedIds: ['a', 'b'] });
    const client = makeClient(page);
    await assert.rejects(() => client.getEmailDetails(-1), /ausserhalb der aktuellen Liste/);
    assert.equal(page.clicks.length, 0);
  });
});

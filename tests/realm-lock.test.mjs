import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { withRealmLock, realmQueueSize, resetRealmLocksForTest } from '../src/realm-lock.js';

// ===========================================================================
// Runde 2 (Live-Befund 05.11.2026): Der geteilte Daemon-Browser hält pro Realm
// EINE Page. Ohne Serialisierung interleaven parallele Tool-Calls ihre
// mehrstufigen UI-Flows — get_email renderte dadurch fremde Mails, Kalender-
// Reads trafen transiente Snapshots. Der Realm-Lock macht jeden Call atomar
// je Realm; verschiedene Realms laufen weiter parallel.
// ===========================================================================

describe('withRealmLock', () => {
  test('gleicher Realm: Calls laufen strikt nacheinander (kein Interleaving)', async () => {
    resetRealmLocksForTest();
    const events = [];
    const flow = (tag, ms) => withRealmLock('d-velop.de', async () => {
      events.push(`${tag}:start`);
      await new Promise((r) => setTimeout(r, ms));
      events.push(`${tag}:end`);
    });
    // Beide starten "gleichzeitig" — der Lock muss sie hintereinander reihen.
    await Promise.all([flow('A', 40), flow('B', 10)]);
    assert.deepEqual(events, ['A:start', 'A:end', 'B:start', 'B:end']);
  });

  test('verschiedene Realms: parallel (Realm-Multiplex bleibt bestehen)', async () => {
    resetRealmLocksForTest();
    const events = [];
    const flow = (tag, realm, ms) => withRealmLock(realm, async () => {
      events.push(`${tag}:start`);
      await new Promise((r) => setTimeout(r, ms));
      events.push(`${tag}:end`);
    });
    await Promise.all([flow('A', 'd-velop.de', 40), flow('B', 'adesso.de', 10)]);
    // Beweis Parallelität: B (kuerzer) startet noch WAEHREND A laeuft.
    assert.ok(events.indexOf('B:start') < events.indexOf('A:end'),
      `B:start (${events.indexOf('B:start')}) muss vor A:end (${events.indexOf('A:end')}) liegen: ${events.join(',')}`);
  });

  test('Fehler eines Vorgaengers blockieren den Nachfolger nicht', async () => {
    resetRealmLocksForTest();
    await assert.rejects(
      () => withRealmLock('x', async () => { throw new Error('Flow A fehlgeschlagen'); }),
      /Flow A fehlgeschlagen/
    );
    const r = await withRealmLock('x', async () => 'B-laeuft');
    assert.equal(r, 'B-laeuft');
    // Cleanup ist ein Microtask nach dem Resolve — einen Tick abwarten.
    await new Promise((res) => setImmediate(res));
    assert.equal(realmQueueSize('x'), 0, 'Queue-Eintrag muss nach Abschluss entfernt sein');
  });

  test('ohne realm: default-Key wird benutzt (gleiche Kette wie null)', async () => {
    resetRealmLocksForTest();
    const events = [];
    const flow = (tag, realm) => withRealmLock(realm, async () => {
      events.push(`${tag}:start`);
      await new Promise((r) => setTimeout(r, 15));
      events.push(`${tag}:end`);
    });
    await Promise.all([flow('A', null), flow('B', '')]);
    assert.deepEqual(events, ['A:start', 'A:end', 'B:start', 'B:end']);
  });
});

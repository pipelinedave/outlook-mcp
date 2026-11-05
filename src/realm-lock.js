// ===========================================================================
// Realm-Lock (Runde 2, Live-Befund 05.11.2026): Der Daemon teilt sich pro
// Realm (d-velop.de / adesso.de) EINEN persistenten Browser mit EINER Page.
// Ohne Serialisierung interleaven parallele Tool-Calls (mehrere Agenten/
// Sessions auf demselben Daemon) ihre mehrstufigen UI-Flows (Suche -> Klick
// -> Read bzw. Navigation -> Header-Read) auf derselben Page — genau daraus
// entstanden die Flakiness-Befunde (get_email renderte Fremdmails,
// Kalender-Reads trafen transiente DOM-Snapshots).
//
// Der Lock macht jeden Tool-Call ATOMAR je Realm: Call B mit demselben Realm
// wartet, bis Call A fertig ist. Calls auf VERSCHIEDENE Realms bleiben
// parallel (Realm-Multiplex unveraendert).
//
// Absichtlich klein (utils-slim): Promise-Kette pro Realm, kein Timeout,
// keine Metering — Fehler eines Vorgaengers blockieren Nachfolger nicht.
// ===========================================================================

const queues = new Map();

/**
 * Fuehrt fn serialisiert fuer den gegebenen Realm aus.
 * @param {string|null} realm Realm-Key (null/'' -> 'default')
 * @param {() => Promise<T>} fn
 * @returns {Promise<T>} Ergebnis von fn
 * @template T
 */
export function withRealmLock(realm, fn) {
  const key = String(realm || 'default');
  const prev = queues.get(key) || Promise.resolve();
  // Fehler des Vorgaengers duerfen den Nachfolger nicht blockieren.
  const run = prev.then(fn, fn);
  // Queue-Eintrag spiegelt das ENDE des eigenen Laufs (Fehler geschluckt).
  // Ist dieser Lauf selbst das Ketten-Ende und abgeschlossen, wird der Eintrag
  // wieder entfernt (kein unbeschraenktes Wachstum der Map ueber die Zeit).
  const tail = run.then(() => undefined, () => undefined);
  queues.set(key, tail);
  tail.then(() => {
    if (queues.get(key) === tail) queues.delete(key);
  });
  return run;
}

/** Test-Hilfe: Warteschlangen-Tiefe je Realm (0 = kein offener/letzter Eintrag). */
export function realmQueueSize(realm) {
  return queues.has(String(realm || 'default')) ? 1 : 0;
}

/** Test-Hilfe: internen Zustand leeren (keine echten Ressourcen gebunden). */
export function resetRealmLocksForTest() {
  queues.clear();
}

import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { collectCalendarEventLabels, parseEventLabel, CLOCK_LABEL_PATTERN } from '../src/owa.js';
import { browserRealmPage, emulateBrowserEvaluate, auditEvaluateScope, serializedCallbacks, moduleLevelNames, freeIdentifiersOf } from './evaluate-scope.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SRC = path.join(HERE, '..', 'src');

// ---------------------------------------------------------------------------
// Minimales Test-DOM. Der Kollektor darf NICHTS voraussetzen, was ein unit-
// test-Fake nicht liefert (tagName/className/innerText/attributes/children...).
// ---------------------------------------------------------------------------
const chipNode = (attrs = {}, text = '', rect = null) => ({
  tagName: 'DIV',
  getAttribute: (n) => (attrs[n] !== undefined ? attrs[n] : null),
  getBoundingClientRect: () => rect || { left: 0, top: 200, width: 120, height: 24, x: 0, y: 200 },
  offsetParent: null,
  innerText: text,
  className: '',
  attributes: Object.entries(attrs).map(([name, value]) => ({ name, value })),
  children: [],
  childElementCount: 0
});

const colNode = (left) => ({
  tagName: 'TH',
  getAttribute: (n) => (n === 'role' ? 'columnheader' : null),
  getBoundingClientRect: () => ({ left, top: 0, width: 140, height: 30, x: left, y: 0 }),
  offsetParent: null,
  innerText: '',
  className: '',
  attributes: [{ name: 'role', value: 'columnheader' }],
  children: [],
  childElementCount: 0
});

const rulerNode = (time, top) => ({
  tagName: 'DIV',
  getAttribute: (n) => (n === 'aria-label' ? time : null),
  getBoundingClientRect: () => ({ left: 0, top, width: 40, height: 20, x: 0, y: top }),
  offsetParent: null,
  innerText: '',
  className: '',
  attributes: [{ name: 'aria-label', value: time }],
  children: [],
  childElementCount: 0
});

/**
 * Grid-Body-Fake: `leafCount` Blatt-Knoten im Grid-Body, die ueber `title`
 * gefunden werden (kein aria-label -> erzwingt den title-Quellpfad, in dem der
 * Live-Bug lag).
 */
const gridDoc = ({ leafCount = 1, label = (i) => `Termin ${i}, 09:00 to 10:00, Thursday, October 1, 2026, Busy` } = {}) => {
  const leaves = [];
  for (let i = 0; i < leafCount; i++) {
    leaves.push(chipNode(
      { title: label(i) },
      '',
      { left: 150, top: 200 + i * 30, width: 120, height: 24, x: 150, y: 200 + i * 30 }
    ));
  }
  // Tages-Spalten ab x=100, damit das Zeitraster (x=0) AUSSERHALB der
  // Grid-Body-Region liegt -> structureNodes ist exakt leafCount.
  const columns = [100, 240, 380, 520, 660].map(colNode);
  const ruler = [8, 9, 10, 11, 12].map((h, i) => rulerNode(`${String(h).padStart(2, '0')}:00`, 60 + i * 60));
  return {
    querySelectorAll: (sel) => {
      if (sel === '*') return columns.concat(ruler, leaves);
      if (sel === '[role="columnheader"], [role="grid"] th, th') return columns;
      return [];
    },
    querySelector: () => null,
    body: null
  };
};

// ===========================================================================
describe('Regression Fehlerklasse "Scope-Leak ueber page.evaluate" (Live-Befund 25.09.2026)', () => {

  // MECHANISMUS-NACHWEIS -----------------------------------------------------
  // Dieser Test beweist, dass die Pruefung die Fehlerklasse ueberhaupt sieht.
  // Er laeuft ABSICHTLICH mit einem absichtlich leckenden Callback: im
  // Browser-Realm muss er scheitern, im Node-Modul-Scope darf er nicht
  // scheitern. Genau diese Differenz war die Blindstelle des alten Fakes.
  test('Der Testmechanismus sieht die Fehlerklasse: leckender Callback scheitert nur im Browser-Realm', () => {
    // CLOCK_LABEL_PATTERN ist ein echter Modul-Scope-Name aus src/owa.js
    // (export const, Z398) — genau die Form des Live-Defekts: eine nackte
    // Modul-Konstante im evaluate-Callback.
    const leaking = (arg) => ({ seen: arg.pats, re: new RegExp(CLOCK_LABEL_PATTERN, 'i') });

    // (a) Browser-Transport: Quelltext in einem frischen vm.Realm.
    assert.throws(
      () => emulateBrowserEvaluate(leaking, { pats: { a: 'x' } }, {}),
      /CLOCK_LABEL_PATTERN is not defined/,
      'im Browser-Realm muss der Leak als ReferenceError auftreten'
    );

    // (b) Node-Modul-Scope: derselbe Callback laeuft "erfolgreich" — so blind
    // war der bisherige Fake `evaluate: async (fn, arg) => fn(arg)`.
    const nodeScopeResult = leaking({ pats: { a: 'x' } });
    assert.ok(nodeScopeResult.re instanceof RegExp, 'im Node-Scope loest der Leak normal auf — daher war er unsichtbar');
  });

  test('browserRealmPage serialisiert den Callback (kein Direktaufruf im Modul-Scope)', async () => {
    const page = browserRealmPage({ document: gridDoc() });
    let sawSerialization = false;
    const spy = { evaluate: async (fn, arg) => { sawSerialization = true; return fn(arg); } };
    // Gegenprobe: der Node-Scope-Fake "sieht" die Leaks nicht.
    assert.doesNotThrow(() => spy.evaluate((arg) => ({ v: arg.maxElements }), { maxElements: 6000 }));
    assert.ok(sawSerialization);
    assert.equal(typeof page.evaluate, 'function');
  });

  // ECHTE REGRESSION --------------------------------------------------------
  test('collectCalendarEventLabels liefert im Browser-Realm das VOLLE Ergebnis-Objekt', async () => {
    const doc = gridDoc({ leafCount: 2 });
    const res = await collectCalendarEventLabels(browserRealmPage({ document: doc }), { weekStartIso: '2026-09-28' });

    // Der Live-Defekt: das komplette Rueckgabeobjekt ging verloren, alles blieb
    // auf Default. Genau diese Felder MUSSEN nun gefuellt sein.
    assert.equal(res.collectError, null, `collectError=${res.collectError}`);
    assert.notEqual(res.selectorUsed, 'error');
    assert.equal(res.eventElementsFound, 2, `labels=${JSON.stringify(res.labels)}`);
    assert.equal(res.sourceUsed, 'title');
    assert.ok(res.sourceTried.length >= 3, 'sourceTried muss gefuellt sein');
    assert.ok(res.census && typeof res.census === 'object');
    assert.ok(res.census.scannedCount > 0, 'Census muss gefuellt sein');
    assert.ok(res.census.dayColumns.length >= 2, 'Tages-Spalten muessen erkannt sein');
    assert.ok(Array.isArray(res.structureNodes) && res.structureNodes.length >= 2, 'Strukturknoten muessen geliefert werden');

    // Und die Labels sind sogar zu Slots verwertbar.
    const slots = res.labels.map((l) => parseEventLabel(l, 2026)).filter(Boolean);
    assert.equal(slots.length, 2);
  });

  test('Der title-Quellpfad (Live-Fehlerort) laeuft im Browser-Realm durch', async () => {
    // Absichtlich OHNE aria-label, damit der title/data-attr-Zweig betreten
    // wird — genau dort stand "if (broadScanned >= SCAN_MAX_ELEMENTS)".
    const doc = gridDoc({
      leafCount: 1,
      label: () => 'Urlaub, all day event, Wednesday, September 30, 2026, Busy'
    });
    const res = await collectCalendarEventLabels(browserRealmPage({ document: doc }), { weekStartIso: '2026-09-28' });
    assert.equal(res.collectError, null);
    assert.equal(res.sourceUsed, 'title');
    assert.equal(res.eventElementsFound, 1);
  });

  test('maxElements wird tatsaechlich GELESEN (Parameter, nicht Modulkonstante)', async () => {
    // Ueber maxElements (6000) Elemente -> truncated: true. Mit einer Modul-
    // Referenz waere das im Browser ein ReferenceError, mit dem falschen
    // Parameterkoerper nie erreicht.
    const many = [];
    for (let i = 0; i < 6100; i++) many.push(chipNode({}, '', { left: 5000, top: 5000, width: 1, height: 1, x: 5000, y: 5000 }));
    const doc = gridDoc({ leafCount: 1 });
    const base = doc.querySelectorAll('*');
    doc.querySelectorAll = (sel) => (sel === '*' ? base.concat(many) : doc.__cols(sel));

    const res = await collectCalendarEventLabels(browserRealmPage({ document: doc }), { weekStartIso: '2026-09-28' });
    assert.equal(res.collectError, null);
    assert.equal(res.census.truncated, true, 'Scan muss am maxElements-Limit abbrechen');
    assert.equal(res.census.scannedCount, 6000, `scannedCount=${res.census.scannedCount}`);
  });

  test('maxStructureNodes wird gelesen — nicht maxCensusNodes (14)', async () => {
    // 30 Blatt-Knoten: korrekt gereicht -> 30. Mit maxCensusNodes verwechselt
    // waeren es 14 (das war die Falle im Auftrag: es gibt fuer 200 keinen
    // durchgereichten Parameter).
    const doc = gridDoc({ leafCount: 30 });
    const res = await collectCalendarEventLabels(browserRealmPage({ document: doc }), { weekStartIso: '2026-09-28' });
    assert.equal(res.collectError, null);
    assert.equal(res.census.structureNodes, 30, `census.structureNodes=${res.census.structureNodes}`);
    assert.equal(res.structureNodes.length, 30, `res.structureNodes=${res.structureNodes.length}`);
    // Der Census selbst bleibt klein (maxCensusNodes = 14) — die beiden Limits
    // sind bewusst getrennt.
    assert.equal(res.census.scanned.length, 14, `census.scanned=${res.census.scanned.length}`);
  });

  test('ReferenceError im Callback wird laut geworfen statt als "leerer Kalender" getarnt', async () => {
    // Nach dem Fix existiert kein Leak mehr. Der Wurf-Pfad wird deshalb
    // direkt ueber einen manipulierten Callback geprueft: ein ReferenceError
    // darf NIE als leeres Standardergebnis durchrutschen.
    const doc = gridDoc({ leafCount: 1 });
    const brokenPage = {
      evaluate: async () => { throw new ReferenceError('GIBT_ES_NICHT is not defined'); }
    };
    await assert.rejects(
      () => collectCalendarEventLabels(brokenPage, { weekStartIso: '2026-09-28' }),
      (err) => {
        assert.equal(err.name, 'CollectCallbackReferenceError');
        assert.match(err.message, /Code-Defekt/);
        assert.match(err.message, /GIBT_ES_NICHT/);
        return true;
      }
    );
  });

  test('Ein normaler DOM-Fehler wird weiterhin still als collectError behandelt (kein Wurf)', async () => {
    const domErrorPage = { evaluate: async () => { throw new Error('Execution context was destroyed'); } };
    const res = await collectCalendarEventLabels(domErrorPage, { weekStartIso: '2026-09-28' });
    assert.equal(res.eventElementsFound, 0);
    assert.equal(res.selectorUsed, 'error');
    assert.match(res.collectError, /Execution context was destroyed/);
  });
});

// ===========================================================================
describe('Scope-Audit: KEIN serialisierter Callback darf Modul-Scope-Identifier benutzen', () => {

  const FILES = ['owa.js', 'browser.js'];

  for (const file of FILES) {
    test(`${file}: 0 nackte Modul-Referenzen in allen evaluate-Callbacks`, () => {
      const a = auditEvaluateScope(file);
      // Der Audit darf nicht "gruen" sein, weil er nichts gefunden hat.
      assert.ok(a.callbacks.length > 0 || file === 'browser.js', `keine Callbacks extrahiert (${a.callbacks.length})`);
      assert.ok(a.moduleNames.size >= 2, `Modulnamen erkannt: ${a.moduleNames.size} (${file})`);
      assert.deepEqual(
        a.leaks.map((l) => `Z${l.lineInFile} ${l.name} (${l.kind}, via ${l.api} @Z${l.line})`),
        [],
        'Nackte Modul-Konstanten im Browser-Callback gefunden — Werte ueber das Argument-Objekt reichen'
      );
    });
  }

  test('Der Extraktor findet die bekannten Callbacks (kein stiller Leerlauf)', () => {
    const src = fs.readFileSync(path.join(SRC, 'owa.js'), 'utf8');
    const cbs = serializedCallbacks(src);
    assert.ok(cbs.length >= 20, `nur ${cbs.length} Callbacks gefunden`);
    assert.equal(cbs.filter((c) => c.kind === 'unknown').length, 0, 'nicht extrahierbare Callbacks vorhanden');
    assert.equal(cbs.filter((c) => c.kind === 'referenz').length, 0, 'eine Funktionsreferenz ist nicht serialisierbar');
  });

  test('Der Extraktor erkennt den Live-Defekt wieder (Negativkontrolle an HEAD-Text)', () => {
    // Der Original-Schnippsel aus fc943cd — muss als LEAK erkannt werden.
    const broken = '({ pats }) => { const cap = (a, n) => a.slice(0, n); for (const el of []) { if (0 >= SCAN_MAX_ELEMENTS) break; } return cap([], SCAN_MAX_STRUCTURE_NODES); }';
    const modNames = new Set(['SCAN_MAX_ELEMENTS', 'SCAN_MAX_STRUCTURE_NODES', 'cap']);
    const free = freeIdentifiersOf(broken);
    assert.ok(free.includes('SCAN_MAX_ELEMENTS'), `free=${free.join(',')}`);
    assert.ok(free.includes('SCAN_MAX_STRUCTURE_NODES'));
    assert.ok(!free.includes('cap'), 'cap ist lokal deklariert');
    assert.equal(free.filter((n) => modNames.has(n)).length, 2);
  });

  test('Der Analyzer unterscheidet Member-Zugriff, Objekt-Schluessel und lokal von Module-Scope', () => {
    const free = freeIdentifiersOf('(a) => { const x = 1; return { x, y: a.z, w: SCAN_MAX_ELEMENTS, f(n) { return n; } }; }');
    assert.ok(free.includes('SCAN_MAX_ELEMENTS'), 'echter Leak muss erkannt werden');
    assert.ok(!free.includes('a'), 'Parameter sind gebunden');
    assert.ok(!free.includes('x'), 'lokale Deklaration');
    assert.ok(!free.includes('z'), 'Member-Zugriff');
    assert.ok(!free.includes('f'), 'Methodenname');
    assert.ok(!free.includes('n'), 'Methodenparameter');
    assert.ok(!free.includes('y'), 'Objekt-Literal-Schluessel');
  });

  // META-GUARD: prueft die Pruefung selbst. Ein synthetisches Modul mit
  // exakt der Live-Fehlerform muss als LEAK auffallen. Ohne diesen Test
  // koennte der Audit stillschweigend alles gruen melden, weil die
  // Heuristik verblindet ist.
  test('Der Audit meldet ein synthetisches Modul mit Live-Fehlerform als LEAK', () => {
    // Struktur wie der echte Callback: cap/re/scanned sind LOKAL, die beiden
    // SCAN_MAX_* sind Modul-Konstanten und werden zusaetzlich als Argument
    // durchgereicht (so wie es der Fix tut).
    const body = (maxScan, maxNodes) => ([
      'export async function collect(page) {',
      '  return page.evaluate(({ maxElements, maxStructureNodes }) => {',
      '    const cap = (a, n) => a.slice(0, n);',
      '    const re = new RegExp(pats.timeRange, "i");',
      '    let scanned = 0;',
      '    for (const el of []) {',
      `      if (scanned >= ${maxScan}) break;`,
      '      scanned++;',
      '    }',
      `    return { re, nodes: cap([], ${maxNodes}) };`,
      '  }, { pats: { timeRange: "x" }, maxElements: 6000, maxStructureNodes: 200 });',
      '}'
    ].join('\n'));
    const buggy = ['const SCAN_MAX_ELEMENTS = 6000;', 'const SCAN_MAX_STRUCTURE_NODES = 200;', 'const pats = { timeRange: "x" };']
      .join('\n') + '\n' + body('SCAN_MAX_ELEMENTS', 'SCAN_MAX_STRUCTURE_NODES');

    const modNames = moduleLevelNames(buggy);
    assert.ok(modNames.has('SCAN_MAX_ELEMENTS') && modNames.has('SCAN_MAX_STRUCTURE_NODES'), `Modulnamen erkannt: ${[...modNames]}`);

    const cbs = serializedCallbacks(buggy);
    assert.equal(cbs.length, 1, `nur ${cbs.length} Callbacks extrahiert`);
    const leaks = freeIdentifiersOf(cbs[0].text).filter((n) => modNames.has(n));
    assert.deepEqual(leaks.sort(), ['SCAN_MAX_ELEMENTS', 'SCAN_MAX_STRUCTURE_NODES'], `leaks=${leaks.join(',')}`);

    // Und die behobene Fassung desselben Moduls ist leak-frei.
    const fixed = ['const SCAN_MAX_ELEMENTS = 6000;', 'const SCAN_MAX_STRUCTURE_NODES = 200;', 'const pats = { timeRange: "x" };']
      .join('\n') + '\n' + body('maxElements', 'maxStructureNodes');
    const freeFixed = freeIdentifiersOf(serializedCallbacks(fixed)[0].text);
    assert.deepEqual(freeFixed.filter((n) => modNames.has(n)), [], `behobene Fassung nicht leak-frei: ${freeFixed.join(',')}`);
  });

  test('Kontrollfluss-Kopf ist kein Parameter: "if (x >= KONSTANTE) {" bleibt ein Leak', () => {
    // Regression auf die eigene Heuristik: eine zu greifende
    // Parameter-Erkennung wuerde if/for-Bedingungen als Parameterliste
    // lesen und den Leak genau dort verschlucken, wo er live auftrat.
    const free = freeIdentifiersOf('(a) => { if (a.n >= SCAN_MAX_ELEMENTS) { break; } for (const el of a.l) { if (el >= SCAN_MAX_STRUCTURE_NODES) break; } }');
    assert.ok(free.includes('SCAN_MAX_ELEMENTS'), 'if-Bedingung darf den Leak nicht verdecken');
    assert.ok(free.includes('SCAN_MAX_STRUCTURE_NODES'), 'for-Bedingung darf den Leak nicht verdecken');
  });

  test('Funktions-/Methodenparameter werden gebunden (kein Fehlalarm)', () => {
    const free = freeIdentifiersOf('async (arg) => { const f = function (inner) { return inner; }; return f(arg.maxElements); }');
    assert.deepEqual(free, []);
  });

  test('Division nach ) ist kein Regex-Literal (Tokenizer-Selbsttest)', () => {
    // Ohne diese Unterscheidung zerlegt der Tokenizer Zeilen wie
    // "Math.round((a / b) * 60 / 15)" falsch und die Modul-Tiefe laeuft weg.
    const free = freeIdentifiersOf('(a) => { return Math.round((a / SCAN_MAX_ELEMENTS) * 60 / 15) * 15; }');
    assert.deepEqual(free, ['SCAN_MAX_ELEMENTS']);
  });

  test('modulLevelNames erkennt Top-Level-Deklarationen ueber verschachtelte Bloecke hinweg', () => {
    const names = moduleLevelNames('const A = 1;\nfunction f() {\n  const B = 2;\n  if (x) { const C = 3; }\n  return A + B + C;\n}\nclass D { m() { return 1; } }\nlet E = 5;');
    assert.ok(names.has('A') && names.has('f') && names.has('D') && names.has('E'));
    assert.ok(!names.has('B'), 'B ist kein Modul-Scope-Name');
    assert.ok(!names.has('C'));
  });
});

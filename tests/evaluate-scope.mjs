// ===========================================================================
// evaluate-scope.mjs — Test-Infrastruktur fuer die Fehlerklasse
// "Identifier-Leak ueber die page.evaluate-Grenze".
//
// PROBLEMKLASSE (Live-Befund 25.09.2026, src/owa.js Z973):
//   Ein an page.evaluate uebergebener Callback laeuft NICHT im Node-Modul-Scope,
//   sondern im Browser-Realm. Playwright serialisiert ihn dazu als QUELLTEXT
//   (fn.toString()) und wertet ihn dort aus. Jede Referenz auf eine
//   modulweite Konstante (SCAN_MAX_ELEMENTS o. Ae.) ist im Browser ein
//   ReferenceError — und NICHT ein "0 Treffer".
//
// WARUM DER BISHERIGE TEST-FAKE BLIND WAR:
//   const fakePage = { evaluate: async (fn, arg) => fn(arg) };
//   -> der Callback laeuft im selben Modul-Scope wie owa.js, in dem
//      SCAN_MAX_ELEMENTS normal aufloest. Ein Scope-Leak ist dort per
//      Konstruktion unsichtbar. Genau darum war der Testlauf 254/254 gruen,
//      obwohl der Live-Lauf mit ReferenceError abbrach.
//
// WIE HIER GEPRUEFT WIRD (zwei unabhaengige Mechanismen):
//   1) TRANSPORTGETREUE EMULATION (runtime): browserRealmPage() uebergibt den
//      Callback als Quelltext in einen FRISCHEN vm.Realm, in dem die
//      Modul-Konstanten nicht existieren. Das ist genau die Semantik von
//      Playwright. Ein Leak wird hier als ReferenceError sichtbar.
//   2) STATISCHE ANALYSE (alle Callbacks, ohne Browser): extrahiert JEDEN
//      serialisierten Callback aus der Quelle und meldet jeden freien
//      Identifier, der ein Modul-Top-Level-Name ist. Das ist die
//      vollstaendige Abdeckung — unabhaengig davon, ob ein Codepfad im Test
//      ueberhaupt erreicht wird.
// ===========================================================================
import vm from 'node:vm';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
export const SRC_DIR = path.join(HERE, '..', 'src');

/**
 * emulateBrowserEvaluate — der Kern von Mechanismismus 1.
 * Führt einen page.evaluate-Callback so aus, wie Playwright es tut:
 * als Quelltext in einem fremden Realm.
 */
export function emulateBrowserEvaluate(fn, arg, browserGlobals = {}) {
  const src = typeof fn === 'function' ? fn.toString() : String(fn);
  const ctx = vm.createContext({ document: undefined, window: undefined, console, ...browserGlobals });
  const browserFn = new vm.Script(`(${src})`, { filename: 'page-evaluate-callback.js' }).runInContext(ctx);
  return browserFn(arg);
}

/**
 * browserRealmPage — page-Ersatz, dessen evaluate den Browser-Transport
 * emuliert. Minimale Playwright-Oberflaeche, die der Kalenderpfad braucht.
 * @param {object} opts.document  Fake-DOM fuer den Browser-Realm
 */
export function browserRealmPage({ document: doc = undefined, onEvaluate = null } = {}) {
  const evaluateCalls = [];
  return {
    evaluateCalls,
    async evaluate(fn, arg) {
      evaluateCalls.push(typeof fn === 'function' ? fn.toString() : String(fn));
      const out = emulateBrowserEvaluate(fn, arg, { document: doc });
      if (onEvaluate) await onEvaluate();
      return out;
    },
    async waitForTimeout() {},
    url() { return 'https://outlook.office.com/calendar/view/month'; },
    locator() { throw new Error('locator() ist im evaluate-Scope-Test nicht unterstuetzt'); }
  };
}

// ---------------------------------------------------------------------------
// Mechanismus 2: statische Analyse aller serialisierten Callbacks
// ---------------------------------------------------------------------------
const KEYWORDS = new Set([
  'break', 'case', 'catch', 'class', 'const', 'continue', 'debugger', 'default', 'delete', 'do',
  'else', 'export', 'extends', 'finally', 'for', 'function', 'if', 'import', 'in', 'instanceof',
  'let', 'new', 'of', 'return', 'static', 'super', 'switch', 'this', 'throw', 'try', 'typeof',
  'var', 'void', 'while', 'with', 'yield', 'await', 'async', 'get', 'set', 'null', 'true', 'false'
]);

// Browser-/JS-Globals: die DUKTION ist absichtlich eine Positivliste. Alles,
// was hier fehlt, gilt als potentieller Scope-Leak-Kandidat.
const BROWSER_GLOBALS = new Set([
  'document', 'window', 'navigator', 'location', 'history', 'screen', 'console',
  'getComputedStyle', 'setTimeout', 'clearTimeout', 'setInterval', 'clearInterval',
  'requestAnimationFrame', 'cancelAnimationFrame', 'requestIdleCallback',
  'Element', 'HTMLElement', 'Node', 'NodeList', 'Event', 'CustomEvent', 'URL',
  'MutationObserver', 'ResizeObserver', 'IntersectionObserver', 'PerformanceObserver',
  'fetch', 'localStorage', 'sessionStorage', 'self', 'top', 'parent', 'frames',
  'innerWidth', 'innerHeight', 'devicePixelRatio', 'matchMedia', 'alert', 'CSS',
  'Array', 'Object', 'String', 'Number', 'Boolean', 'RegExp', 'Date', 'Math', 'JSON',
  'Set', 'Map', 'WeakMap', 'WeakSet', 'Promise', 'Symbol', 'BigInt', 'Error', 'TypeError',
  'RangeError', 'SyntaxError', 'ReferenceError', 'Function', 'Proxy', 'Reflect', 'Intl',
  'ArrayBuffer', 'Uint8Array', 'Int32Array', 'Float64Array', 'isNaN', 'isFinite',
  'parseInt', 'parseFloat', 'encodeURIComponent', 'decodeURIComponent', 'encodeURI', 'decodeURI',
  'undefined', 'NaN', 'Infinity', 'globalThis', 'structuredClone', 'queueMicrotask'
]);

const DECL_KEYWORDS = new Set(['const', 'let', 'var']);
const CONTROL_HEADS = new Set(['if', 'while', 'for', 'with', 'switch', 'catch']);

/** Minimaler JS-Tokenizer: Strings, Templates, Regex-Literale und Kommentare korrekt uebersprungen. */
export function tokenize(src) {
  const toks = [];
  const parenHeads = [];
  let i = 0;
  let prev = null;
  const n = src.length;
  const regexAllowed = () => {
    if (!prev) return true;
    if (prev.type === 'num' || prev.type === 'str' || prev.type === 'template' || prev.type === 'regex') return false;
    if (prev.type === 'name') return KEYWORDS.has(prev.value);
    if (prev.value === ')' || prev.value === ']' || prev.value === '}') {
      // "if (x) /re/" ist ein Regex, "(a + b) / 2" ist eine Division.
      if (prev.value !== ')') return false;
      return parenHeads.length > 0 && CONTROL_HEADS.has(parenHeads[parenHeads.length - 1]);
    }
    if (prev.value === '++' || prev.value === '--') return false;
    return true;
  };
  const push = (type, value, pos) => { const t = { type, value, pos }; toks.push(t); prev = t; return t; };
  while (i < n) {
    const c = src[i];
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r' || c === '\f' || c === '\v') { i++; continue; }
    if (c === '/' && src[i + 1] === '/') { while (i < n && src[i] !== '\n') i++; continue; }
    if (c === '/' && src[i + 1] === '*') { const e = src.indexOf('*/', i + 2); i = e < 0 ? n : e + 2; continue; }
    if (c === "'" || c === '"') {
      const s = i; const q = c; i++;
      while (i < n) { if (src[i] === '\\') i += 2; else if (src[i] === q) { i++; break; } else i++; }
      push('str', src.slice(s, i), s); continue;
    }
    if (c === '`') {
      const s = i; i++; let d = 0;
      while (i < n) {
        if (src[i] === '\\') { i += 2; continue; }
        if (src[i] === '`' && d === 0) { i++; break; }
        if (src[i] === '$' && src[i + 1] === '{') { d++; i += 2; continue; }
        if (src[i] === '}' && d > 0) { d--; i++; continue; }
        i++;
      }
      push('template', src.slice(s, i), s); continue;
    }
    if (c === '/' && regexAllowed()) {
      const s = i; i++; let cls = false; let ok = false;
      while (i < n) {
        const ch = src[i];
        if (ch === '\\') { i += 2; continue; }
        if (ch === '\n') break;
        if (ch === '[') cls = true;
        else if (ch === ']') cls = false;
        else if (ch === '/' && !cls) { i++; ok = true; break; }
        i++;
      }
      if (ok) { while (i < n && /[a-z]/.test(src[i])) i++; push('regex', src.slice(s, i), s); continue; }
      i = s;
    }
    if (/[0-9]/.test(c) || (c === '.' && /[0-9]/.test(src[i + 1] || ''))) {
      const s = i; i++;
      while (i < n && /[0-9a-fA-FxXoObBn._]/.test(src[i])) i++;
      push('num', src.slice(s, i), s); continue;
    }
    if (/[A-Za-z_$]/.test(c)) {
      const s = i; i++;
      while (i < n && /[\w$]/.test(src[i])) i++;
      push('name', src.slice(s, i), s); continue;
    }
    const three = src.slice(i, i + 3);
    const two = src.slice(i, i + 2);
    if (['...', '===', '!==', '**=', '<<=', '>>=', '&&=', '||=', '??=', '>>>'].includes(three)) { push('punct', three, i); i += 3; continue; }
    if (['=>', '==', '!=', '<=', '>=', '&&', '||', '??', '?.', '++', '--', '+=', '-=', '*=', '/=', '%=', '**', '<<', '>>'].includes(two)) { push('punct', two, i); i += 2; continue; }
    if ('{}()[];,.:?=+-*/%<>!&|^~'.includes(c)) {
      let head = null;
      if (c === '(') {
        head = prev && prev.type === 'name' ? prev.value : null;
        parenHeads.push(head);
      } else if (c === ')') {
        head = parenHeads.pop();
      }
      const t = push('punct', c, i);
      // Der Kopf dieser Klammer-Liste: 'if'/'for'/'while' = Kontrollfluss,
      // null = Pfeil-Parameter, sonst Funktions-/Methodenname. Ohne das ist
      // "if (x >= KONSTANTE) {" von einem Funktionskopf nicht unterscheidbar.
      t.parenHead = head;
      i++;
      continue;
    }
    i++;
  }
  return toks;
}

function collectPattern(toks, start, names) {
  let depth = 0;
  for (let k = start; k < toks.length; k++) {
    const t = toks[k];
    if (['{', '[', '('].includes(t.value)) depth++;
    else if (['}', ']', ')'].includes(t.value)) { if (depth === 0) return; depth--; }
    else if (depth === 0 && (t.value === ',' || t.value === ';')) return;
    else if (t.type === 'name' && !KEYWORDS.has(t.value) && t.value !== 'as') {
      const nxt = toks[k + 1];
      if (nxt && nxt.value === ':') continue;               // Objekt-Muster-Schluessel
      if (toks[k - 1] && toks[k - 1].value === '.') continue; // Member
      names.add(t.value);
    }
  }
}

/** Alle Deklarationsnamen im Callback (flach, ueber Hoisting hinweg). */
function declaredNames(toks) {
  const names = new Set();
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type === 'name' && DECL_KEYWORDS.has(t.value)) {
      let k = i + 1;
      while (k < toks.length && toks[k].value === ',') {
        collectPattern(toks, k + 1, names); k++;
        while (k < toks.length && !['=', ',', ';'].includes(toks[k].value)) k++;
      }
      collectPattern(toks, k, names);
      continue;
    }
    if (t.type === 'name' && (t.value === 'function' || t.value === 'class') && toks[i + 1] && toks[i + 1].type === 'name') {
      names.add(toks[i + 1].value);
    }
  }
  // Parameter einer Klammer-Liste, die von einem Funktions-/Methodenkoepfe
  // oder einem Pfeil gefolgt wird: (x) =>, function (x) {}, { m(x) {} }.
  // Kontrollfluss-Köpfe (if/for/while/with/switch/catch) sind ausdruecklich
  // AUSGENOMMEN — "if (x >= KONSTANTE) {" ist kein Parameter, sondern genau
  // die Referenz, die wir suchen.
  for (let i = 0; i < toks.length; i++) {
    if (toks[i].value !== ')') continue;
    const nxt = toks[i + 1];
    const isArrow = nxt && nxt.value === '=>';
    const isBody = nxt && nxt.value === '{';
    if (!isArrow && !isBody) continue;
    if (CONTROL_HEADS.has(toks[i].parenHead)) continue;
    let depth = 0;
    for (let k = i; k >= 0; k--) {
      if (toks[k].value === ')') depth++;
      else if (toks[k].value === '(') {
        depth--;
        if (depth === 0) {
          for (let m = k + 1; m < i; m++) {
            if (toks[m].type === 'name' && !KEYWORDS.has(toks[m].value)) names.add(toks[m].value);
          }
          break;
        }
      }
    }
  }
  return names;
}

/** Freie Identifier eines Callback-Quelltexts (lokal + Browser-Globals abgezogen). */
export function freeIdentifiersOf(fnText) {
  const toks = tokenize(fnText);
  const declared = declaredNames(toks);
  const free = [];
  const seen = new Set();
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'name') continue;
    const v = t.value;
    if (KEYWORDS.has(v) || BROWSER_GLOBALS.has(v) || declared.has(v)) continue;
    const p = toks[i - 1];
    const n = toks[i + 1];
    if (p && (p.value === '.' || p.value === '?.')) continue;      // Member-Zugriff
    // Schluessel eines Objekt-Literals ({ k: v }, { k, v }, { k() {} }) sind
    // keine Referenzen. { k } (Shorthand) IST eine und bleibt deshalb erfasst.
    if (p && (p.value === '{' || p.value === ',') && n && (n.value === ':' || n.value === ',' || n.value === '(')) continue;
    if (n && n.value === ':' && p && p.value !== '.') continue;      // Label
    if (!seen.has(v)) { seen.add(v); free.push(v); }
  }
  return free;
}

/** Modul-Top-Level-Namen (const/let/var/function/class) einer Quelldatei. */
export function moduleLevelNames(src) {
  const toks = tokenize(src);
  const names = new Set();
  let depth = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (depth === 0 && t.type === 'name' && (DECL_KEYWORDS.has(t.value) || t.value === 'function' || t.value === 'class')) {
      const n = toks[i + 1];
      if (n && n.type === 'name') names.add(n.value);
    }
    if (t.type === 'punct') {
      if ('{([('.includes(t.value)) depth++;
      else if ('}])'.includes(t.value)) depth--;
    }
  }
  return names;
}

// ----------------------- Callback-Extraktion aus dem Quelltext ---------------
const CALL_APIS = ['.evaluateHandle(', '.evaluate(', '.waitForFunction(', '.$$eval(', '.$eval('];

function scanPair(s, i, open, close) {
  if (i < 0 || i >= s.length || s[i] !== open) return -1;
  let d = 0;
  for (let k = i; k < s.length; k++) {
    const c = s[k];
    if (c === '\\') { k++; continue; }
    if (c === "'" || c === '"' || c === '`') { const q = c; k++; while (k < s.length) { if (s[k] === '\\') k++; else if (s[k] === q) break; k++; } continue; }
    if (c === '/' && s[k + 1] === '*') { const e = s.indexOf('*/', k + 2); if (e < 0) return -1; k = e + 1; continue; }
    if (c === '/' && s[k + 1] === '/') { while (k < s.length && s[k] !== '\n') k++; continue; }
    if (c === open) d++;
    else if (c === close) { d--; if (d === 0) return k; }
  }
  return -1;
}

function splitTopLevelArgs(s, openIdx) {
  let d = 0; const out = []; let cur = openIdx + 1;
  for (let k = openIdx; k < s.length; k++) {
    const c = s[k];
    if (c === '\\') { k++; continue; }
    if (c === "'" || c === '"' || c === '`') { const q = c; k++; while (k < s.length) { if (s[k] === '\\') k++; else if (s[k] === q) break; k++; } continue; }
    if (c === '/' && s[k + 1] === '*') { const e = s.indexOf('*/', k + 2); if (e < 0) break; k = e + 1; continue; }
    if (c === '/' && s[k + 1] === '/') { while (k < s.length && s[k] !== '\n') k++; continue; }
    if (c === '(' || c === '[' || c === '{') d++;
    else if (c === ')' || c === ']' || c === '}') { d--; if (d === 0) { out.push(s.slice(cur, k)); return out; } }
    else if (c === ',' && d === 1) { out.push(s.slice(cur, k)); cur = k + 1; }
  }
  return out;
}

function parseCallbackArg(arg) {
  const a = arg.trim();
  let j = 0;
  const asy = /^async\s+/.exec(a);
  if (asy) j += asy[0].length;
  if (a.startsWith('function', j)) return { kind: 'function', text: a.slice(j) };
  if (a[j] === '(') {
    const pe = scanPair(a, j, '(', ')');
    if (pe < 0) return null;
    const af = a.slice(pe + 1).match(/^\s*=>/);
    if (!af) return null;
    const bs = pe + 1 + af[0].length;
    if (a[bs] === '{') { const be = scanPair(a, bs, '{', '}'); return be < 0 ? null : { kind: 'arrow', text: a.slice(j, be + 1) }; }
    return { kind: 'arrow-expr', text: a.slice(j) };
  }
  const idm = /^[A-Za-z_$][\w$]*/.exec(a.slice(j));
  if (idm) {
    const nm = idm[0];
    const af = a.slice(j + nm.length).match(/^\s*=>/);
    if (af) {
      const bs = j + nm.length + af[0].length;
      if (a[bs] === '{') { const be = scanPair(a, bs, '{', '}'); return be < 0 ? null : { kind: 'arrow-1p', text: a.slice(j, be + 1) }; }
      return { kind: 'arrow-1p-expr', text: a.slice(j) };
    }
    return { kind: 'referenz', name: nm, text: nm }; // Funktionsreferenz: noch schlimmer
  }
  return null;
}

/** Alle serialisierten Callback-Stellen einer Quelldatei. */
export function serializedCallbacks(src) {
  const rows = [];
  for (const api of CALL_APIS) {
    let from = 0;
    for (;;) {
      const at = src.indexOf(api, from);
      if (at < 0) break;
      from = at + api.length;
      const line = src.slice(0, at).split('\n').length;
      const args = splitTopLevelArgs(src, at + api.length - 1);
      let picked = null;
      for (const a of args) { const p = parseCallbackArg(a); if (p) { picked = p; break; } }
      if (!picked) { rows.push({ line, api, kind: 'unknown', text: '', offset: at }); continue; }
      const offset = src.indexOf(picked.text, at);
      rows.push({ line, api, kind: picked.kind, name: picked.name || null, text: picked.text, offset: offset >= 0 ? offset : at });
    }
  }
  return rows.sort((a, b) => a.line - b.line);
}

/**
 * Scope-Audit fuer eine Quelldatei.
 * @returns {{ file, moduleNames, callbacks, leaks: Array }}
 *   leaks: { line, api, name, lineInFile } — nackte Modul-Referenz im Callback
 */
export function auditEvaluateScope(file) {
  const abs = path.join(SRC_DIR, file);
  const src = fs.readFileSync(abs, 'utf8');
  const moduleNames = moduleLevelNames(src);
  const callbacks = serializedCallbacks(src);
  const leaks = [];
  for (const cb of callbacks) {
    if (cb.kind === 'referenz') {
      if (moduleNames.has(cb.name)) leaks.push({ line: cb.line, api: cb.api, name: cb.name, lineInFile: cb.line, kind: 'funktionsreferenz' });
      continue;
    }
    if (!cb.text) continue;
    for (const name of freeIdentifiersOf(cb.text)) {
      if (!moduleNames.has(name)) continue;
      const re = new RegExp(`\\b${name}\\b`, 'g');
      const lines = [];
      let m;
      while ((m = re.exec(cb.text))) lines.push(cb.line + cb.text.slice(0, m.index).split('\n').length);
      for (const lineInFile of lines) leaks.push({ line: cb.line, api: cb.api, name, lineInFile, kind: 'nackte-referenz' });
    }
  }
  return { file, moduleNames, callbacks, leaks };
}

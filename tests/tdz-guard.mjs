// ===========================================================================
// tdz-guard.mjs — erkennt "Use-before-declaration" (TDZ) fuer const/let.
//
// Fehlerklasse aus fc943cd (25.09.2026): beim Refactor sind Statement-Bloecke
// vor ihre eigene const-Deklaration gerutscht. Zwei Instanzen, beide in
// getFreeBusy:
//   1) der `let note = '...' +`-Verkettungsstring wurde mitten im Satz
//      zerrissen -> ESM-SyntaxError, `node index.js` startet nicht
//   2) `rawLabels`/`eventElementsFound`/`selectorUsed` wurden aus
//      `eventCollection` gelesen, BEVOR dessen const-Deklaration stand
//      -> ReferenceError: Cannot access 'eventCollection' before initialization
//
// Beide fielen nur nicht auf, weil Live-Lauf und Tests gegen den
// Arbeitsbaum liefen, in dem die Reihenfolge schon korrekt war.
//
// Kriterium: eine Referenz auf N liegt VOR der Deklaration von N, bei
// GLEICHER Block-Tiefe, und zwischen beiden wird keine Blockgrenze
// unterschritten. Runde () zaehlen bewusst NICHT als Tiefe — Funktions-
// koepfe wuerden sonst als Blockgrenze taeuschen.
// ===========================================================================
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { tokenize } from './evaluate-scope.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
export const SRC_DIR = path.join(HERE, '..', 'src');

const KEYWORD = new Set([
  'if', 'else', 'for', 'while', 'do', 'switch', 'case', 'return', 'function', 'class', 'new',
  'typeof', 'void', 'delete', 'in', 'of', 'try', 'catch', 'finally', 'throw', 'break',
  'continue', 'await', 'async', 'export', 'import', 'default', 'extends', 'this', 'super',
  'null', 'true', 'false', 'get', 'set', 'static', 'const', 'let', 'var', 'instanceof',
  'yield', 'debugger', 'with'
]);

/** Namen, die als Parameter gebunden werden (Pfeil, Funktions-koerper, Methode). */
function parameterNames(toks) {
  const names = new Set();
  for (let i = 0; i < toks.length; i++) {
    if (toks[i].value !== ')') continue;
    const nxt = toks[i + 1];
    if (!(nxt && (nxt.value === '=>' || nxt.value === '{'))) continue;
    if (['if', 'for', 'while', 'switch', 'catch', 'with'].includes(toks[i].parenHead)) continue;
    let depth = 0;
    for (let k = i; k >= 0; k--) {
      if (toks[k].value === ')') depth++;
      else if (toks[k].value === '(') {
        depth--;
        if (depth === 0) {
          for (let m = k + 1; m < i; m++) if (toks[m].type === 'name' && !KEYWORD.has(toks[m].value)) names.add(toks[m].value);
          break;
        }
      }
    }
  }
  return names;
}

/**
 * TDZ-Stellen in einem Quelltext.
 * @param {string} src
 * @returns {Array<{ name, useLine, declLine }>}
 */
export function findTdzInSource(src) {
  const toks = tokenize(src);
  const lineOf = (i) => src.slice(0, i).split('\n').length;

  // Block-Tiefe: nur {} zaehlen.
  const depth = new Int32Array(toks.length + 1);
  let d = 0;
  for (let i = 0; i < toks.length; i++) {
    if (toks[i].type === 'punct' && toks[i].value === '{') { depth[i] = ++d; continue; }
    if (toks[i].type === 'punct' && toks[i].value === '}') { depth[i] = --d; continue; }
    depth[i] = d;
  }

  const decls = new Map();
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'name' || (t.value !== 'const' && t.value !== 'let')) continue;
    const n = toks[i + 1];
    if (!n || n.type !== 'name') continue;
    const arr = decls.get(n.value) || [];
    arr.push({ i: i + 1, depth: depth[i] });
    decls.set(n.value, arr);
  }

  // Parameter schatten gleichnamige Deklarationen in aeusseren Bloecken.
  const params = parameterNames(toks);

  const hits = [];
  const seen = new Set();
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'name' || KEYWORD.has(t.value)) continue;
    const p = toks[i - 1];
    if (p && (p.value === '.' || p.value === '?.')) continue;
    if (p && (p.value === '{' || p.value === ',') && toks[i + 1] && toks[i + 1].value === ':') continue;
    const arr = decls.get(t.value);
    if (!arr) continue;
    for (const decl of arr) {
      if (decl.i <= i) continue;
      if (decl.depth !== depth[i]) continue;
      let closed = false;
      for (let k = i; k < decl.i; k++) {
        if (toks[k].type === 'punct' && toks[k].value === '}') { closed = true; break; }
      }
      if (closed) continue;
      // Parametername: die aeussere Deklaration ist hier nicht die Bindung.
      if (params.has(t.value)) break;
      const key = `${t.value}@${lineOf(t.pos)}`;
      if (seen.has(key)) break;
      seen.add(key);
      hits.push({ name: t.value, useLine: lineOf(t.pos), declLine: lineOf(toks[decl.i].pos) });
      break;
    }
  }
  return hits;
}

/** TDZ-Stellen in einer Quelldatei (relativ zu src/ oder absolut). */
export function findTdzSites(file) {
  const abs = path.isAbsolute(file) ? file : path.join(SRC_DIR, file);
  return findTdzInSource(fs.readFileSync(abs, 'utf8'));
}

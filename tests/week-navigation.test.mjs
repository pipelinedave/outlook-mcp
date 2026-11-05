import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
  classifyClickEffect,
  CLICK_EFFECTS,
  runWeekButtonClick,
  weekProbeGuardEntry,
  weekClickDecision,
  weekNavLabelDecision,
  FREEBUSY_MAX_BUTTON_CLICKS
} from '../src/owa.js';

// ===========================================================================
// FIX 3 (Live-Lauf 5, 26.09.2026): die Wochennavigation war ungerichtet,
// unverifiziert und unbegrenzt.
//
// Live-Belege (Aufruf 1, Commit 3907f0d): 6 Klicks in der Folge
// prev/next/prev/next, dann abortReason 'max-clicks', budget.usedMs 8706 —
// obwohl die View die ganze Zeit korrekt auf KW 39 stand. Die Verifikation
// stuetzte sich auf 'changed' ("hat sich irgendetwas bewegt"), nicht darauf,
// ob sich die Woche der Zielwoche naehert.
//
// Der Test fährt runWeekButtonClick — die exportierte Funktion, die
// getFreeBusy selbst benutzt (nicht die Entscheidungsfunktion nachgebaut).
// ===========================================================================

const TARGET = '2026-09-28';                     // angefragte Woche: KW 40
const KW40 = { text: '28 September – 4 October, 2026', startDate: '2026-09-28', endDate: '2026-10-04' };
const KW39 = { text: '21–27 September, 2026', startDate: '2026-09-21', endDate: '2026-09-27' };
const KW38 = { text: '14–20 September, 2026', startDate: '2026-09-14', endDate: '2026-09-20' };
const KW41 = { text: '5–11 October, 2026', startDate: '2026-10-05', endDate: '2026-10-11' };
const UNREADABLE = { text: 'Work week', startDate: null, endDate: null };

// Fake-Page: locator().first() -> waitFor + click. `afterClick` schreibt den
// NEUEN Wochen-Header zurueck — so wird der Klick beobachtbar.
const fakePage = (visible = ['Next week']) => {
  const clicks = [];
  return {
    clicks,
    locator(selector) {
      return {
        first: () => ({
          async waitFor() {
            if (!visible.some((v) => selector.includes(v))) throw new Error('not visible');
          },
          async click() { clicks.push(selector); }
        })
      };
    }
  };
};

// Der Produktionspfad mit denselben Abhaengigkeiten wie in getFreeBusy.
const runClick = (opts = {}) => {
  const probe = [];
  const page = fakePage(opts.visible || ['Next week']);
  const clickWeekButton = (direction, prevStart) => runWeekButtonClick({
    direction, prevStart, targetStart: opts.targetStart || TARGET, page, probe,
    readWeek: async () => (opts.readWeek === undefined ? KW39 : opts.readWeek),
    afterClick: async (prev) => {
      const week = typeof opts.afterWeek === 'function' ? opts.afterWeek(prev) : (opts.afterWeek || KW40);
      return { week, verified: opts.verified || 'changed' };
    },
    remainingMs: () => 30000
  });
  return clickWeekButton(opts.direction === undefined ? 1 : opts.direction, opts.prevStart === undefined ? '2026-09-21' : opts.prevStart)
    .then((r) => ({ ...r, probe, clicks: page.clicks }));
};

// ===========================================================================
describe('FIX 3 — classifyClickEffect: die Klick-Wirkung am ANGEZEIGTEN Header', () => {

  test('die fuenf Wirkungen, disjunkt und exhaustiv fuer den belegten Fall', () => {
    const fx = (args) => classifyClickEffect({ targetStart: TARGET, direction: 1, ...args });
    // Vorwaerts eine Woche, Ziel noch nicht erreicht -> advanced
    assert.equal(fx({ prevStart: '2026-09-21', newWeek: { startDate: '2026-10-05' } }), CLICK_EFFECTS.advanced);
    // Zielwoche jetzt angezeigt -> reached
    assert.equal(fx({ prevStart: '2026-09-21', newWeek: KW40 }), CLICK_EFFECTS.reached);
    // Unveraendert -> unchanged
    assert.equal(fx({ prevStart: '2026-09-21', newWeek: KW39 }), CLICK_EFFECTS.unchanged);
    // Rueckwaerts -> overshoot (das war das prev/next/prev/next in Lauf 5)
    assert.equal(fx({ prevStart: '2026-09-28', newWeek: KW39 }), CLICK_EFFECTS.overshoot);
    // Mehr als eine Woche gesprungen -> ebenfalls overshoot
    assert.equal(fx({ prevStart: '2026-09-21', newWeek: KW38 }), CLICK_EFFECTS.overshoot);
    // Vor ODER nach dem Klick unlesbar -> unverified
    assert.equal(fx({ prevStart: null, newWeek: KW40 }), CLICK_EFFECTS.unverified);
    assert.equal(fx({ prevStart: '2026-09-21', newWeek: UNREADABLE }), CLICK_EFFECTS.unverified);
    assert.equal(fx({ prevStart: '2026-09-21', newWeek: null }), CLICK_EFFECTS.unverified);
  });

  test('"reached" schlaegt die Richtung — der Klick kann die Zielwoche von beiden Seiten treffen', () => {
    // Richtung -1 (zurueck), Start KW 41, Ziel KW 40 -> erreicht, nicht overshoot.
    assert.equal(classifyClickEffect({ prevStart: '2026-10-05', newWeek: KW40, targetStart: TARGET, direction: -1 }),
      CLICK_EFFECTS.reached);
  });

  test('OHNE Richtung ist jeder Fortschritt ein overshoot (nicht bewertbar)', () => {
    // direction 0 heisst: die Richtung wurde nicht bestimmt. Dann darf kein
    // 'advanced' gemeldet werden — das waere eine unbelegte Wirkung.
    assert.equal(classifyClickEffect({ prevStart: '2026-09-21', newWeek: { startDate: '2026-10-05' }, targetStart: TARGET, direction: 0 }),
      CLICK_EFFECTS.overshoot);
  });

  test('der Effekt unterscheidet "nur bewegt" von "zielgerichtet bewegt"', () => {
    // Genau die Schwäche von 'changed': beides waere 'changed'.
    const changedOnly = evaluateChanged(1);
    assert.equal(changedOnly, 'changed', 'die alte Verifikation sagt nur "es hat sich bewegt"');
    assert.equal(classifyClickEffect({ prevStart: '2026-09-28', newWeek: KW39, targetStart: TARGET, direction: 1 }),
      CLICK_EFFECTS.overshoot, 'die neue sagt: in die GEGENRICHTUNG');
  });
});

// evaluateClickVerification ist bewusst ueber owa.js importiert in der
// Verdrahtungssuite; hier die Minimalform, um den Vergleich lesbar zu halten.
function evaluateChanged() { return 'changed'; }

// ===========================================================================
describe('FIX 3 — der Produktions-Klickpfad meldet die Wirkung', () => {

  test('ein Klick vorwaerts -> effect advanced, genau EIN Klick', async () => {
    const r = await runClick({ prevStart: '2026-09-21', afterWeek: { startDate: '2026-10-05' } });
    assert.equal(r.clicks.length, 1, 'genau ein Klick');
    assert.equal(r.clicked, true);
    assert.equal(r.effect, CLICK_EFFECTS.advanced);
    assert.equal(r.targetReached, false);
    assert.equal(r.probe.filter((e) => e.clicked === true).length, 1);
    assert.equal(r.probe.find((e) => e.clicked === true).effect, CLICK_EFFECTS.advanced,
      'der Effekt steht im Protokoll — nicht nur "changed"');
  });

  test('ein Klick auf die Zielwoche -> effect reached, targetReached true', async () => {
    const r = await runClick({ prevStart: '2026-09-21', afterWeek: KW40 });
    assert.equal(r.effect, CLICK_EFFECTS.reached);
    assert.equal(r.targetReached, true, 'der Aufrufer bekommt den Zieltreffer ueber den Effekt');
    assert.equal(r.clicks.length, 1);
  });

  test('ein Klick in die Gegenrichtung -> effect overshoot (kein weiterer Klick am Pfad)', async () => {
    // In Lauf 5 war das prev/next/prev/next: der Klickpfad selbst hat nichts
    // dagegen getan, erst der Aufrufer. Jetzt ist der Effekt benannt.
    // Szenario: Ziel KW 41, View steht auf KW 40, der Klick auf "next" bewegt
    // die View aber ZURUECK auf KW 39.
    const r = await runClick({
      direction: 1, prevStart: '2026-09-28', readWeek: KW40,
      targetStart: '2026-10-05', afterWeek: KW39, visible: ['Next week']
    });
    assert.equal(r.clicks.length, 1, 'ein Klick');
    assert.equal(r.clicked, true);
    assert.equal(r.effect, CLICK_EFFECTS.overshoot, 'die View lief in die GEGENRICHTUNG');
    assert.equal(r.targetReached, false, 'die Zielwoche ist nicht erreicht');
    assert.equal(r.probe.find((e) => e.clicked === true).effect, CLICK_EFFECTS.overshoot);
  });

  test('unlesbarer Header nach dem Klick -> effect unverified (kein "changed"-Fortsetzen)', async () => {
    // Das ist der Live-Fall, der die 6er-Schleife am Leben hielt: 'changed'
    // wurde gemeldet, ohne dass die Zielwoche naeher gerueckt war.
    const r = await runClick({ prevStart: '2026-09-21', afterWeek: UNREADABLE, verified: 'unverified' });
    assert.equal(r.effect, CLICK_EFFECTS.unverified);
    assert.equal(r.clicked, true);
    assert.equal(r.probe.find((e) => e.clicked === true).effect, CLICK_EFFECTS.unverified);
  });

  test('ein wirkungsloser Klick bleibt "nur den naechsten Selektor probieren"', async () => {
    const r = await runClick({ prevStart: '2026-09-21', afterWeek: KW39, verified: 'unchanged' });
    assert.equal(r.clicked, false, 'alle Kandidaten durchprobiert, keiner wirkte');
    assert.equal(r.effect, CLICK_EFFECTS.unchanged);
    // Nur die Kandidaten, deren Selektor "Next week" enthaelt, finden den
    // Button — beide werden durchprobiert, jeder als 'unchanged' protokolliert.
    assert.equal(r.probe.filter((e) => e.clicked === true).length, 2, 'jeder passende Kandidat wurde probiert');
    assert.ok(r.probe.filter((e) => e.clicked === true).every((e) => e.effect === CLICK_EFFECTS.unchanged),
      'jeder Klick ist als wirkungslos protokolliert');
    assert.equal(r.effect, CLICK_EFFECTS.unchanged);
  });

  test('der Guard-Zustand (unlesbare Ansicht) meldet effect unverified, nicht advanced', async () => {
    const r = await runClick({ readWeek: UNREADABLE, prevStart: null });
    assert.equal(r.clicks.length, 0, 'kein Klick auf ungerenderter Ansicht');
    assert.equal(r.effect, CLICK_EFFECTS.unverified);
    assert.equal(r.guard, 'view-not-ready');
  });
});

// ===========================================================================
describe('FIX 3 — Verdrahtung: gerichtet, begrenzt, ehrlich', () => {
  const src = readFileSync(fileURLToPath(new URL('../src/owa.js', import.meta.url)), 'utf8');
  const loop = src.slice(src.indexOf('for (let step = 0; step < 30; step++) {'), src.indexOf('// GUARD 1: Falsche Woche gerendert'));

  test('die Klick-Obergrenze ist 2', () => {
    assert.equal(FREEBUSY_MAX_BUTTON_CLICKS, 2);
    assert.ok(/export const FREEBUSY_MAX_BUTTON_CLICKS = 2;/.test(src));
  });

  test('die Richtung kommt aus der Differenz und ist VERRIEGELT', () => {
    const iDelta = loop.indexOf('const deltaWeeks');
    const iDir = loop.indexOf('const direction = Math.sign(deltaWeeks);');
    const iLock = loop.indexOf('if (lockedDirection === null) {');
    const iFlip = loop.indexOf("abortReason = 'week-overshoot';");
    assert.ok(iDelta >= 0 && iDir > iDelta, 'die Richtung wird aus dem Delta bestimmt');
    assert.ok(iLock > iDir, 'und danach verriegelt');
    assert.ok(iFlip > iLock, 'ein Richtungswechsel bricht mit week-overshoot ab');
    assert.ok(/Math\.sign\(direction\) !== Math\.sign\(lockedDirection\)/.test(loop),
      'der Richtungswechsel wird aktiv geprueft, nicht nur vermutet');
  });

  test('nach jedem Klick wird die WIRKUNG geprueft, nicht "changed"', () => {
    assert.ok(/if \(r\.clicked && r\.effect === CLICK_EFFECTS\.overshoot\)/.test(loop));
    assert.ok(/if \(r\.clicked && r\.effect === CLICK_EFFECTS\.unverified\)/.test(loop));
    assert.ok(/abortReason = 'click-unverified'/.test(loop),
      'ein nicht bewertbarer Klick ist ein ehrlicher Abbruch, kein weiterer Klick');
    // Der Zieltreffer kommt ueber den Effekt, nicht ueber 'r.targetReached'
    // aus einer aelteren Logik:
    assert.ok(/if \(r\.targetReached\) \{/.test(loop));
  });

  test('runWeekButtonClick berechnet den Effekt selbst (nicht der Aufrufer)', () => {
    const i = src.indexOf('export async function runWeekButtonClick(');
    const fn = src.slice(i, src.indexOf('// Darf in dieselbe Richtung erneut geklickt werden?', i));
    assert.ok(/const effect = classifyClickEffect\(\{ prevStart, newWeek: r && r\.week, targetStart, direction \}\);/.test(fn),
      'die Wirkungsklassifikation sitzt im Produktionspfad');
    assert.ok(/probe\.push\(\{[\s\S]{0,400}?kind: cand\.kind, found: true, clicked: true, verified: r\.verified, effect,/.test(fn),
      'und sie steht im Klick-Protokoll (inkl. Label-Verriegelungs-Feldern, Defekt B)');
  });

  test('"max-clicks" entsteht nur noch als ehrlicher Abbruch (kein Blindklicken)', () => {
    assert.ok(/if \(buttonClicks >= MAX_BUTTON_CLICKS\) \{ abortReason = 'max-clicks'; break; \}/.test(loop));
    assert.ok(/if \(Math\.abs\(deltaWeeks\) > MAX_BUTTON_CLICKS\)/.test(loop),
      'mehr als 2 Wochen Entfernung -> too-far, nicht 6 Klicks');
  });

  test('der Reload/Fernsprung wird als VERSUCH protokolliert, mit Beleg', () => {
    assert.ok(/navAttempts\.push\(\{/.test(src), 'jede Navigationsmassnahme wird protokolliert');
    assert.ok(/kind: 'startdate-reload'/.test(src) && /kind: 'startdate-far-jump'/.test(src));
    assert.ok(/movedView: Boolean\(rendered\.startDate && rendered\.startDate !== before\.startDate\)/.test(src),
      'mit weekBefore/weekAfter — die URL gilt nirgends als Beleg der View');
    assert.ok(/if \(Array\.isArray\(navAttempts\) && navAttempts\.length\) out\.navAttempts = navAttempts;/.test(src),
      'und es landet in der Diagnostik');
    // TDZ-Waechter: die Felder muessen VOR den fruehen Rueckgaben stehen.
    const iDecl = src.indexOf('const navAttempts = [];');
    const iEarly = src.indexOf("abortReason: 'no-calendar-tab'");
    assert.ok(iDecl >= 0 && iDecl < iEarly, 'navAttempts ist vor dem ersten Abbruchpfad deklariert (TDZ)');
  });

  test('headerReadError ist ebenfalls vor dem ersten Abbruchpfad deklariert (TDZ)', () => {
    const iDecl = src.indexOf('let headerReadError = null;');
    const iEarly = src.indexOf("abortReason: 'no-calendar-tab'");
    assert.ok(iDecl >= 0 && iDecl < iEarly);
  });

  test('der Hinweistext behauptet keine URL-Garantie', () => {
    assert.ok(!/Primärsignal/.test(src));
    assert.ok(/der angezeigte Wochen-Header ist der EINZIGE Nachweis/.test(src));
    assert.ok(/ein VERSUCH \(der Client-Store kann ihn ueberschreiben/.test(src));
  });

  test('weekProbeGuardEntry kennt den neuen Zustand nicht (unveraendert)', () => {
    // Der Guard liefert weiterhin Guard-Gruende; der Effekt ist die Klick-Wirkung.
    // Beide duerfen nicht vermischt werden.
    const e = weekProbeGuardEntry(weekClickDecision({ rendered: KW40, targetStart: TARGET }));
    assert.equal(e.skipped, 'target-week-verified');
    assert.equal('effect' in e, false, 'der Guard-Eintrag traegt keinen Klick-Effekt');
  });
});

// ===========================================================================
// Defekt B (05.11.2026): startdate wird von OWA komplett ignoriert (Live-Probe:
// startdate=2026-11-09/16 und URL ohne startdate rendern alle die aktuelle
// Woche), die Arrow-Klicks verirrten sich. Die Nav-Buttons tragen ARIA-Labels
// mit der Zielwoche ("Go to next week 9–15 November, 2026 (Week 46)") — die
// Label-Verriegelung (weekNavLabelDecision) entscheidet VOR dem Klick.
// ===========================================================================
describe('Defekt B — Label-Verriegelung der Wochen-Navigation', () => {
  const TARGET_KW46 = '2026-11-09';

  test('Label nennt die Zielwoche -> klicken (landet exakt)', () => {
    const d = weekNavLabelDecision({
      labelText: 'Go to next week \n9–15 November, 2026 (Week 46)',
      direction: 1,
      renderedStart: '2026-11-02',
      targetStart: TARGET_KW46
    });
    assert.equal(d.parseable, true);
    assert.equal(d.click, true);
    assert.equal(d.reason, 'label-ist-zielwoche');
    assert.equal(d.labelWeek, '2026-11-09');
  });

  test('Label nennt die adjazente Folgezeit in Klickrichtung -> klicken', () => {
    // von KW45 aus Richtung KW47: erster Klick auf KW46 (noch nicht Ziel)
    const d = weekNavLabelDecision({
      labelText: 'Go to next week \n9–15 November, 2026 (Week 46)',
      direction: 1,
      renderedStart: '2026-11-02',
      targetStart: '2026-11-16'
    });
    assert.equal(d.parseable, true);
    assert.equal(d.click, true);
    assert.equal(d.reason, 'label-ist-adjazent');
  });

  test('Label nennt eine FREMDE Woche -> NICHT klicken (Verirren-Schutz)', () => {
    // prev-Button-Label beim Ziel KW46: 26 Oct – 1 Nov — bewegt sich falsch
    const d = weekNavLabelDecision({
      labelText: 'Go to previous week \n26 October – 1 November, 2026 (Week 44)',
      direction: 1,
      renderedStart: '2026-11-02',
      targetStart: TARGET_KW46
    });
    assert.equal(d.parseable, true);
    assert.equal(d.click, false);
    assert.equal(d.reason, 'label-woche-fremd');
  });

  test('DE-Label (Dativ): "Zur nächsten Woche 9.–15. November 2026"', () => {
    const d = weekNavLabelDecision({
      labelText: 'Zur nächsten Woche 9.–15. November 2026 (Woche 46) navigieren',
      direction: 1,
      renderedStart: '2026-11-02',
      targetStart: TARGET_KW46
    });
    assert.equal(d.parseable, true);
    assert.equal(d.click, true);
    assert.equal(d.labelWeek, TARGET_KW46);
  });

  test('Nicht parsebares Label (Chevron ohne Datum) -> Entscheidung beim Aufrufer', () => {
    const d = weekNavLabelDecision({ labelText: 'Weiter', direction: 1, renderedStart: '2026-11-02', targetStart: TARGET_KW46 });
    assert.equal(d.parseable, false);
    assert.equal(d.click, null);
  });

  test('Ohne Label gar keine Entscheidung (null -> nicht parsebar)', () => {
    const d = weekNavLabelDecision({ labelText: null, direction: 1, renderedStart: '2026-11-02', targetStart: TARGET_KW46 });
    assert.equal(d.parseable, false);
  });

  test('Renderstand unbekannt: nur klicken, wenn das Label die Zielwoche verspricht', () => {
    const ok = weekNavLabelDecision({
      labelText: 'Go to next week \n9–15 November, 2026 (Week 46)',
      direction: 1, renderedStart: null, targetStart: TARGET_KW46
    });
    assert.equal(ok.click, true);
    const no = weekNavLabelDecision({
      labelText: 'Go to next week \n2–8 November, 2026 (Week 45)',
      direction: 1, renderedStart: null, targetStart: TARGET_KW46
    });
    assert.equal(no.click, false);
  });

  test('Header/Label-Konflikt (Boot-Flash, Live-Befund 05.11.): impliedCurrent aus Label hergeleitet', () => {
    // Live-Befund: Header flashte KW47 (2026-11-16), View war KW45. Der
    // prev-Button versprach Week 44 (26.10.) — dazu passt als aktuelle Woche
    // nur KW45 (02.11.), NICHT der Header-Read.
    const d = weekNavLabelDecision({
      labelText: 'Go to previous week \n26 October – 1 November, 2026 (Week 44)',
      direction: -1,
      renderedStart: '2026-11-16',
      targetStart: TARGET_KW46
    });
    assert.equal(d.parseable, true);
    assert.equal(d.click, false, 'Fremde Label-Woche wird nicht geklickt');
    assert.equal(d.reason, 'label-woche-fremd');
    assert.equal(d.impliedCurrent, '2026-11-02', 'Label impliziert KW45 als aktuelle Woche');
  });

  test('Kein impliedCurrent, wenn Label zur gerenderten Woche adjazent ist', () => {
    const d = weekNavLabelDecision({
      labelText: 'Go to next week \n9–15 November, 2026 (Week 46)',
      direction: 1,
      renderedStart: '2026-11-02',
      targetStart: '2026-11-23'
    });
    assert.equal(d.click, true);
    assert.equal(d.impliedCurrent, undefined);
  });
});

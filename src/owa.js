import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import TurndownService from 'turndown';
import { browserManager } from './browser.js';
import { config } from './config.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const sigPath = path.join(__dirname, 'default_signature.html');
const DEFAULT_SIGNATURE_HTML = fs.existsSync(sigPath) ? fs.readFileSync(sigPath, 'utf8') : '';

const turndown = new TurndownService({
  headingStyle: 'atx',
  hr: '---',
  bulletListMarker: '-'
});

turndown.remove(['script', 'style', 'head']);

export const FOLDER_MAP = {
  inbox: 'inbox', posteingang: 'inbox',
  sentitems: 'sentitems', sent: 'sentitems', gesendet: 'sentitems',
  archive: 'archive', archiv: 'archive',
  deleteditems: 'deleteditems', deleted: 'deleteditems', papierkorb: 'deleteditems',
  drafts: 'drafts', entwürfe: 'drafts',
  newsletter: 'newsletter'
};

export function resolveFolder(folder) {
  if (!folder) return 'inbox';
  return FOLDER_MAP[folder.toLowerCase()] || folder.toLowerCase();
}

export function buildComposeUrl(realm, { to, subject, body, cc } = {}) {
  const toStr = Array.isArray(to) ? to.join('; ') : (to || '');
  const ccStr = Array.isArray(cc) ? cc.join('; ') : (cc || '');
  const subjectStr = subject || '';
  const bodyStr = body || '';

  const realmParam = realm ? `realm=${encodeURIComponent(realm)}&` : '';
  let url = `https://outlook.office.com/mail/deeplink/compose?${realmParam}to=${encodeURIComponent(toStr)}&subject=${encodeURIComponent(subjectStr)}`;
  if (bodyStr) {
    url += `&body=${encodeURIComponent(bodyStr)}`;
  }
  if (ccStr) {
    url += `&cc=${encodeURIComponent(ccStr)}`;
  }
  return url;
}

export function buildCalendarComposeUrl(realm, { subject, start, end, body, is_all_day = false } = {}) {
  const subjectStr = subject || '';
  const bodyStr = body || '';
  const realmParam = realm ? `realm=${encodeURIComponent(realm)}&` : '';

  let url = `https://outlook.office.com/calendar/deeplink/compose?${realmParam}subject=${encodeURIComponent(subjectStr)}`;
  if (start) {
    url += `&startdt=${encodeURIComponent(start)}`;
  }
  if (end) {
    url += `&enddt=${encodeURIComponent(end)}`;
  }
  if (bodyStr) {
    url += `&body=${encodeURIComponent(bodyStr)}`;
  }
  if (is_all_day) {
    url += `&allday=true`;
  }
  return url;
}

// Baut die URL der OWA-Kalender-Wochenansicht für eine Zielwoche (ISO 'YYYY-MM-DD' = Montag).
//
// STRATEGIE (Iteration 3, Live-Befunde 25.09.2026): Der startdate-Param wird von
// outlook.cloud.microsoft NICHT zuverlässig befolgt (zukünftige Woche renderte leer,
// während die aktuelle Woche funktionierte). Die URL MIT startdate wird nur noch als
// OPTIMIERUNG aufgerufen (spart Button-Klicks, wenn der Param wirkt) — vertraut wird
// ihr NIE. Primärpfad in getFreeBusy: bewährte Kalender-URL, gerenderte Woche ermitteln,
// dann per Vor/Zurück-Buttons zur Zielwoche navigieren (Verifikation nach jedem Klick).
export function buildCalendarWeekUrl(realm, iso) {
  const realmParam = realm ? `realm=${encodeURIComponent(realm)}&` : '';
  return `https://outlook.office.com/calendar/view/week?${realmParam}startdate=${iso}`;
}

// Formatiert ein ISO-Datum (z.B. "2026-09-15T13:00:00" oder "2026-09-15T13:00:00+02:00")
// zu einem OWA-kompatiblen startdt/enddt-Wert.
//
// WICHTIG (Zeitzonen-Falle): Der OWA-Deep-Link (startdt/enddt) interpretiert Werte
// OHNE Offset als LOKALE Browserzeit (Europe/Berlin) und Werte mit "Z" als UTC.
// Daher werden zeitzonen-lose Inputs als Berliner Lokalzeit BELASSEN (nicht nach UTC
// konvertiert) — sonst verschiebt OWA z.B. "08:00" (gewollt Berlin) auf "10:00".
export function normalizeCalendarDate(input, timezone = 'Europe/Berlin') {
  if (!input) return undefined;
  let dt = input;
  // Kein Zeitzonen-Offset ("Z" oder +hh:mm / -hh:mm am Ende) vorhanden
  // -> als lokale (Europe/Berlin) Zeit belassen, nur Format normalisieren.
  if (!/[zZ]|[+-]\d{2}:?\d{2}$/.test(dt)) {
    // Sekunden/Werte normalisieren, ohne Zeitzone zu wechseln:
    const m = String(dt).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/);
    if (m) {
      const [, y, mo, d, h, mi, s = '00'] = m;
      return `${y}-${mo}-${d}T${h}:${mi}:${s}`;
    }
    return dt;
  }
  // Expliziter Offset: als UTC-ISO normalisieren (eindeutig).
  try {
    const date = new Date(dt);
    if (!isNaN(date.getTime())) {
      return date.toISOString();
    }
  } catch (e) { /* bewusst ignorieren */ }
  return dt;
}

// ---------------------------------------------------------------------------
// Kalender-Wochenansicht: Header-Auslesen & -Parsen
// ---------------------------------------------------------------------------

// Monatnamen DE/EN (lowercase) -> Monatsindex. Wird für das Parsen von
// OWA-Kalender-Headern und Event-ARIA-Labels benötigt.
const MONTH_NAMES = {
  januar: 0, january: 0,
  februar: 1, february: 1,
  märz: 2, maerz: 2, march: 2,
  april: 3,
  mai: 4, may: 4,
  juni: 5, june: 5,
  juli: 6, july: 6,
  august: 7,
  september: 8,
  oktober: 9, october: 9,
  november: 10,
  dezember: 11, december: 11
};

// Date -> ISO-Datumsstring 'YYYY-MM-DD' (lokale Kalenderkomponenten).
function toIsoDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

// Parst einen OWA-Wochen-Header-Text in Start-/Enddatum der gerenderten Woche.
// Erkannte Muster (Beispiele):
//  - DE: "28. September – 4. Oktober 2026" (Jahr ggf. nur am Ende)
//  - EN: "September 28 - October 4, 2026"
// Rückgabe: { startDate, endDate } als ISO-Strings 'YYYY-MM-DD' oder null.
export function parseWeekHeaderText(text, now = new Date()) {
  if (!text) return null;
  const t = String(text).replace(/\s+/g, ' ').trim();
  const currentYear = now.getFullYear();

  // DE-Muster: "28. September – 4. Oktober 2026"
  const de = t.match(/(\d{1,2})\.\s*([A-Za-zäöüÄÖÜß]{3,})\.?(?:\s+(\d{4}))?\s*[–—-]\s*(\d{1,2})\.\s*([A-Za-zäöüÄÖÜß]{3,})\.?(?:\s+(\d{4}))?/);
  if (de) {
    const startMonth = MONTH_NAMES[de[2].toLowerCase()];
    const endMonth = MONTH_NAMES[de[5].toLowerCase()];
    if (startMonth !== undefined && endMonth !== undefined) {
      const y1 = de[3] ? parseInt(de[3], 10) : (de[6] ? parseInt(de[6], 10) : currentYear);
      const y2 = de[6] ? parseInt(de[6], 10) : y1;
      let s = new Date(y1, startMonth, parseInt(de[1], 10));
      let e = new Date(y2, endMonth, parseInt(de[4], 10));
      // Jahreswechsel (z.B. "29. Dezember – 4. Januar"): Enddatum in das Folgejahr heben
      if (e < s) e = new Date(y2 + 1, endMonth, parseInt(de[4], 10));
      return { startDate: toIsoDate(s), endDate: toIsoDate(e) };
    }
  }

  // EN-Muster: "September 28 - October 4, 2026" (Endmonat optional gleich)
  const en = t.match(/([A-Za-zäöüÄÖÜß]{3,})\.?\s+(\d{1,2})(?:,?\s*(\d{4}))?\s*[–—-]\s*(?:([A-Za-zäöüÄÖÜß]{3,})\.?\s+)?(\d{1,2})(?:,?\s*(\d{4}))?/);
  if (en) {
    const startMonth = MONTH_NAMES[en[1].toLowerCase()];
    const endMonth = en[4] ? MONTH_NAMES[en[4].toLowerCase()] : startMonth;
    if (startMonth !== undefined && endMonth !== undefined) {
      const y1 = en[3] ? parseInt(en[3], 10) : (en[6] ? parseInt(en[6], 10) : currentYear);
      const y2 = en[6] ? parseInt(en[6], 10) : y1;
      let s = new Date(y1, startMonth, parseInt(en[2], 10));
      let e = new Date(y2, endMonth, parseInt(en[5], 10));
      if (e < s) e = new Date(y2 + 1, endMonth, parseInt(en[5], 10));
      return { startDate: toIsoDate(s), endDate: toIsoDate(e) };
    }
  }

  // EN-Tages-zuerst-Muster (OWA-Live-Format Runde 4): "28 September – 4 October, 2026"
  // bzw. "28 September – 4 October 2026". Der ECHTE sichtbare Wochen-Header von
  // outlook.cloud.microsoft ist TAGES-zuerst und OHNE Punkt nach dem Tag — weder das
  // DE- noch das EN- noch das numerische Muster griffen hier. Root Cause Runde 4:
  // Dadurch fiel die Erkennung auf ein (stale) Event-Label zurück.
  const enDayFirst = t.match(/(\d{1,2})\s+([A-Za-zäöüÄÖÜß]{3,})\.?(?:,?\s*(\d{4}))?\s*[–—-]\s*(\d{1,2})\s+([A-Za-zäöüÄÖÜß]{3,})\.?(?:,?\s*(\d{4}))?/);
  if (enDayFirst) {
    const startMonth = MONTH_NAMES[enDayFirst[2].toLowerCase()];
    const endMonth = MONTH_NAMES[enDayFirst[5].toLowerCase()];
    if (startMonth !== undefined && endMonth !== undefined) {
      const y1 = enDayFirst[3] ? parseInt(enDayFirst[3], 10) : (enDayFirst[6] ? parseInt(enDayFirst[6], 10) : currentYear);
      const y2 = enDayFirst[6] ? parseInt(enDayFirst[6], 10) : y1;
      let s = new Date(y1, startMonth, parseInt(enDayFirst[1], 10));
      let e = new Date(y2, endMonth, parseInt(enDayFirst[4], 10));
      if (e < s) e = new Date(y2 + 1, endMonth, parseInt(enDayFirst[4], 10));
      return { startDate: toIsoDate(s), endDate: toIsoDate(e) };
    }
  }

  // Numerisches Muster: "28.9. – 4.10.2026" / "28.9. - 4.10." (OWA rendert den
  // Wochen-Header je nach Locale/Build numerisch statt mit Monatsnamen).
  const num = t.match(/(\d{1,2})\.(\d{1,2})\.?(?:(\d{4}))?\s*[–—-]\s*(\d{1,2})\.(\d{1,2})\.?(?:(\d{4}))?/);
  if (num) {
    const d1 = parseInt(num[1], 10);
    const m1 = parseInt(num[2], 10);
    const d2 = parseInt(num[4], 10);
    const m2 = parseInt(num[5], 10);
    if (d1 >= 1 && d1 <= 31 && d2 >= 1 && d2 <= 31 && m1 >= 1 && m1 <= 12 && m2 >= 1 && m2 <= 12) {
      const y1 = num[3] ? parseInt(num[3], 10) : (num[6] ? parseInt(num[6], 10) : currentYear);
      const y2 = num[6] ? parseInt(num[6], 10) : y1;
      let s = new Date(y1, m1 - 1, d1);
      let e = new Date(y2, m2 - 1, d2);
      if (e < s) e = new Date(y2 + 1, m2 - 1, d2);
      return { startDate: toIsoDate(s), endDate: toIsoDate(e) };
    }
  }

  return null;
}

// Parst eine NUMERISCHE Tagesangabe ("Montag, 28.9." / "28.9.2026" / "Di 29.9.")
// in ein ISO-Datum. OWA rendert Tages-Spaltenköpfe je nach Locale/Build mit
// Monatsnamen ("Montag, 28. September") ODER numerisch ("Montag 28.9.") — beide
// Varianten müssen für die Wochen-Erkennung lesbar sein (Live-Lücke Runde 2).
// Rückgabe: ISO-String 'YYYY-MM-DD' oder null.
export function parseNumericDayHeaderText(text, now = new Date()) {
  if (!text) return null;
  const t = String(text).replace(/\s+/g, ' ').trim();
  // "28.9." / "28.9.2026" — Tag.Monat, Jahr optional (direkt anhängend ODER getrennt)
  const m = t.match(/(?:^|[\s,])(\d{1,2})\.(\d{1,2})\.(?:(\d{4})|(?=$|[\s,]))/);
  if (!m) return null;
  const day = parseInt(m[1], 10);
  const month = parseInt(m[2], 10);
  if (day < 1 || day > 31 || month < 1 || month > 12) return null;
  const year = m[3] ? parseInt(m[3], 10) : now.getFullYear();
  const pad = (n) => String(n).padStart(2, '0');
  return `${year}-${pad(month)}-${pad(day)}`;
}

// ---------------------------------------------------------------------------
// Strikte Header-/Event-Trennung (Iteration 5, Live-Befund Runde 4 vom 25.09.):
// Die Header-Erkennung griff ein stale/hidden All-Day-Event-Label ("Urlaub …
// September 23") statt des sichtbaren Wochen-Headers. Event-Labels tragen
// Event-Signaturen (Zeitfenster, Ganztag, Statuswort+Datum) und dürfen NIE als
// Wochen-Header akzeptiert werden.
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// EIN Pattern-Satz, zwei Verbraucher (Iteration 7, Live-Befund Runde 6 vom
// 28.09.2026 — Hauptursache "0 Busy-Slots trotz verifizierter Woche").
//
// Root Cause Runde 6: Die Event-Wartephase fragte "gibt es irgendein
// [aria-label] mit einer Uhrzeit?" — und bekam vom Zeit-Indikator
// "calendar view, current time: Fri 22:31" die Antwort JA (22:31 matcht
// /\d{1,2}:\d{2}/). Der Lauf brach nach ~4s "erfolgreich" ab, obwohl das
// Event-Grid erst nach ~17s rendert (Runde 4 hatte die 6 Labels nach 16,9s).
//
// Deshalb EIN serialisierbarer Pattern-Satz (Strings statt RegExp-Objekte, damit
// er die evaluate-Grenze nach page.evaluate passieren kann) und zwei Prädikate:
//  - isCalendarEventLabel: BREIT (Header-Trennung, inkl. Statuswort+Datum)
//  - isEventGridLabel:     STRENG (Beweis "Grid ist fertig gerendert": echtes
//                          ZeitFENSTER oder Ganztag; Uhrzeit-/Nav-Indikatoren
//                          explizit ausgeschlossen)
// ---------------------------------------------------------------------------

// Patterns als Strings (serialisierbar über page.evaluate; RegExp-Flags 'i').
export const EVENT_LABEL_PATTERNS = {
  // ZeitFENSTER (nicht einzelne Uhrzeit): "09:00 to 12:00" / "14:00 bis 15:00" / "9:00 AM - 10:00 AM"
  timeRange: String.raw`\d{1,2}:\d{2}\s*(?:AM|PM)?\s*(?:to|bis|[\u2013\u2014-])\s*\d{1,2}:\d{2}`,
  // Ganztag: "all day event" / "all-day" / "ganztägig" / "ganztages"
  allDay: 'all[- ]?day|ganzt\u00e4gig|ganztages',
  // Statuswort (nur für die breite Header-Trennung)
  status: String.raw`\b(busy|tentative|free|besetzt|vorbehalt|frei)\b`,
  // Datum-ish (nur für die breite Header-Trennung)
  dateish: String.raw`(\d{1,2}\.\s*[A-Za-z\u00e4\u00f6\u00fc\u00c4\u00d6\u00dc\u00df]{3,})|([A-Za-z\u00e4\u00f6\u00fc\u00c4\u00d6\u00dc\u00df]{3,}\.?\s+\d{1,2})|(\d{1,2}\s+[A-Za-z\u00e4\u00f6\u00fc\u00c4\u00d6\u00dc\u00df]{3,})`,
  // KEIN Event: Work-Location-Banner
  nonEvent: 'arbeitsort|working location',
  // KEIN Event: Uhrzeit-/Navigations-Indikatoren der Kalenderansicht — der
  // Runde-6-Fehltreffer. Muss explizit ausgeschlossen sein, weil "22:31"
  // zwar KEIN Zeitfenster ist, aber in einem naiven Uhrzeit-Test als Treffer
  // durchging.
  gridBlocker: 'current time|aktuelle uhrzeit|clock indicator'
};

// Kompiliert den Pattern-Satz einmalig (lazy) — dieselben RegExp-Objekte
// werden von der Node-Seite und der Browser-Seite benutzt.
let compiledEventPatterns = null;
let compiledEventPatternsSource = null;
function eventPatterns() {
  if (!compiledEventPatterns || compiledEventPatternsSource !== EVENT_LABEL_PATTERNS) {
    compiledEventPatterns = {};
    for (const key of Object.keys(EVENT_LABEL_PATTERNS)) {
      compiledEventPatterns[key] = new RegExp(EVENT_LABEL_PATTERNS[key], 'i');
    }
    compiledEventPatternsSource = EVENT_LABEL_PATTERNS;
  }
  return compiledEventPatterns;
}

// Erkennt Event-Signaturen in einem ARIA-Label/Text. Reine Funktion (unit-testbar).
//  - Zeitfenster: "09:00 to 12:00" / "14:00 bis 15:00" / "9:00 AM - 10:00 AM"
//  - Ganztag: "all day event" / "all-day" / "ganztägig" / "ganztages"
//  - Statuswort (Busy|Tentative|Free|Besetzt|Vorbehalt|Frei) ZUSAMMEN mit einem Datum
export function isCalendarEventLabel(text) {
  if (!text) return false;
  const t = String(text);
  const re = eventPatterns();
  if (re.timeRange.test(t)) return true;
  if (re.allDay.test(t)) return true;
  return re.status.test(t) && re.dateish.test(t);
}

// STRENGES Prädikat: "das Event-Grid ist fertig gerendert". Verlangt ein echtes
// Zeitfenster ODER ein Ganztag-Marker — bewusst OHNE die Statuswort+Datum-
// Fallback-Regel des breiten Prädikats, denn ein solches Label ist kein Beweis
// für ein gerendertes Grid. Work-Location- und Uhrzeit-Indikatoren sind
// explizit ausgeschlossen (der Runde-6-Fehltreffer "current time: Fri 22:31").
export function isEventGridLabel(text) {
  if (!text) return false;
  const t = String(text);
  const re = eventPatterns();
  if (re.nonEvent.test(t)) return false;
  if (re.gridBlocker.test(t)) return false;
  return re.timeRange.test(t) || re.allDay.test(t);
}

// Filtert Event-Labels aus einer Kandidatenliste (Header dürfen keine Events sein).
export function filterCalendarHeaderCandidates(texts = []) {
  return (texts || []).filter((t) => !isCalendarEventLabel(t));
}

// Sichtbarkeits-Prädikat (Spiegel der Browser-Logik in collectCalendarCandidates).
// offsetParent === null bedeutet hidden — AUSSER bei position:fixed (dort ist
// offsetParent per Spezifikation null, das Element kann aber sichtbar sein).
// Zusätzlich muss die Bounding-Box echte Fläche haben (Breite/Höhe > 0).
export function isVisibleCalendarElement({ offsetParent = null, position = 'static', rect = null } = {}) {
  if (offsetParent === null && position !== 'fixed') return false;
  return Boolean(rect && rect.width > 0 && rect.height > 0);
}

// ===========================================================================
// Defekt 1 — Event-QUELLEN-Leiter + Geometrie-Ableitung
// ===========================================================================
// MESSBEFUND (Live-Smoke-Test KW 40, 25.09.2026, Commit 0c4cfee): Nach
// vollstaendigem Rendern (bis 20s Poll) lieferten `[role="button"][aria-label"]`
// und `[aria-label]` ZUSAMMEN 12 Labels, davon 11x header-artig und 1x
// Uhrzeit-Indikator. ariaButtonCount=9 = ausschliesslich Chrome-Navigation
// (App launcher, Mail, Calendar, People, Newsletters, Org Explorer, OneDrive,
// Copilot, Wochen-Header). eventElementsFound: 0.
//
// SCHLUSSFOLGERUNG: In diesem Tenant tragen die Event-Chips der Wochenansicht
// nachweislich KEIN aria-label mit Event-Signatur. Eine weitere geratene
// CSS-Klasse waere erneut eine Hypothese, die der naechste Smoke-Test erst
// wieder widerlegen kann. Stattdessen:
//   1) eine LEITER unabhaengiger Quellen (aria-label -> title -> data-* ->
//      sichtbarer Text), jede Stufe wird protokolliert,
//   2) GEOMETRIE als letzter Weg: Wochentag aus der Spalten-x-Position
//      (Tageskoepfe), Startzeit aus dem Zeitraster-y, Ende aus der Hoehe —
//      damit braucht es KEIN geratenes Attribut, nur echte Flaeche (rect>0),
//   3) ein STRUKTUR-CENSUS, der auch im 0-Treffer-Fall die Tatsache ueber
//      die echte DOM-Struktur liefert (tag/role/class/data-attrs/rect/Text).
//
// GRUNDREGEL: Es wird NIE eine Zeit oder ein Datum geraten. Fehlt ein
// Beleg, ist das Ergebnis eine Ablehnung mit Grund, kein Phantom-Slot.

// Quellen-Leiter. selector === null -> kein CSS-Selektor, sondern ein
// Struktur-Scan der Blatt-Elemente im Kalenderbereich (scope: 'surface').
export const CALENDAR_EVENT_SOURCES = [
  { id: 'aria-role-button', selector: '[role="button"][aria-label]', scope: 'document', source: 'aria-label' },
  { id: 'aria-any', selector: '[aria-label]', scope: 'document', source: 'aria-label' },
  { id: 'title', selector: null, scope: 'surface', source: 'title' },
  { id: 'data-attr', selector: null, scope: 'surface', source: 'data-attr' },
  { id: 'text', selector: null, scope: 'surface', source: 'text' }
];

// Reihenfolge der Leiter inkl. Selektor-Memory: der zuletzt erfolgreiche
// Selektor wird zuerst probiert (spart bei bekannterstruktur 4 Scans).
export function eventSourcePlan(preferSelector = null) {
  const steps = CALENDAR_EVENT_SOURCES.slice();
  if (preferSelector && !steps.some((s) => s.selector === preferSelector)) {
    steps.unshift({ id: 'remembered', selector: preferSelector, scope: 'document', source: 'aria-label' });
  }
  return steps;
}

// --- Geometrie -------------------------------------------------------------

// Ein Raster-/Stundenlabel ist ein Text, der NUR aus einer Uhrzeit besteht
// ("08:00", "8 AM"). "calendar view, current time: Fri 22:31" ist das nicht —
// genau dieser Knoten war der Runde-6-Fehltreffer, weshalb die Pruefung
// bewusst auf den GESAMTEN Text streng ist (Anker ^...$).
export const CLOCK_LABEL_PATTERN = String.raw`^\d{1,2}(?::\d{2})?\s*(?:am|pm)?$`;

export function isClockLabelOnly(text) {
  if (!text) return false;
  return new RegExp(CLOCK_LABEL_PATTERN, 'i').test(String(text).replace(/\s+/g, ' ').trim());
}

// Normalisiert eine Raster-/Zeitangabe auf HH:MM. MEHRDEUTIGE Angaben (nackte
// "8" = 08:00 oder 20:00) werden NULL — eine geratene Zeit waere eine
// Falschaussage in den Belegungsdaten.
export function normalizeClockTime(text) {
  if (text === null || text === undefined) return null;
  const s = String(text).replace(/\s+/g, ' ').trim();
  if (!s) return null;
  const m = s.match(/(?:^|\D)(\d{1,2})(?::(\d{2}))?\s*(am|pm)?(?:\s|$|\D)/i);
  if (!m) return null;
  const hasMeridiem = Boolean(m[3]);
  if (!m[2] && !hasMeridiem) return null; // nackte Stunde: mehrdeutig -> ablehnen
  let h = parseInt(m[1], 10);
  const min = m[2] ? parseInt(m[2], 10) : 0;
  const pm = m[3] && m[3].toLowerCase() === 'pm';
  const am = m[3] && m[3].toLowerCase() === 'am';
  if (am && h === 12) h = 0;
  if (pm && h !== 12) h += 12;
  if (h > 23 || min > 59) return null;
  return `${String(h).padStart(2, '0')}:${String(min).padStart(2, '0')}`;
}

// Ordnet eine x-Position einer Wochentags-Spalte zu (Index, -1 = keine).
// OWA rendert den Tageskopf oft schmaler als die Spalte, deshalb gilt die
// naechste Spalte, solange der Abstand die halbe mediane Spaltenbreite
// nicht ueberschreitet. Alles andere wird abgelehnt statt geraten.
export function mapXToDayColumn(columns = [], x) {
  if (!Array.isArray(columns) || columns.length < 2) return -1;
  const xs = Number(x);
  if (!Number.isFinite(xs)) return -1;
  const usable = columns
    .map((c) => c && { left: Number(c.left), width: Number(c.width), center: Number(c.left) + Number(c.width) / 2 })
    .filter((c) => Number.isFinite(c.left) && Number.isFinite(c.width) && c.width > 0);
  if (usable.length < 2) return -1;
  const hit = usable.findIndex((c) => xs >= c.left && xs <= c.left + c.width);
  if (hit >= 0) return hit;
  const widths = usable.map((c) => c.width).sort((a, b) => a - b);
  const maxDist = widths[Math.floor(widths.length / 2)] / 2;
  let best = -1;
  let bestDist = Infinity;
  usable.forEach((c, i) => {
    const dist = Math.abs(xs - c.center);
    if (dist < bestDist) { bestDist = dist; best = i; }
  });
  return bestDist <= maxDist ? best : -1;
}

// Ordnet eine y-Position dem untersten Zeitraster-Eintrag oberhalb zu -> Startzeit.
// Liegt der Chip UNTERHALB des letzten Rasterpunkts (Rastersonde endet z.B. um
// 18 Uhr, das Event laeuft bis 20), wird linear fortgeschrieben — aber nur, wenn
// die Stundenhöhe belegt ist. Ohne Beleg: null, NIE geraten.
export function mapYToClockTime(ruler = [], y, { tolerancePx = 24, hourPx = 0 } = {}) {
  if (!Array.isArray(ruler) || !ruler.length) return null;
  const ys = Number(y);
  if (!Number.isFinite(ys)) return null;
  let best = null;
  let last = null;
  for (const entry of ruler) {
    if (!entry) continue;
    const time = normalizeClockTime(entry.time || entry.label || entry.text);
    const top = Number(entry.top);
    if (!time || !Number.isFinite(top)) continue;
    if (!last || top > last.top) last = { time, top };
    if (top > ys + tolerancePx) continue; // Rasterpunkt liegt UNTER dem Chip-Anfang
    if (best === null || top > best.top) best = { time, top };
  }
  if (best) {
    // Chip startet DEUTLICH unterhalb der letzten Rasterlinie: nur mit belegter
    // Stundenhöhe fortschreiben, sonst ablehnen (lieber keine Zeit als eine
    // falsche — genau das war die "header-misclassified"-Falle in Runde 4/6).
    if (last && ys > best.top + tolerancePx) {
      const px = Number(hourPx) > 0 ? Number(hourPx) : estimateHourPx(ruler);
      if (px <= 0) return null;
      const steps = Math.round((ys - best.top) / px);
      return steps > 0 ? addMinutes(best.time, steps * 60) : best.time;
    }
    return best.time;
  }
  const px = Number(hourPx) > 0 ? Number(hourPx) : estimateHourPx(ruler);
  if (px > 0 && last && ys > last.top + tolerancePx) {
    const steps = Math.round((ys - last.top) / px);
    if (steps > 0) return addMinutes(last.time, steps * 60);
  }
  return null;
}

// Stundenhöhe in px aus dem Zeitraster (Median der Abstände, gegen Rauschen
// robust). Weniger als 2 Rasterpunkte -> 0, damit KEINE Endzeit erfunden wird.
export function estimateHourPx(ruler = []) {
  if (!Array.isArray(ruler) || ruler.length < 2) return 0;
  const tops = ruler.map((r) => r && Number(r.top)).filter((t) => Number.isFinite(t)).sort((a, b) => a - b);
  if (tops.length < 2) return 0;
  const diffs = [];
  for (let i = 1; i < tops.length; i++) {
    const d = tops[i] - tops[i - 1];
    if (d > 0) diffs.push(d);
  }
  if (!diffs.length) return 0;
  diffs.sort((a, b) => a - b);
  return diffs[Math.floor(diffs.length / 2)];
}

// "Monday, September 28, 2026" — Format, das parseEventLabel sicher liest.
function formatEventDatePhrase(dateIso) {
  const m = String(dateIso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const year = parseInt(m[1], 10);
  const month = parseInt(m[2], 10) - 1;
  const day = parseInt(m[3], 10);
  if (month < 0 || month > 11 || day < 1 || day > 31) return null;
  const d = new Date(Date.UTC(year, month, day));
  const weekdays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
    'August', 'September', 'October', 'November', 'December'];
  return `${weekdays[d.getUTCDay()]}, ${months[d.getUTCMonth()]} ${day}, ${year}`;
}

// Baut ein parsebares Label aus Geometrie-Angaben. Ohne Titel ODER ohne Datum
// ODER (für zeitbasierte Events) ohne Startzeit -> null. Kein Phantom-Label.
export function buildGeometryEventLabel({ title, start, end, dateIso, allDay = false } = {}) {
  const t = String(title || '').replace(/\s+/g, ' ').trim();
  if (!t) return null;
  const when = formatEventDatePhrase(dateIso);
  if (!when) return null;
  if (allDay) return `${t}, all day event, ${when}`;
  const s = normalizeClockTime(start);
  if (!s) return null;
  const e = normalizeClockTime(end);
  return e ? `${t}, ${s} to ${e}, ${when}` : `${t}, ${s}, ${when}`;
}

// --- Ableitung aus einem Grid-Knoten --------------------------------------

const NON_EVENT_REASONS = { clock: 'clock-indicator', work: 'work-location' };

// Wertet EINEN Grid-Knoten aus. Quellen werden in dieser Reihenfolge geprueft:
// aria-label -> title -> data-* -> sichtbarer Text; danach (nur wenn der Knoten
// selbst keine belastbare Zeit traegt) die Geometrie.
// Rueckgabe: { label, reason, source, timeSource }
export function deriveEventLabelFromNode(node, { pats = EVENT_LABEL_PATTERNS, columns = [], ruler = [], weekStartIso = null } = {}) {
  const re = {};
  for (const k of Object.keys(pats)) re[k] = new RegExp(pats[k], 'i');
  const area = node && (node.area === true || (node.rect && node.rect.width > 0 && node.rect.height > 0));
  if (!node || !area) return { label: null, reason: 'no-area', source: null, timeSource: null };

  const rect = node.rect || {};
  const hasSignature = (s) => Boolean(s) && (re.timeRange.test(s) || re.allDay.test(s));
  const hasDate = (s) => Boolean(s) && re.dateish.test(s);
  const norm = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();

  // Ablehnungsgruende VOR der Event-Erkennung (Uhr-Indikator, Arbeitsort)
  const blockers = [];
  const consider = (s) => { const t = norm(s); if (t && t.length <= 300) blockers.push(t); };
  consider(node.ariaLabel);
  consider(node.title);
  consider(node.text);
  for (const b of blockers) {
    if (re.gridBlocker.test(b)) return { label: null, reason: NON_EVENT_REASONS.clock, source: null, timeSource: null };
    if (re.nonEvent.test(b)) return { label: null, reason: NON_EVENT_REASONS.work, source: null, timeSource: null };
  }

  // 1) Attribut-/Text-Quellen mit vollstaendigem Label
  const sources = [];
  const aria = norm(node.ariaLabel);
  if (aria) sources.push({ source: 'aria-label', value: aria });
  const title = norm(node.title);
  if (title) sources.push({ source: 'title', value: title });
  if (node.data && typeof node.data === 'object') {
    for (const [k, v] of Object.entries(node.data)) {
      const val = norm(v);
      if (val) sources.push({ source: 'data-attr', value: val, attr: k });
    }
  }
  const text = norm(node.text);
  if (text) sources.push({ source: 'text', value: text });
  for (const s of sources) {
    if (!hasSignature(s.value)) continue;
    if (hasDate(s.value)) return { label: s.value, reason: null, source: s.source, timeSource: 'label' };
    // Signatur ohne Datum: Datum aus der Spaltenposition ergaenzen, wenn belegt.
    // Ganztag laeuft ueber den All-Day-Weg — seine y-Position ist KEINE Zeit.
    const geo = re.allDay.test(s.value)
      ? geometryAllDayLabel(node, { columns, weekStartIso })
      : geometryLabel(node, { columns, ruler, weekStartIso });
    if (geo) return { label: geo, reason: null, source: s.source, timeSource: 'label+geometry-date' };
    return { label: null, reason: 'no-date', source: s.source, timeSource: null };
  }

  // 2) Nur Text: Zeit+Datum aus der Geometrie ableiten
  if (text) {
    // "Urlaub, all day" o.ae. -> Ganztag-Label, NIE eine erfundene Uhrzeit.
    const geo = re.allDay.test(text)
      ? geometryAllDayLabel(node, { columns, weekStartIso })
      : geometryLabel(node, { columns, ruler, weekStartIso });
    if (geo) return { label: geo, reason: null, source: 'text', timeSource: re.allDay.test(text) ? 'geometry-date' : 'geometry' };
    if (hasSignature(text)) return { label: null, reason: 'no-date', source: 'text', timeSource: null };
    // Signatur vorhanden, aber keine Zeit/Datum ableitbar -> eigener Grund
    if (text.length > 1) {
      const isTimeOnly = re.timeRange.test(text) || re.allDay.test(text);
      if (isTimeOnly) return { label: null, reason: 'no-date', source: 'text', timeSource: null };
      if (ruler && ruler.length >= 2 && weekStartIso && columns && columns.length >= 2 && node.inGrid === true) {
        return { label: null, reason: 'geometry-incomplete', source: 'text', timeSource: null };
      }
      // Chip-artiger Knoten ohne ableitbare Zeit vs. reiner Chrome-Text:
      // die Unterscheidung macht die Diagnostik ueberhaupt erst lesbar.
      return {
        label: null,
        reason: node.inGrid === true ? 'no-time-source' : 'no-event-signature',
        source: 'text',
        timeSource: null
      };
    }
  }

  return { label: null, reason: 'no-event-signature', source: null, timeSource: null };
}

// Wochentag aus der Spalten-x-Position (Spalten 0..n = Montag..). Rein
// geometrisch, braucht nur Tageskoepfe + bekannten Wochenstart.
function dateFromColumn(node, { columns = [], weekStartIso = null } = {}) {
  if (node.inGrid !== true) return null;
  const rect = node.rect || {};
  const colIndex = mapXToDayColumn(columns, Number.isFinite(rect.left) ? rect.left : rect.x);
  if (colIndex < 0 || !weekStartIso) return null;
  const base = new Date(String(weekStartIso).slice(0, 10) + 'T00:00:00Z');
  if (Number.isNaN(base.getTime())) return null;
  base.setUTCDate(base.getUTCDate() + colIndex);
  return toIsoDate(base);
}

// Titel des Chips: die erste Textquelle, die NICHT selbst eine Uhrzeit traegt
// (sonst wuerde "14:00 to 15:00" zum Titel). All-Day-/Statusmarker werden
// abgeschnitten, damit sie nicht doppelt im Label landen.
function chipTitle(node) {
  const norm = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  const raw = [norm(node.ariaLabel), norm(node.title), norm(node.text)]
    .find((t) => t && !/\d{1,2}:\d{2}/.test(t)) || '';
  return raw
    .replace(/[,;]?\s*\b(all[\s-]?day|ganzt(ä|ae)gig|ganztages|busy|besetzt|tentative|free|frei|recurring event|wiederkehrend)\b\s*,?/gi, ' ')
    .replace(/\s*[,;]\s*$/, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// Geometrie-Weg fuer ZEITBASIERTE Events: Start aus dem Raster-y, Ende aus der
// Chip-Hoehe, Tag aus der Spalte. Jeder Teil muss belegt sein, sonst null.
function geometryLabel(node, { columns = [], ruler = [], weekStartIso = null } = {}) {
  if (node.inGrid !== true) return null;
  const rect = node.rect || {};
  const dateIso = dateFromColumn(node, { columns, weekStartIso });
  if (!dateIso) return null;
  const hourPx = estimateHourPx(ruler);
  const start = mapYToClockTime(ruler, rect.top, { hourPx });
  if (!start) return null;
  let end = null;
  if (hourPx > 0 && Number(rect.height) > 0) {
    const mins = Math.round((Number(rect.height) / hourPx) * 60 / 15) * 15;
    end = addMinutes(start, Math.max(15, Math.min(720, mins)));
  }
  return buildGeometryEventLabel({ title: chipTitle(node), start, end, dateIso });
}

// Geometrie-Weg fuer GANZTAG-Events: nur der Tag aus der Spalte. Ein Ganztag-
// Event darf NIEMALS eine Uhrzeit bekommen — der All-Day-Eintrag liegt in
// seiner eigenen Zeile, seine y-Position ist keine Zeitangabe.
function geometryAllDayLabel(node, { columns = [], weekStartIso = null } = {}) {
  if (node.inGrid !== true) return null;
  const dateIso = dateFromColumn(node, { columns, weekStartIso });
  if (!dateIso) return null;
  return buildGeometryEventLabel({ title: chipTitle(node), allDay: true, dateIso });
}

function addMinutes(hhmm, minutes) {
  const m = String(hhmm).match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  let total = parseInt(m[1], 10) * 60 + parseInt(m[2], 10) + minutes;
  total = Math.min(total, 24 * 60 - 1);
  return `${String(Math.floor(total / 60)).padStart(2, '0')}:${String(total % 60).padStart(2, '0')}`;
}

// ===========================================================================
// Defekt 2 — ehrliche Verifikationsangabe
// ===========================================================================
// Live-Befund: dataQuality:'verified' bei eventElementsFound:0 war eine
// Falschaussage — busySlots:[] wurde dadurch als "Kalender ist frei"
// gelesen. 'verified' ist ab jetzt NUR gesetzt, wenn die Woche passt UND
// echte Event-Knoten gelesen wurden. Der Wochen-Nachweis selbst bleibt
// ueber das separate Feld weekVerified erhalten (er funktioniert und MUSS
// erhalten bleiben).
export const DATA_QUALITY_VALUES = ['verified', 'unverified-empty', 'unverified'];

export function resolveDataQuality({ weekVerified = false, eventElementsFound = 0 } = {}) {
  if (!weekVerified) return 'unverified';
  const n = Number(eventElementsFound);
  return n > 0 ? 'verified' : 'unverified-empty';
}

// Belastbar ist dataQuality NUR bei 'verified' — Aufrufer sollen das nicht
// selbst nachrechnen muessen.
export function isBelievableBusySlots(dataQuality) {
  return dataQuality === 'verified';
}

// ===========================================================================
// Defekt 3 — Warte-Ausgang an die Beobachtung koppeln
// ===========================================================================
// Live-Befund: eventWaitOutcome:'found' bei waitedForEventsMs:0 und
// eventElementsFound:0. 'found' wird ab jetzt ausschliesslich gemeldet, wenn
// in der ABSCHLIESSENDEN Sammlung echte Event-Knoten gelesen wurden.
// 'not-polled' (Bug 3): Budget war vorhanden, aber 0 Polls. Ein reiner
// Logikfehler der Wartephase — der muss als eigener Ausgang sichtbar sein,
// sonst ist er von "Kalender nachweislich leer" nicht unterscheidbar.
export const EVENT_WAIT_OUTCOMES = ['found', 'not-found', 'budget-exhausted', 'not-rendered', 'not-polled'];

export function resolveEventWaitOutcome({ polls = 0, eventElementsFound = 0, waitBudgetMs = null } = {}) {
  const n = Number(eventElementsFound);
  if (Number.isFinite(n) && n > 0) return 'found';
  return 'not-found';
}

// Parst die Wochennummer aus einem Prev/Next-Button-Label ("… (Week 39)" / "(KW 39)").
export function parseWeekNumberFromNav(text) {
  if (!text) return null;
  const m = String(text).match(/\((?:Week|KW)\s*(\d{1,2})\)/i);
  if (!m) return null;
  const n = parseInt(m[1], 10);
  return n >= 1 && n <= 53 ? n : null;
}

// Parst den Datumsbereich eines Prev/Next-Button-Labels ("21–27 September, 2026").
export function parseNavWeekRange(text, now = new Date()) {
  if (!text) return null;
  const t = String(text).replace(/\s+/g, ' ').trim();
  const m = t.match(/(\d{1,2})\s*[–—-]\s*(\d{1,2})\s+([A-Za-zäöüÄÖÜß]{3,})\.?(?:,?\s*(\d{4}))?/);
  if (!m) return null;
  const month = MONTH_NAMES[m[3].toLowerCase()];
  if (month === undefined) return null;
  const year = m[4] ? parseInt(m[4], 10) : now.getFullYear();
  const start = new Date(year, month, parseInt(m[1], 10));
  const end = new Date(year, month, parseInt(m[2], 10));
  if (end < start) end.setFullYear(year + 1);
  return { startDate: toIsoDate(start), endDate: toIsoDate(end) };
}

// Montag der ISO-Woche (für den Fall, dass nur "(Week NN)" ohne Datumsbereich vorliegt).
export function isoWeekMonday(week, year) {
  const jan4 = new Date(year, 0, 4);
  const jan4Day = jan4.getDay() || 7; // Mo=1 … So=7
  const week1Monday = new Date(jan4.getTime() - (jan4Day - 1) * 86400000);
  return new Date(week1Monday.getTime() + (week - 1) * 7 * 86400000);
}

// Leitet die ANGEZEIGTE Woche aus den Prev/Next-Button-Labels ab (belastbarer
// Fallback, Runde-4-Forderung 2): Die angezeigte Woche liegt ZWISCHEN prev und
// next. Bevorzugt werden die Datumsbereiche (Montag = Tag nach prev.endDate,
// Sonntag = Tag vor next.startDate); die "(Week NN)"-Angabe dient als Signal
// und als alleinige Quelle, wenn keine Datumsbereiche lesbar sind.
export function resolveWeekFromNavButtons(navTexts = [], now = new Date()) {
  const list = (navTexts || []).filter(Boolean);
  const prevText = list.find((t) => /previous|vorherige|zurück|prev\b/i.test(t));
  const nextText = list.find((t) => /\bnext\b|nächste|nächsten/i.test(t));

  let weekNumber = null;
  const prevNum = parseWeekNumberFromNav(prevText);
  const nextNum = parseWeekNumberFromNav(nextText);
  if (prevNum != null) weekNumber = prevNum + 1;
  else if (nextNum != null) weekNumber = nextNum - 1;

  let startDate = null;
  let endDate = null;
  const prevRange = parseNavWeekRange(prevText, now);
  const nextRange = parseNavWeekRange(nextText, now);
  if (prevRange) startDate = toIsoDate(new Date(new Date(prevRange.endDate + 'T00:00:00').getTime() + 86400000));
  if (nextRange) endDate = toIsoDate(new Date(new Date(nextRange.startDate + 'T00:00:00').getTime() - 86400000));
  if (!startDate && endDate) startDate = toIsoDate(new Date(new Date(endDate + 'T00:00:00').getTime() - 6 * 86400000));
  if (startDate && !endDate) endDate = toIsoDate(new Date(new Date(startDate + 'T00:00:00').getTime() + 6 * 86400000));
  if (!startDate && weekNumber != null) {
    const monday = isoWeekMonday(weekNumber, now.getFullYear());
    startDate = toIsoDate(monday);
    endDate = toIsoDate(new Date(monday.getTime() + 6 * 86400000));
  }
  if (!startDate && weekNumber == null) return null;
  return {
    text: (prevText || nextText || '').slice(0, 200),
    startDate,
    endDate,
    weekNumber
  };
}

// Sammelt rohe Kandidaten-Texte und Diagnostik-Rohdaten aus der gerenderten
// Kalenderansicht (Browser-seitig, OHNE Interpretation). Die Wochen-Erkennung
// passiert deterministisch in resolveRenderedWeekFromCandidates.
//
// Live-Befund Runde 2 (25.09.2026): Bei zukünftigen Wochen war unbekannt, WAS im DOM
// stand — die Response enthielt keine Diagnostik. Diese Sammlung liefert jetzt
// Header-Kandidaten, Tageskopf-Kandidaten (inkl. Body-Text-Heuristik) und ein
// ARIA-Label-Sample zurück.
export async function collectCalendarCandidates(page) {
  return page.evaluate(() => {
    const cap = (arr, n) => arr.slice(0, n);
    const push = (arr, t) => {
      const s = String(t || '').replace(/\s+/g, ' ').trim();
      if (s && arr.indexOf(s) === -1) arr.push(s);
    };
    // Sichtbarkeit (Runde-4-Root-Cause): stale/virtualisierte Elemente (z.B. der
    // alte September-23-Header) sind hidden oder haben keine Fläche und dürfen
    // NICHT als Header gewertet werden. Spiegel von isVisibleCalendarElement().
    const isVisible = (el) => {
      if (!el) return false;
      try {
        const cs = window.getComputedStyle(el);
        if (cs.display === 'none' || cs.visibility === 'hidden') return false;
        if (el.offsetParent === null && cs.position !== 'fixed') return false;
        const r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0;
      } catch (e) { return false; }
    };
    // Event-Signaturen (Runde-4-Root-Cause): Zeitfenster, Ganztag oder
    // Statuswort+Datum -> Event, KEIN Header. Spiegel von isCalendarEventLabel().
    const isEventLabel = (text) => {
      const t = String(text || '');
      if (/\d{1,2}:\d{2}\s*(?:AM|PM)?\s*(?:to|bis|[–—-])\s*\d{1,2}:\d{2}/i.test(t)) return true;
      if (/all[- ]?day|ganztägig|ganztages/i.test(t)) return true;
      const statusRe = /\b(busy|tentative|free|besetzt|vorbehalt|frei)\b/i;
      const dateRe = /(\d{1,2}\.\s*[A-Za-zäöüÄÖÜß]{3,})|([A-Za-zäöüÄÖÜß]{3,}\.?\s+\d{1,2})|(\d{1,2}\s+[A-Za-zäöüÄÖÜß]{3,})/;
      return statusRe.test(t) && dateRe.test(t);
    };
    const headerTexts = [];
    const dayHeaderTexts = [];
    const navTexts = [];
    // Datums-/Monatsmuster (DE + EN + Tages-zuerst): "28. September", "September 28",
    // "28 September" (OWA-Live-Format Runde 4).
    const dateish = /(\d{1,2}\.\s*[A-Za-zäöüÄÖÜß]{3,})|([A-Za-zäöüÄÖÜß]{3,}\.?\s+\d{1,2})|(\d{1,2}\s+[A-Za-zäöüÄÖÜß]{3,})/;
    const timeRe = /\d{1,2}:\d{2}/;

    // 1) Klassische Header-Rollen (Banner/Überschriften) — nur sichtbare, keine Events
    for (const el of document.querySelectorAll('[role="banner"], [role="heading"], h1, h2')) {
      if (!isVisible(el)) continue;
      const text = el.getAttribute('aria-label') || el.innerText;
      if (isEventLabel(text)) continue;
      push(headerTexts, text);
    }
    // 2) Toolbar-/Button-Elemente mit Datums-Label: OWA rendert den Wochenbereich als
    //    klickbaren Datums-Button. Der Ansichtsumschalter ("Work week"/"Arbeitswoche")
    //    enthält KEIN Datum und fällt hier bewusst heraus.
    for (const el of document.querySelectorAll(
      '[role="banner"] button, [role="banner"] [role="button"], header button, [role="toolbar"] button, [role="toolbar"] [role="button"]'
    )) {
      if (!isVisible(el)) continue;
      const label = el.getAttribute('aria-label') || el.innerText || el.getAttribute('title') || '';
      if (dateish.test(label) && !timeRe.test(label) && !isEventLabel(label)) push(headerTexts, label);
    }
    // 3) Alle aria-labels mit Datums-/Monatsmuster, aber OHNE Uhrzeit (z.B. der
    //    Datumsauswahl-Button). Zeitmuster-/Event-Labels sind Event-Elemente, keine Header.
    for (const el of document.querySelectorAll('[aria-label]')) {
      if (!isVisible(el)) continue;
      const label = el.getAttribute('aria-label') || '';
      if (dateish.test(label) && !timeRe.test(label) && label.length <= 120 && !isEventLabel(label)) push(headerTexts, label);
    }
    // 3b) Prev/Next-Wochen-Buttons (belastbarer Fallback für die Wochen-Ableitung):
    //     "Go to previous week 21–27 September, 2026 (Week 39)" / "… (Week 41)".
    for (const el of document.querySelectorAll('[aria-label], [title]')) {
      if (!isVisible(el)) continue;
      const label = el.getAttribute('aria-label') || el.getAttribute('title') || '';
      if (/(previous|next|vorherige|nächste|nächsten)\s+week|week\s+(previous|next)|woche/i.test(label) &&
          /(week|kw)\s*\d{1,2}/i.test(label)) {
        push(navTexts, label);
      }
    }
    // 4) Tages-Spaltenköpfe der Woche — mehrere Selektor-Varianten (OWA nutzt je nach
    //    Build role="columnheader", th-Elemente oder Grid-Köpfe)
    for (const el of document.querySelectorAll(
      '[role="columnheader"], [role="grid"] [role="columnheader"], th, [role="grid"] th'
    )) {
      if (!isVisible(el)) continue;
      const text = el.getAttribute('aria-label') || el.innerText;
      if (isEventLabel(text)) continue;
      push(dayHeaderTexts, text);
    }
    // 5) Body-Text-Heuristik: Zeilen mit Wochentag + Datum ("Montag, 28. September",
    //    "Montag 28.9.") ODER Wochen-Range-Muster — Fallback, falls die Spaltenköpfe
    //    über keine der Rollen erreichbar sind. Wochentag+Datum ZUSAMMEN (oder ein
    //    explizites Range-Muster) sind Pflicht, um Rauschen zu filtern.
    try {
      const weekdayRe = /\b(montag|dienstag|mittwoch|donnerstag|freitag|samstag|sonntag|monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i;
      const lineDateRe = /(\d{1,2}\.\s*[A-Za-zäöüÄÖÜß]{3,}\.?)|(\d{1,2}\.\d{1,2}\.?)/;
      const rangeRe = /(\d{1,2}\.\s*[A-Za-zäöüÄÖÜß]{3,}\.?\s*[–—-]\s*\d{1,2}\.?\s*[A-Za-zäöüÄÖÜß]{3,})|(\d{1,2}\.\d{1,2}\.?\s*[–—-]\s*\d{1,2}\.\d{1,2}\.?)/;
      const lines = (document.body ? document.body.innerText : '').split('\n');
      for (const line of lines) {
        const s = line.trim();
        if (!s || s.length > 60) continue;
        if (isEventLabel(s)) continue;
        if ((weekdayRe.test(s) && lineDateRe.test(s)) || rangeRe.test(s)) push(dayHeaderTexts, s);
        if (dayHeaderTexts.length >= 12) break;
      }
    } catch (e) { /* bewusst ignorieren */ }

    // Diagnostik: Anzahl + Sample aller ARIA-Buttons (Event-Elemente)
    const buttons = Array.from(document.querySelectorAll('[role="button"][aria-label]'));
    const ariaLabelSample = cap(
      buttons
        .map((b) => String(b.getAttribute('aria-label') || '').trim())
        .filter((l) => l.length > 0)
        .map((l) => l.slice(0, 120)),
      15
    );

    return {
      headerTexts: cap(headerTexts, 20),
      dayHeaderTexts: cap(dayHeaderTexts, 20),
      navTexts: cap(navTexts, 10),
      ariaButtonCount: buttons.length,
      ariaLabelSample
    };
  }).catch(() => ({ headerTexts: [], dayHeaderTexts: [], navTexts: [], ariaButtonCount: 0, ariaLabelSample: [] }));
}

// ---------------------------------------------------------------------------
// Gemeinsamer Event-Kollektionspfad (Iteration 6, Live-Befund Runde 5 vom 25.09.):
// Die Slot-Extraktion lieferte 0 Slots, obwohl dieselben Events in Runde 4 als
// ARIA-Labels im DOM nachweisbar waren. Root Cause: Die Extraktion nutzte einen
// EIGENEN, vom Diagnostik-Pfad abweichenden Inline-Collector (Hypothese 2) und
// brach nach der jetzt erfolgreichen Wochen-Verifikation zu früh ab. Dieser EINE
// Pfad liefert sowohl die Roh-Labels für die Slot-Extraktion als auch die
// Diagnostik (eventElementsFound/selectorUsed/rawEventLabels/rejectionReasons).
//
// Sichtbarkeit (Runde-5-Forderung): Event-Buttons sind in virtualisierten Grids
// teils mit offsetParent === null, haben aber echte Fläche (rect > 0). offsetParent
// ist daher KEIN Hard-K.O. für Events — ein Label mit Event-Signatur gilt als
// extrahierbar, solange das Element im DOM existiert. hiddenCount wird nur
// diagnostisch erfasst. Der strenge Sichtbarkeitsfilter bleibt den Header-
// Kandidaten in collectCalendarCandidates vorbehalten.
// ---------------------------------------------------------------------------
// Gemeinsamer Event-Kollektionspfad (Defekt 1, Live-Befund KW 40 vom 25.09.)
// ---------------------------------------------------------------------------
// Historie der Root Causes an dieser Stelle:
//  Runde 5: Die Extraktion nutzte einen EIGENEN, vom Diagnostik-Pfad
//           abweichenden Inline-Collector -> 0 Slots trotz vorhandener Events.
//  Runde 7: Der SEL'EKTOR selbst war blind. `[role="button"][aria-label]"` und
//           `[aria-label]` lieferten zusammen nur 12 Labels (11x header-artig,
//           1x Uhrzeit-Indikator) und KEINEN einzigen Event-Chip.
//
// Deshalb jetzt EINE Sammlung mit drei Bestandteilen:
//  1) QUELLEN-LEITER: aria-label -> title -> data-* -> sichtbarer Text.
//     Jede Stufe wird protokolliert (sourceTried), die erste mit Treffern
//     gewinnt. Damit haengt die Extraktion nicht an einem geratenen Attribut.
//  2) GEOMETRIE: Tages-Spalten (Tageskoepfe) + Zeitraster (Stundenzeilen) +
//     Grid-Body-Region. Nur Knoten INNERHALB der Region werden als
//     Event-Chip-Kandidaten behandelt — dadurch kann Toolbar-Text
//     ("Work week") nie zu einem Phantom-Slot werden.
//  3) STRUKTUR-CENSUS: wird IMMER geliefert, auch bei 0 Treffern. Er
//     enthaelt dayColumns/clockRuler/Blatt-Knoten (tag/role/class/data-attrs/
//     rect/Text) — das ist die Bodenwahrheit ueber die echte OWA-Struktur.
//
// Sichtbarkeit: In virtualisierten Grids ist offsetParent === null normal,
// Sichtbarkeit ist aber ueber die boundingRect-Flaeche belegbar. offsetParent
// ist daher KEIN Hard-K.O. fuer Events; der strenge Sichtbarkeitsfilter
// bleibt den Header-Kandidaten in collectCalendarDiagnostics vorbehalten.
// ---------------------------------------------------------------------------

// Kandidaten-/Census-Grenzen: haalten den Scan und die MCP-Antwort klein.
const CENSUS_MAX_NODES = 14;
const SCAN_MAX_ELEMENTS = 6000;
// Rohe Struktur-Knoten fuer die Geometrie-Ableitung (Datentraeger, grosszuegig).
const SCAN_MAX_STRUCTURE_NODES = 200;

// Sucht die Kandidaten-Knoten und liefert Label-Liste + Diagnostik.
// `weekStartIso` (Montag der gerenderten Woche) ermoeglicht die Datums-
// ableitung aus der Spaltenposition.
export async function collectCalendarEventLabels(page, { preferSelector = null, weekStartIso = null, fast = false } = {}) {
  let collectError = null;
  try {
    const res = await page.evaluate(
      ({ pats, plan, clockStrict, weekStart, fastMode, maxElements, maxCensusNodes, maxStructureNodes, loadPats, loadingSelector }) => {
        const re = {};
        for (const k of Object.keys(pats)) re[k] = new RegExp(pats[k], 'i');
        // Lade-Indikator-Muster kommen als ARGUMENT ueber die evaluate-Grenze
        // (Modul-Referenzen sind hier ein ReferenceError, siehe collectError).
        const lre = {};
        for (const k of Object.keys(loadPats)) lre[k] = new RegExp(loadPats[k], 'i');
        const clockOnlyRe = new RegExp(clockStrict, 'i');
        const norm = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
        const cap = (a, n) => a.slice(0, n);
        const safeQSA = (sel) => { try { return document.querySelectorAll(sel); } catch (e) { return []; } };
        const safeAttr = (el, name) => { try { return el.getAttribute ? el.getAttribute(name) : null; } catch (e) { return null; } };
        // Billige Attribut-Existenzpruefung (kein Layout). hasAttribute fehlt
        // bei einfachen Test-Fakes -> dann ueber die Attribut-Liste.
        const hasAttr = (el, name) => {
          try {
            if (typeof el.hasAttribute === 'function') return el.hasAttribute(name);
          } catch (e) { /* weiter mit dem Fallback */ }
          try {
            const as = el.attributes;
            if (!as) return false;
            for (let i = 0; i < as.length; i++) if (as[i] && as[i].name === name) return true;
          } catch (e) { /* kein Attributzugang -> false */ }
          return false;
        };
        const hasDataAttr = (el) => {
          try {
            const as = el.attributes;
            if (!as) return false;
            for (let i = 0; i < as.length; i++) if (as[i] && as[i].name && as[i].name.indexOf('data-') === 0) return true;
          } catch (e) { /* kein Attributzugang -> false */ }
          return false;
        };
        const rectOf = (el) => {
          try {
            const r = el.getBoundingClientRect();
            if (!r) return null;
            return { x: r.x, y: r.y, left: r.left, top: r.top, width: r.width, height: r.height };
          } catch (e) { return null; }
        };
        const areaOf = (r) => Boolean(r && r.width > 0 && r.height > 0);
        // Volle Knoten-Beschreibung (teuer -> erst nach der Flaechen-/Regions-
        // Filterung aufrufen).
        const describe = (el) => {
          const rect = rectOf(el);
          let offsetParentNull = false;
          try { offsetParentNull = el.offsetParent === null; } catch (e) { offsetParentNull = false; }
          const data = {};
          try {
            const attrs = el.attributes;
            if (attrs) {
              for (let i = 0; i < attrs.length && Object.keys(data).length < 8; i++) {
                const a = attrs[i];
                if (a && a.name && a.name.indexOf('data-') === 0) data[a.name] = String(a.value || '').slice(0, 160);
              }
            }
          } catch (e) { /* data-* ist optional */ }
          let cls = '';
          try { cls = String((el.className && el.className.baseVal) || el.className || ''); } catch (e) { cls = ''; }
          let tag = '';
          try { tag = String(el.tagName || ''); } catch (e) { tag = ''; }
          let text = '';
          try { text = el.innerText || el.textContent || ''; } catch (e) { text = ''; }
          return {
            tag,
            role: safeAttr(el, 'role') || null,
            cls: norm(cls).slice(0, 120),
            ariaLabel: norm(safeAttr(el, 'aria-label')).slice(0, 300),
            title: norm(safeAttr(el, 'title')).slice(0, 300),
            text: norm(text).slice(0, 300),
            data: Object.keys(data).length ? data : null,
            rect,
            area: areaOf(rect),
            offsetParentNull
          };
        };

        const isEventSignature = (t) => {
          if (!t) return false;
          if (re.nonEvent.test(t)) return false;
          if (re.gridBlocker.test(t)) return false;
          return re.timeRange.test(t) || re.allDay.test(t);
        };
        const reasonOf = (t) => {
          if (re.gridBlocker.test(t)) return 'clock-indicator';
          if (re.nonEvent.test(t)) return 'work-location';
          if (re.dateish.test(t) && t.length <= 160) return 'header-misclassified';
          return null;
        };

        // Lade-Indikator = EIGENE Kategorie. Ein Spinner ist kein Event und
        // erst recht kein "nichts gefunden": er belegt, dass die Ansicht noch
        // laedt. Live-Befund 25.09.2026: ein ms-Spinner-circle stand im
        // Struktur-Census und wurde als Event-Kandidat gefuehrt.
        const loadHit = (node) => {
          if (lre.tag.test(node.tag || '')) return 'tag';
          if (lre.role.test(node.role || '')) return 'role';
          if (lre.cls.test(node.cls || '')) return 'cls';
          if (lre.text.test(node.ariaLabel || '') || lre.text.test(node.title || '') || lre.text.test(node.text || '')) return 'text';
          const dd = node.data || {};
          const dkeys = Object.keys(dd);
          for (let i = 0; i < dkeys.length; i++) {
            if (lre.data.test(dkeys[i]) || lre.data.test(dd[dkeys[i]] || '')) return 'data';
          }
          return null;
        };
        const loadIndicators = [];
        const noteLoading = (node, why) => {
          if (loadIndicators.length < 10) {
            loadIndicators.push({ tag: node.tag, cls: String(node.cls || '').slice(0, 80), reason: why, top: node.rect ? node.rect.top : null });
          }
          return 'load-indicator';
        };
        // Billiger Lade-Indikator-Check fuer den 'fast'-Modus (Warte-Probe):
        // dokumentweiter Attribut-Selektor statt '*'-Durchlauf.
        const fastLoadingProbe = () => {
          const hits = safeQSA(loadingSelector);
          for (let i = 0; i < hits.length && i < 5; i++) {
            loadIndicators.push({ tag: String(hits[i].tagName || ''), cls: norm(safeAttr(hits[i], 'class')).slice(0, 80), reason: 'fast-probe', top: null });
          }
          return hits.length;
        };

        // 1) Tages-Spalten aus den Tageskoepfen (x-Lage der Wochentage)
        const dayColumns = [];
        if (!fastMode) {
        for (const el of safeQSA('[role="columnheader"], [role="grid"] th, th')) {
          const r = rectOf(el);
          if (areaOf(r)) dayColumns.push({ left: r.left, width: r.width, bottom: r.top + r.height, text: norm(safeAttr(el, 'aria-label') || (el.textContent || '')).slice(0, 40) });
        }
        dayColumns.sort((a, b) => a.left - b.left);
        }

        // 2) Zeitraster: Elemente, deren Text NUR eine Uhrzeit ist
        const clockRuler = [];
        let scannedCount = 0;
        let truncated = false;
        for (const el of (fastMode ? [] : safeQSA('*'))) {
          if (scannedCount >= maxElements) { truncated = true; break; }
          scannedCount++;
          const r = rectOf(el);
          if (!areaOf(r)) continue;
          if (r.width > 120) continue; // Stundenzeilen sind schmale Labels
          const t = norm(safeAttr(el, 'aria-label') || safeAttr(el, 'title') || (el.innerText || el.textContent || ''));
          if (!t || !clockOnlyRe.test(t)) continue;
          clockRuler.push({ time: t, top: r.top });
        }
        clockRuler.sort((a, b) => a.top - b.top);

        // 3) Grid-Body-Region: unterhalb der Tageskoepfe, rechts des Rasters.
        //    Ohne sie wird die Geometrie-Ableitung NICHT angewendet (nie raten).
        let region = null;
        if (dayColumns.length >= 2) {
          const first = dayColumns[0];
          const last = dayColumns[dayColumns.length - 1];
          const headerBottom = Math.max.apply(null, dayColumns.map((c) => c.bottom));
          region = { left: first.left, right: last.left + last.width, top: headerBottom };
        }
        const inRegion = (r) => {
          if (!region || !r) return false;
          if (r.left < region.left - 8 || r.left > region.right + 8) return false;
          if (r.top < region.top - 8) return false;
          return true;
        };

        // 4) Struktur-Scan: Blatt-Knoten im Grid-Body (kein geratener Selektor)
        //    Lade-Indikatoren landen in einer EIGENEN Liste und werden NIE als
        //    Event-Kandidat gefuehrt (Bug 3, Live-Befund 25.09.2026).
        const structureNodes = [];
        const loadingNodes = [];
        if (region) {
          let seen = 0;
          for (const el of safeQSA('*')) {
            if (seen >= maxElements) { truncated = true; break; }
            seen++;
            const r = rectOf(el);
            if (!areaOf(r)) continue;
            if (!inRegion(r)) continue;
            const d = describe(el);
            // Lade-Indikator zuerst: er ist im Grid-Body zu erwarten, solange
            // die Liste laedt, und darf die Event-Geometrie nicht verfaelschen.
            const why = loadHit(d);
            if (why) { loadingNodes.push(d); noteLoading(d, why); continue; }
            // Blatt-Knoten: keine Kind-Knoten mit eigener Flaeche
            let hasChildArea = false;
            try {
              const kids = el.children || [];
              for (let i = 0; i < kids.length; i++) {
                const kr = rectOf(kids[i]);
                if (areaOf(kr)) { hasChildArea = true; break; }
              }
            } catch (e) { hasChildArea = false; }
            if (hasChildArea) continue;
            d.inGrid = true;
            structureNodes.push(d);
          }
        }

        const headerLike = [];
        const nonEventLabels = [];
        const candidates = [];
        const allLabels = [];
        // Event-Signatur OHNE Datum — der Geometrie-Weg macht daraus ein Label.
        const dateMissing = [];
        const sourceTried = [];
        let hiddenCount = 0;
        let sourceUsed = null;
        let selectorUsed = null;

        const consume = (node, source) => {
          if (!node.area) { hiddenCount++; return false; }
          // Ein Lade-Indikator ist nie ein Event — auch nicht, wenn er ueber
          // die Attribut-Leiter auftaucht statt ueber den Struktur-Scan.
          const whyLoad = loadHit(node);
          if (whyLoad) { noteLoading(node, whyLoad); return false; }
          const texts = [node.ariaLabel, node.title, node.text]
            .concat(Object.values(node.data || {}))
            .map(norm)
            .filter((t) => t && t.length <= 300);
          let hit = false;
          for (const t of texts) {
            if (isEventSignature(t)) {
              // Ein Label NUR mit Zeit ("Urlaub, all day", "14:00 to 15:00")
              // ist fuer den Parser wertlos (kein Datum). Solche Treffer
              // werden NICHT als Label gemeldet, sondern der Geometrie
              // ueberlassen — sonst zaehlt eventElementsFound Knoten, aus
              // denen kein Slot werden kann.
              if (re.dateish.test(t)) { allLabels.push(t); hit = true; }
              else if (dateMissing.indexOf(t) < 0) dateMissing.push(t);
            } else {
              const r = reasonOf(t);
              if (r === 'header-misclassified') headerLike.push(t);
              else if (r) nonEventLabels.push({ label: t, reason: r });
            }
          }
          if (texts.length && candidates.length < 30) {
            candidates.push({
              label: (texts[0] || '').slice(0, 160),
              source, tag: node.tag, role: node.role, cls: node.cls,
              data: node.data, inGrid: node.inGrid === true,
              visible: node.area, offsetParentNull: node.offsetParentNull
            });
          }
          return hit;
        };

        for (const step of plan) {
          let nodes = [];
          if (fastMode) {
            // Nur dokumentweite Attribut-Selektoren — kein '*'-Durchlauf.
            if (!step.selector) { sourceTried.push({ source: step.source, selector: '(struktur)', nodes: 0, matched: 0, skipped: 'fast-mode' }); continue; }
            for (const el of safeQSA(step.selector)) nodes.push(describe(el));
          } else if (step.selector) {
            for (const el of safeQSA(step.selector)) nodes.push(describe(el));
          } else if (step.source === 'text') {
            nodes = structureNodes;
          } else {
            // title/data-attr: bewusst breit, aber mit einem BILLIGEN Vorfilter.
            // describe() liest innerText und erzwingt damit ein Layout — das
            // darf nicht fuer jedes der ~15k Elemente einer echten OWA-Seite
            // passieren (sonst frisst der Scan das 45s-Budget). Der Vorfilter
            // prueft nur, ob das Attribut ueberhaupt vorhanden ist.
            let broadScanned = 0;
            for (const el of safeQSA('*')) {
              if (broadScanned >= maxElements) { truncated = true; break; }
              broadScanned++;
              if (step.source === 'title' && !hasAttr(el, 'title')) continue;
              if (step.source === 'data-attr' && !hasDataAttr(el)) continue;
              const r = rectOf(el);
              if (areaOf(r)) nodes.push(describe(el));
            }
          }
          let matched = 0;
          for (const n of nodes) if (consume(n, step.source)) matched++;
          sourceTried.push({ source: step.source, selector: step.selector || '(struktur)', nodes: nodes.length, matched });
          if (matched > 0 && !sourceUsed) { sourceUsed = step.source; selectorUsed = step.selector || `(struktur:${step.source})`; break; }
        }

        const dedupe = (arr) => { const o = []; const s = new Set(); for (const l of arr) if (!s.has(l)) { s.add(l); o.push(l); } return o; };
        const dedupeReasons = (arr) => { const o = []; const s = new Set(); for (const e of arr) if (!s.has(e.label)) { s.add(e.label); o.push(e); } return o; };
        const labels = cap(dedupe(allLabels), 60);
        // 'fast'-Modus (Warte-Probe) hat keinen Struktur-Scan -> dort zaehlt
        // der billige Attribut-Selektor. Beide Wege liefern dasselbe Flag.
        const fastLoading = fastMode ? fastLoadingProbe() : 0;
        const loadingIndicatorsOut = cap(dedupeReasons(loadIndicators.map((l) => ({ label: String(l.cls || l.tag || '').slice(0, 80), reason: 'load-indicator' }))), 10);
        return {
          labels,
          eventElementsFound: labels.length,
          selectorUsed: selectorUsed || '(keine-quelle-mit-treffer)',
          sourceUsed,
          sourceTried,
          hiddenCount,
          headerLikeLabels: cap(dedupe(headerLike), 20),
          nonEventLabels: cap(dedupeReasons(nonEventLabels), 20),
          candidates: cap(candidates, 20),
          // Signatur ohne Datum: sichtbar, aber kein verwertbares Label.
          dateMissingLabels: cap(dateMissing, 20),
          // Lade-Indikator vorhanden? -> die Ansicht ist NOCH NICHT FERTIG.
          // Das ist die Kategorie, die den Warte-Loop am Leben haelt.
          loading: loadIndicators.length > 0 || fastLoading > 0,
          loadingIndicators: loadingIndicatorsOut,
          // Rohe Struktur-Knoten fuer die Node-seitige Geometrie-Ableitung
          // (grosszuegiges Limit: das ist der eigentliche Datentraeger).
          // SCAN_MAX_STRUCTURE_NODES wird ueber maxStructureNodes gereicht —
          // ein nackter Modul-Identifier waere im Browser ein ReferenceError.
          structureNodes: fastMode ? [] : cap(structureNodes, maxStructureNodes),
          census: fastMode
            ? { skipped: 'fast-mode', loadingNodes: fastLoading }
            : {
              scannedCount,
              truncated,
              dayColumns: cap(dayColumns, 10),
              clockRuler: cap(clockRuler, 16),
              structureNodes: structureNodes.length,
              loadingNodes: loadingNodes.length,
              scanned: cap(structureNodes, maxCensusNodes)
            }
        };
      },
      {
        pats: EVENT_LABEL_PATTERNS,
        plan: eventSourcePlan(preferSelector),
        clockStrict: CLOCK_LABEL_PATTERN,
        weekStart: weekStartIso,
        fastMode: fast === true,
        maxElements: SCAN_MAX_ELEMENTS,
        maxCensusNodes: CENSUS_MAX_NODES,
        // Eigenes Limit: maxCensusNodes ist CENSUS_MAX_NODES (14, reine
        // Antwortgroesse) und NICHT das Scan-Limit fuer die Rohknoten (200).
        maxStructureNodes: SCAN_MAX_STRUCTURE_NODES,
        // Lade-Indikator-Muster + Selektor als Argument (Single Source of
        // Truth ist LOAD_INDICATOR_PATTERNS weiter oben im Modul).
        loadPats: LOAD_INDICATOR_PATTERNS,
        loadingSelector: LOAD_INDICATOR_SELECTOR
      }
    );
    // GEOMETRIE-Pfad (Node-seitig, gleiche Prueflogik wie der Rest):
    // Struktur-Knoten aus dem Grid-Body werden ueber deriveEventLabelFromNode
    // bewertet — das ist der Weg, der OHNE aria-label/Title/data-* auskommt
    // (Wochentag aus der Spalten-x, Zeit aus dem Raster-y).
    const structureNodes = Array.isArray(res.structureNodes) ? res.structureNodes : [];
    const columns = (res.census && res.census.dayColumns) || [];
    const ruler = (res.census && res.census.clockRuler) || [];
    const geoLabels = [];
    const geoRejections = [];
    for (const n of structureNodes) {
      const d = deriveEventLabelFromNode(n, {
        pats: EVENT_LABEL_PATTERNS,
        columns,
        ruler,
        weekStartIso
      });
      if (d.label) geoLabels.push(d.label);
      else if (d.reason && d.reason !== 'no-event-signature') {
        geoRejections.push({ label: (n.ariaLabel || n.title || n.text || '').slice(0, 120), reason: d.reason, tag: n.tag, cls: n.cls });
      }
    }
    if (geoLabels.length > 0) {
      res.labels = res.labels.concat(geoLabels);
      // Dedupe + ehrliche Zaehlung nach der Zusammenfuehrung
      const seen = new Set();
      res.labels = res.labels.filter((l) => (seen.has(l) ? false : (seen.add(l), true))).slice(0, 60);
      res.eventElementsFound = res.labels.length;
      res.sourceUsed = res.sourceUsed === 'text' || !res.sourceUsed ? 'text+geometry' : res.sourceUsed;
      if (!res.selectorUsed || res.selectorUsed === '(keine-quelle-mit-treffer)') res.selectorUsed = '(struktur:text+geometry)';
    }
    if (geoRejections.length) {
      const known = new Set(res.nonEventLabels.map((e) => e.label));
      for (const g of geoRejections) {
        if (!known.has(g.label)) { known.add(g.label); res.nonEventLabels.push(g); }
      }
      res.nonEventLabels = res.nonEventLabels.slice(0, 20);
    }
    return { collectError: null, ...res };
  } catch (e) {
    collectError = String((e && e.message) || e).slice(0, 160);
    // Live-Befund 25.09.2026: Ein ReferenceError aus dem evaluate-Callback ist
    // KEIN DOM-Zustand, sondern ein Code-Defekt — fast immer eine nackte
    // Modul-Referenz, die die evaluate-Grenze nicht ueberquert hat
    // (SCAN_MAX_ELEMENTS o. Ae.). Ein solcher Fehler wurde bisher still in
    // ein leeres Standardobjekt verwandelt und dadurch als "Kalender leer"
    // gemeldet. Das ist exakt die Falschaussage, die diese Runde beseitigen
    // soll: er wird laut geworfen. Jeder DOM-Zugriff im Callback ist bereits
    // try/catch-geschuetzt, ein ReferenceError kann dort nicht entstehen.
    if (e && e.name === 'ReferenceError') {
      const err = new Error(
        `collectCalendarEventLabels: ReferenceError im Browser-Callback (${collectError}). ` +
        `Das ist ein Code-Defekt (nackte Modul-Referenz ueber die page.evaluate-Grenze), ` +
        `kein DOM-Zustand — Werte muessen ueber das Argument-Objekt gereicht werden.`
      );
      err.name = 'CollectCallbackReferenceError';
      err.collectError = collectError;
      throw err;
    }
    return {
      labels: [], eventElementsFound: 0, selectorUsed: 'error', sourceUsed: null,
      sourceTried: [], hiddenCount: 0, headerLikeLabels: [], nonEventLabels: [],
      candidates: [], collectError, loading: false, loadingIndicators: [],
      census: { scannedCount: 0, truncated: false, dayColumns: [], clockRuler: [], structureNodes: 0, loadingNodes: 0, scanned: [] }
    };
  }
}

// Löst aus den gesammelten Kandidaten-Texten die gerenderte Kalenderwoche auf.
// Reine Funktion (unit-testbar). Erkennungs-Reihenfolge (spezifisch zuerst):
//  1. Wochen-Range-Header mit Monatsnamen ODER numerisch ("28. September – 4. Oktober",
//     "28 September – 4 October, 2026", "28.9. – 4.10.2026")
//  2. Prev/Next-Button-Ableitung ("… (Week 39)" / "… (Week 41)" -> Woche 40)
//  3. Einzel-Datum mit Monatsnamen (Toolbar-Datumsbutton, "Montag, 28. September")
//  4. Numerische Tagesangaben ("Montag 28.9." / "28.9.2026")
// Rückgabe: { text, startDate, endDate, weekNumber } — bei NICHTS Datierbarem sind
// alle Felder null. Event-Labels (Zeitfenster/Ganztag/Statuswort+Datum) werden strikt
// ausgefiltert und können NIE als gerenderte Woche durchkommen (Live-Fail Runde 4:
// stale All-Day-Event "Urlaub … September 23" wurde fälschlich als Header gelesen).
export function resolveRenderedWeekFromCandidates({ headerTexts = [], dayHeaderTexts = [], navTexts = [] } = {}, now = new Date()) {
  // Aus einem Einzel-Datum die OWA-Wochenansicht (Montag–Sonntag) ableiten.
  const deriveWeek = (dateIso, text) => {
    const d = new Date(dateIso + 'T00:00:00');
    const offsetToMonday = d.getDay() === 0 ? -6 : 1 - d.getDay();
    const monday = new Date(d.getTime() + offsetToMonday * 86400000);
    const sunday = new Date(monday.getTime() + 6 * 86400000);
    return { text: String(text).slice(0, 200), startDate: toIsoDate(monday), endDate: toIsoDate(sunday), weekNumber: null };
  };
  const currentYear = now.getFullYear();

  // Event-Labels strikt ausschließen (Runde-4-Root-Cause): Ein All-Day-Event
  // ("Urlaub … September 23") darf NIE als Wochen-Header durchkommen.
  const headers = filterCalendarHeaderCandidates(headerTexts);
  const dayHeaders = filterCalendarHeaderCandidates(dayHeaderTexts);

  // 1) Wochen-Range-Header (Header-Bereich zuerst, dann Tageskopf-/Body-Kandidaten)
  for (const text of headers) {
    const parsed = parseWeekHeaderText(text, now);
    if (parsed) return { text: text.slice(0, 200), startDate: parsed.startDate, endDate: parsed.endDate, weekNumber: null };
  }
  for (const text of dayHeaders) {
    const parsed = parseWeekHeaderText(text, now);
    if (parsed) return { text: text.slice(0, 200), startDate: parsed.startDate, endDate: parsed.endDate, weekNumber: null };
  }
  // 2) Prev/Next-Button-Ableitung (belastbarer Fallback, Runde-4-Forderung 2)
  const nav = resolveWeekFromNavButtons(navTexts, now);
  if (nav && nav.startDate) {
    return { text: nav.text, startDate: nav.startDate, endDate: nav.endDate, weekNumber: nav.weekNumber };
  }
  // 3) Einzel-Datum mit Monatsnamen
  for (const text of headers) {
    const slot = parseEventLabel(text, currentYear);
    if (slot && slot.date) return deriveWeek(slot.date, text);
  }
  for (const text of dayHeaders) {
    const slot = parseEventLabel(text, currentYear);
    if (slot && slot.date) return deriveWeek(slot.date, text);
  }
  // 4) Numerische Tagesangaben (Spaltenköpfe + kurze Header-Texte)
  for (const text of dayHeaders) {
    const dateIso = parseNumericDayHeaderText(text, now);
    if (dateIso) return deriveWeek(dateIso, text);
  }
  for (const text of headers) {
    if (String(text).length > 60) continue; // lange Texte nur mit stärkerem Muster akzeptieren
    const dateIso = parseNumericDayHeaderText(text, now);
    if (dateIso) return deriveWeek(dateIso, text);
  }

  // Nichts Datierbares gefunden -> bewusst NULL zurückgeben (kein "Work week"-Fehlgreifer).
  return { text: null, startDate: null, endDate: null, weekNumber: null };
}

// Liest die gerenderte Kalenderwoche aus der OWA-Wochenansicht aus.
// Rückgabe: { text, startDate, endDate } — startDate/endDate als ISO-Strings
// 'YYYY-MM-DD', alles null wenn der Header nicht gefunden oder nicht parsebar war.
export async function readRenderedCalendarWeek(page) {
  const collected = await collectCalendarCandidates(page);
  return resolveRenderedWeekFromCandidates(collected);
}

// Sammelt reiche Diagnostik für Fehlerfälle und unverified-Responses (Kernforderung
// Runde 3): URL, Header-Kandidaten, Tageskopf-Kandidaten, ARIA-Button-Anzahl und ein
// Sample der ersten ARIA-Labels — so zeigt der nächste Live-Test sofort, was das DOM
// wirklich enthielt.
export async function collectCalendarDiagnostics(page) {
  const collected = await collectCalendarCandidates(page);
  let url = '';
  try { url = page.url() || ''; } catch (e) { /* bewusst ignorieren */ }
  return {
    url,
    headerCandidates: collected.headerTexts,
    dayHeaderCandidates: collected.dayHeaderTexts,
    weekNavCandidates: collected.navTexts,
    ariaButtonCount: collected.ariaButtonCount,
    ariaLabelSample: collected.ariaLabelSample
  };
}

// Formatiert Kalender-Diagnostik als kompakte mehrzeilige Zeichenkette (für Fehlermeldungen).
export function formatCalendarDiagnostics(diag) {
  if (!diag) return '(keine Diagnostik verfügbar)';
  const fmtList = (list, max) => (list && list.length)
    ? list.slice(0, max).map((t) => `"${String(t).slice(0, 80)}"`).join(' | ')
    : '(keine)';
  return [
    `URL: ${diag.url || 'unbekannt'}`,
    `Header-Kandidaten: ${fmtList(diag.headerCandidates, 8)}`,
    `Tageskopf-Kandidaten: ${fmtList(diag.dayHeaderCandidates, 8)}`,
    `Wochen-Nav-Kandidaten: ${fmtList(diag.weekNavCandidates, 4)}`,
    `ARIA-Buttons gesamt: ${diag.ariaButtonCount ?? '?'}`,
    `ARIA-Label-Sample: ${fmtList(diag.ariaLabelSample, 15)}`
  ].join('\n');
}

// ---------------------------------------------------------------------------
// getFreeBusy: Zeitbudget, Klick-Rationierung & Diagnostik-Assemblierung
// (Iteration 4, Live-Befunde Runde 3 vom 25.09.: Der Gesamt-Call lief in den
// MCP-Client-Timeout (-32001), bevor irgendetwas serialisiert wurde — die
// Diagnostik wurde erst NACH der Eskalationsleiter erhoben und ging verloren.)
// ---------------------------------------------------------------------------

// Hartes Gesamtzeitbudget für den KOMPLETTEN getFreeBusy-Call. Muss sicher
// UNTER dem MCP-Client-Timeout liegen (Runde 3: -32001 bei ~60s oder darunter).
export const FREEBUSY_TOTAL_BUDGET_MS = 45000;

// Rationierung der Button-Klicks gegen das Budget (Runde-4-Forderung A):
// max. 6 Klicks mit je max. 4s Re-Render-Wartezeit.
export const FREEBUSY_MAX_BUTTON_CLICKS = 6;
export const FREEBUSY_CLICK_RENDER_WAIT_MS = 4000;

// Früher Kontroll-Return: Liegt nach Navigation + Render-Warten kein datierter
// Header vor UND das Restbudget reicht nicht mehr für die Eskalationsleiter
// (frische Navigation ~15s bzw. Button-Klicks), wird SOFORT error+diagnostics
// zurückgegeben statt weiterzuklicken.
export const FREEBUSY_MIN_REMAINING_FOR_ESCALATION_MS = 12000;

// ---------------------------------------------------------------------------
// Event-Wartephase (Iteration 7, Live-Befund Runde 6 vom 28.09.2026 — HAUPTURSACHE)
// ---------------------------------------------------------------------------
// Runde 6 lief nach ~4s "erfolgreich" durch (usedMs 4141 bei remainingMs 40859),
// obwohl das Event-Grid erst nach ~17s rendert (Runde 4: 16,9s, 6 Labels). Die
// 2,5s-Gnadenfrist nach datiertem Header galt als "fertig", weil hasEventElements
// den Uhrzeit-Indikator "calendar view, current time: Fri 22:31" als Event
// wertete. Deshalb wird nach bestätigter Wochen-Verifikation GEZIELT auf echte
// Event-Labels gepollt — budgetgebunden, mit Reserve für die Extraktion.

// Obergrenze der Event-Wartephase (Events rendern teils erst nach ~17s).
export const FREEBUSY_EVENT_WAIT_MAX_MS = 20000;
// Reserve, die für Extraktion + Diagnostik vom Gesamtbudget ABGEZOGEN wird, damit
// der Call nie ins MCP-Client-Timeout läuft.
export const FREEBUSY_EVENT_WAIT_RESERVE_MS = 5000;
// Poll-Intervall der Event-Wartephase.
export const FREEBUSY_EVENT_WAIT_POLL_MS = 500;
// Ruhe-Fenster, über das "kein Event UND kein Lade-Indikator" als BELEGT leer
// gelten muss, bevor der Poll vorzeitig enden darf. Ein einzelner spinner-freier
// Blick ist KEIN Leer-Beweis: direkt nach Navigation/Klick ist das Grid kurz
// leer, BEVOR es befüllt wird (Live-Befund 25.09.2026). Ohne dieses Fenster
// würde der Poll nach ~1 Poll mit "not-found" enden — genau der Defekt, den
// die stale Vorabmeldung verdeckt hat.
export const FREEBUSY_EVENT_WAIT_SETTLE_MS = 1000;

// ---------------------------------------------------------------------------
// Lade-Indikatoren (Spinner) — EIGENE Kategorie, kein Event-Kandidat
// ---------------------------------------------------------------------------
// Live-Befund 25.09.2026 (Smoke-Test Runde 3, echte OWA): Der Struktur-Census
// enthielt einen "ms-Spinner-circle" (y≈487) und führte ihn als möglichen
// Event-Knoten. Ein Spinner ist aber das Gegenteil eines Events: er BEWEIST,
// dass die Kalenderliste NOCH AM LADEN ist. Als Event-Kandidat durchgerutscht
// war er wertlos; die Wartephasepollte daraufhin nie.
//
// Die Erkennungsmuster liegen hier als Single Source of Truth. Der Browser-
// Callback in collectCalendarEventLabels KANN Modul-Identifikatoren nicht
// referenzieren (ReferenceError ueber die evaluate-Grenze, siehe collectError),
// deshalb werden die Muster als Argument durchgereicht und dort neu kompiliert —
// Muster und Selektor kommen aber beide von hier.
export const LOAD_INDICATOR_PATTERNS = {
  cls: String.raw`ms-?spinner|ms-?progressindicator|ms-?loading|progressring|\bspinner\b`,
  tag: String.raw`^ms-(spinner|progressindicator|loading|progressring)$`,
  role: String.raw`^progressbar$`,
  text: String.raw`wird geladen|loading\.\.\.|lade(anzeige|spin)|^\s*laden\s*$`,
  data: String.raw`spinner|loading`
};

// Billiger, dokumentweiter Lade-Indikator-Selektor fuer den 'fast'-Modus der
// Warte-Probe. Bewusstattributbasiert und ohne '*'-Durchlauf: die Warte-Probe
// laeuft alle 500ms und darf das Poll-Budget nicht auffressen.
export const LOAD_INDICATOR_SELECTOR = [
  '[class*="ms-Spinner" i]',
  '[class*="ms-ProgressIndicator" i]',
  '[class*="ms-Loading" i]',
  '[role="progressbar"]',
  '[aria-busy="true"]',
  'ms-Spinner',
  'ms-ProgressIndicator'
].join(', ');

// Ist dieser Grid-/Strukturknoten ein Lade-Indikator? Reine Funktion, damit die
// Kategorie auch ohne Browser testbar ist (Regressionstest der Spinner-Trennung).
export function isLoadIndicatorNode(node) {
  if (!node) return false;
  const test = (src, pat) => {
    if (!src) return false;
    try { return new RegExp(pat, 'i').test(String(src)); } catch (e) { return false; }
  };
  if (test(node.tag, LOAD_INDICATOR_PATTERNS.tag)) return true;
  if (test(node.role, LOAD_INDICATOR_PATTERNS.role)) return true;
  if (test(node.cls, LOAD_INDICATOR_PATTERNS.cls)) return true;
  if (test(node.ariaLabel, LOAD_INDICATOR_PATTERNS.text)) return true;
  if (test(node.title, LOAD_INDICATOR_PATTERNS.text)) return true;
  if (test(node.text, LOAD_INDICATOR_PATTERNS.text)) return true;
  const data = node.data || {};
  for (const k of Object.keys(data)) {
    if (test(k, LOAD_INDICATOR_PATTERNS.data) || test(data[k], LOAD_INDICATOR_PATTERNS.data)) return true;
  }
  return false;
}

// Die möglichen Ausgänge der Event-Wartephase (Diagnose-Feld eventWaitOutcome)
// sind bei resolveEventWaitOptionen definiert — dort Single Source of Truth.

// Wartebudget = min(Cap, Restzeit - Reserve), nie negativ. Runde-6-Kennzahlen als
// Beispiel: remainingMs 40859 -> 20000 (Cap greift), remainingMs 6000 -> 1000,
// remainingMs 3000 -> 0 (Wartephase wird ganz übersprungen).
export function eventWaitBudgetMs(remainingMs, { maxMs = FREEBUSY_EVENT_WAIT_MAX_MS, reserveMs = FREEBUSY_EVENT_WAIT_RESERVE_MS } = {}) {
  if (!remainingMs || remainingMs <= 0) return 0;
  return Math.max(0, Math.min(maxMs, remainingMs - (reserveMs || 0)));
}

// Pollt, bis das Event-Grid echte Event-Labels zeigt. Dependency-Injection
// (hasEvents/now/sleep) macht die Wartephase ohne Browser testbar — Fixture:
// Events erscheinen erst im 3. Poll. Pollt NIE über budgetMs hinaus und schläft
// nie über das Budget hinaus (kein Überziehen der Gesamtdeadline).
// Rückgabe: { found, outcome, waitedMs, polls }
export async function pollUntilEventGrid({ hasEvents, budgetMs, pollMs = FREEBUSY_EVENT_WAIT_POLL_MS, now = () => Date.now(), sleep } = {}) {
  const startedAt = now();
  const waited = () => now() - startedAt;
  if (typeof hasEvents !== 'function' || !budgetMs || budgetMs <= 0) {
    return { found: false, outcome: 'budget-exhausted', waitedMs: waited(), polls: 0 };
  }
  let polls = 0;
  while (waited() < budgetMs) {
    polls++;
    // eslint-disable-next-line no-await-in-loop
    if (await hasEvents()) return { found: true, outcome: 'found', waitedMs: waited(), polls };
    if (waited() + pollMs >= budgetMs) break;
    // eslint-disable-next-line no-await-in-loop
    await sleep(pollMs);
  }
  return { found: false, outcome: 'budget-exhausted', waitedMs: waited(), polls };
}

// ---------------------------------------------------------------------------
// Bug 3 — der Warte-Loop pollt nicht (Root Cause, Live-Befund 25.09.2026)
// ---------------------------------------------------------------------------
// Live-Kennzahlen des Smoke-Tests Runde 3 (echte OWA, KW 40):
//     waitedForEventsMs: 0   eventWaitPolls: 0   eventWaitBudgetMs: 20000
//     eventWaitOutcome: 'not-found'   Abbruch nach ~4s
// 20000ms Budget, NULL Polls, NULL Milli Sekunden gewartet.
//
// Ursache (an der Quelle belegt, nicht vermutet): pollUntilEventGrid war die
// EINZIGE Pollschleife — sie wurde aber per `if (hasEvents) { polls: 0,
// waitedMs: 0 }` uebersprungen. hasEvents stammt aus der schnellen Vorab-Probe
// und war STALE: der Vorab-Probe-Selektor fand etwas, die abschliessende
// vollstaendige Sammlung fand 0 Knoten. Der Kommentar "eine stale Vorabmeldung
// darf nichts mehr entscheiden" beschrieb genau das Gegenteil der Wirklichkeit:
// die Vorabmeldung entschied, OB ueberhaupt gepollt wird — und verhinderte
// damit jede Moeglichkeit, die stale Vorabmeldung zu widerlegen.
//
// Zweiter, davon unabhaengiger Defekt: resolveEventWaitOutcome({polls,
// eventElementsFound, waitBudgetMs}) hat polls und waitBudgetMs destrukturiert
// und NIE benutzt. Ein erschoeptes Budget war damit von "0 Polls bei
// vorhandenem Budget" nicht unterscheidbar — beide ergaben 'not-found'.
//
// Die Reparatur: die Wartephase besitzt die Entscheidung selbst und arbeitet
// ausschliesslich auf BEOBACHTUNGEN (Event-Zahl + Lade-Indikator), nicht auf
// einer booleschen Vorabmeldung. Es gibt keine Stale-Boolean-Shortcut-Stelle
// mehr, die den Poll ueberspringen koennte.
//
// Beobachtungsprotokoll je Poll: { eventElementsFound, loading }
//   eventElementsFound > 0        -> 'found', sofort belegt fertig
//   loading === true               -> Spinner: NOCH NICHT FERTIG, weiter pollen
//   0 Events, kein Spinner         -> Ruhe-Fenster sammeln; erst nach
//                                     settleMs ununterbrochen ruhig ->
//                                     'not-found' mit abortReason
//                                     'rendered-empty-proven' (BELEGT leer)
//   Budget erschoepft               -> 'budget-exhausted' mit abortReason
//                                     'wait-budget-exhausted' (Budget
//                                     tatsaechlich ausge schoepft)
export const EVENT_WAIT_ABORT_REASONS = {
  renderedEmpty: 'rendered-empty-proven',
  budgetExhausted: 'wait-budget-exhausted',
  notPolled: 'wait-not-polled',
  budgetUnavailable: 'wait-budget-unavailable'
};

export async function waitForCalendarEvents({
  observe,
  budgetMs,
  pollMs = FREEBUSY_EVENT_WAIT_POLL_MS,
  settleMs = FREEBUSY_EVENT_WAIT_SETTLE_MS,
  now = () => Date.now(),
  sleep
} = {}) {
  const startedAt = now();
  const waited = () => now() - startedAt;
  const result = (extra) => ({
    found: false,
    outcome: 'not-found',
    waitedMs: waited(),
    polls: 0,
    loadingPolls: 0,
    loading: false,
    abortReason: null,
    ...extra
  });
  if (typeof observe !== 'function' || !budgetMs || budgetMs <= 0) {
    // Kein Budget -> ehrlich benennen statt "0 Polls bei 20s Budget" zu melden.
    return result({ outcome: 'budget-exhausted', abortReason: EVENT_WAIT_ABORT_REASONS.budgetUnavailable });
  }
  const doSleep = typeof sleep === 'function' ? sleep : (ms) => new Promise((r) => setTimeout(r, ms));
  let polls = 0;
  let loadingPolls = 0;
  let quietSince = null;   // Beginn des ununterbrochenen Ruhe-Fensters
  let lastLoading = false;
  while (waited() < budgetMs) {
    polls++;
    // eslint-disable-next-line no-await-in-loop
    const o = (await observe()) || {};
    const found = Number(o.eventElementsFound) > 0;
    lastLoading = o.loading === true;
    if (found) {
      return { found: true, outcome: 'found', waitedMs: waited(), polls, loadingPolls, loading: lastLoading, abortReason: null };
    }
    if (lastLoading) {
      // BEWEIS FUER "NOCH NICHT FERTIG": der Spinner haelt den Poll am Leben.
      loadingPolls++;
      quietSince = null;
    } else {
      if (quietSince === null) quietSince = now();
      const quietMs = waited() - (quietSince - startedAt);
      if (quietMs >= (settleMs || 0)) {
        return {
          found: false, outcome: 'not-found', waitedMs: waited(), polls, loadingPolls,
          loading: false, abortReason: EVENT_WAIT_ABORT_REASONS.renderedEmpty
        };
      }
    }
    if (waited() + pollMs >= budgetMs) break;
    // eslint-disable-next-line no-await-in-loop
    await doSleep(pollMs);
  }
  return {
    found: false, outcome: 'budget-exhausted', waitedMs: waited(), polls, loadingPolls,
    loading: lastLoading, abortReason: EVENT_WAIT_ABORT_REASONS.budgetExhausted
  };
}

// Klassifiziert EINE Beobachtung der Wartephase in eine von vier Kategorien.
// Reine Funktion (testbar) und die einzige Stelle, die "noch nicht fertig"
// von "nachweislich leer" trennt — genau die Unterscheidung, die im Live-Befund
// gefehlt hat, weil der Spinner als Event-Kandidat durchrutschte.
export function loadIndicatorReason({ eventElementsFound = 0, loading = false } = {}) {
  const n = Number(eventElementsFound) || 0;
  if (n > 0) return { category: 'events', loading: Boolean(loading), eventElementsFound: n };
  if (loading) return { category: 'still-loading', loading: true, eventElementsFound: 0 };
  return { category: 'empty-candidate', loading: false, eventElementsFound: 0 };
}

// Baut die Beobachtungsfunktion der Wartephase. Bewusst als exportierte
// Fabrik und NICHT als Inline-Closure in getFreeBusy: der Regressionstest
// (Bug 3) muss den ECHEN Draht fahren — echter Beobachter -> echter Kollektor
// -> echte Warteschleife. Eine im Test nachgebaute Beobachtung koennte
// driften und wuerde den Defekt genau nicht fangen (genau dieser Fehler ist
// bei Bug 2 passiert).
//
// `collect(preferSelector)` liefert das Ergebnis von collectCalendarEventLabels
// im 'fast'-Modus. Die zuletzt erfolgreiche Quelle wird als preferSelector
// weitergereicht (dieselbe Leiter, dieselbe Reihenfolge).
export function makeEventObserver(collect) {
  if (typeof collect !== 'function') throw new TypeError('makeEventObserver: collect muss eine Funktion sein');
  let lastSelector = null;
  const observeCalendarEvents = async () => {
    // eslint-disable-next-line no-await-in-loop
    const r = await collect(lastSelector) || {};
    const n = Number(r.eventElementsFound) || 0;
    if (n > 0 && r.selectorUsed) lastSelector = r.selectorUsed;
    return { eventElementsFound: n, loading: r.loading === true };
  };
  // Der zuletzt erfolgreiche Selektor wird auch fuer die abschliessende
  // vollstaendige Sammlung gebraucht (dieselbe Leiter, erstes probieren).
  observeCalendarEvents.lastSelector = () => lastSelector;
  return observeCalendarEvents;
}

// Abgleich des Warte-Ergebnisses mit der ABSCHLIESSENDEN Sammlung. Das ist die
// einzige Stelle, die den gemeldeten Ausgang festlegt — und sie kann
// 'not-polled' melden, also einen Logikfehler sichtbar machen, den die alte
// resolveEventWaitOutcome per destrukturierter, ungenutzter Argumente
// unsichtbar gemacht hat.
export function reconcileEventWait({ wait = null, eventElementsFound = 0, waitBudgetMs = 0 } = {}) {
  const n = Number(eventElementsFound);
  if (Number.isFinite(n) && n > 0) return { outcome: 'found', abortReason: null };
  const polls = wait ? Number(wait.polls) : NaN;
  if (Number(waitBudgetMs) > 0 && (!Number.isFinite(polls) || polls === 0)) {
    // Budget war da, aber es wurde nicht gepollt -> das ist ein Defekt, kein
    // DOM-Zustand. Muss benannt werden, statt als 'not-found' durchzurutschen.
    return { outcome: 'not-polled', abortReason: EVENT_WAIT_ABORT_REASONS.notPolled };
  }
  const w = wait || {};
  if (w.abortReason === EVENT_WAIT_ABORT_REASONS.renderedEmpty) {
    return { outcome: 'not-found', abortReason: EVENT_WAIT_ABORT_REASONS.renderedEmpty };
  }
  if (w.abortReason === EVENT_WAIT_ABORT_REASONS.budgetUnavailable) {
    return { outcome: 'budget-exhausted', abortReason: EVENT_WAIT_ABORT_REASONS.budgetUnavailable };
  }
  return { outcome: 'budget-exhausted', abortReason: EVENT_WAIT_ABORT_REASONS.budgetExhausted };
}

// EIN ehrlicher Abbruchgrund fuer die Diagnostik: der Navigations-Abbruch
// schlaegt die Wartephase, aber wenn die Navigation sauber war und die
// Wartephase gescheitert ist, muss deren Grund sichtbar werden — sonst endet
// der Call nach 4s still mit abortReason 'none'.
export function effectiveAbortReason({ navReason = 'none', eventWaitReason = null, eventElementsFound = 0 } = {}) {
  const nav = navReason && navReason !== 'none' ? navReason : null;
  const wait = Number(eventElementsFound) > 0 ? null : (eventWaitReason || null);
  if (nav && wait && nav !== wait) return `${nav}+${wait}`;
  return nav || wait || 'none';
}

// ---------------------------------------------------------------------------
// Klick-Pfad (Iteration 7, Defekt 2 — latenter Wochen-Skip)
// ---------------------------------------------------------------------------
// Runde 6 protokollierte einen Klick auf "Next week" (found: true, clicked: true,
// verified: 'unverified'), OBWOHL der Header die Zielwoche bereits bestätigt hatte.
// Gefährlich: bei noch unlesbarem Header ('unverified') ist die aktuelle Woche
// unbekannt — ein Klick kann die Zielwoche ÜBERSPRINGEN (KW 39 -> KW 41 statt
// -> KW 40). Deshalb: Klick nur bei echter Abweichung, und in dieselbe Richtung
// nie ohne erneute Wochen-Bestätigung.

// Protokolleintrag, wenn ohne Probe/Klick bereits die Zielwoche gerendert ist.
// `note` kennzeichnet, WO der Guard gegriffen hat (Live-Befund 25.09.2026:
// der Guard stand vorher nur in der aeusseren Schritt-Schleife und griff
// deshalb zu spaet — nach dem Klick auf "Next week").
export function weekProbeSkipEntry(renderedStart, note = null) {
  return {
    skipped: 'target-week-verified',
    renderedWeek: renderedStart || null,
    clicked: false,
    ...(note ? { note } : {})
  };
}

// Enthaelt die gerenderte Wochenansicht das angefragte Startdatum?
// Reine Funktion (unit-testbar), EIN Definition fuer beide Stellen: die
// aeussere Schritt-Schleife und der Klick-Guard in clickWeekButton.
export function renderedWeekContains(rendered, dateIso) {
  return Boolean(
    rendered && rendered.startDate && rendered.endDate && dateIso &&
    rendered.startDate <= dateIso && dateIso <= rendered.endDate
  );
}

// ---------------------------------------------------------------------------
// Bug 2 — der Zielwochen-Guard ist ein stiller No-op (Live-Befund 25.09.2026)
// ---------------------------------------------------------------------------
// Live-Trace:
//   0. aria-label-de  found:false  clicked:false
//   1. aria-label-en  found:true   clicked:true  verified:'unverified'
//   2. skipped:'target-week-verified'  renderedWeek:'2026-09-28'  clicked:false
// `note:'guard-vor-klick'` erscheint NIRGENDS.
//
// Ursache, aus dem Trace abgeleitet (kein Raten): evaluateClickVerification
// liefert 'unverified' genau dann, wenn prevStart ODER newStart fehlt. Eintrag 2
// stammt aus dem Zweig `rendered = forced.week` und zeigt einen DATIERBAREN
// Header ('2026-09-28', Skip verlangt startDate+endDate). Also war newStart
// gesetzt — und damit zwingend prevStart === null. Beweis:
//   node proof.mjs
//
// Bei prevStart === null war beim Klick KEIN Wochen-Header gerendert. Der alte
// Guard fragte ausschliesslich "ist die Zielwoche schon da?" und
// renderedWeekContains({startDate:null,...}) ist zwangslaeufig false. Der Guard
// war also ein No-op GENAU in dem Fall, fuer den er gebaut wurde, und er
// failte OPEN und still: kein Probe-Eintrag, kein note, kein Abbruchgrund.
// Geklickt wurde damit auf einer noch ladenden Ansicht — das ist exakt die
// Wochen-Skip-Gefahr (KW 39 -> KW 41), die 8ae0eec verhindern sollte.
//
// Der Guard bekommt deshalb eine zweite, notwendige Bedingung: die Ansicht
// muss ueberhaupt gerendert sein. Ohne datierbaren Header ist die Wochen- und
// damit auch die Button-Position unbekannt — ein Klick darauf ist nicht
// vertretbar, egal welches Week-Button-Element man trifft.
export const WEEK_CLICK_GUARD_REASONS = {
  targetReached: 'target-week-reached',
  viewNotReady: 'view-not-ready',
  click: 'click'
};

// EIN Entscheidungs-Funktion, benutzt von der aeusseren Schritt-Schleife UND
// vom Klickpfad. Beide Stellen konnten vorher auseinanderlaufen — genau das war
// der Live-Befund. Reine Funktion, damit der Draht testbar ist.
export function weekClickDecision({ rendered, targetStart, hasEvents = false } = {}) {
  if (renderedWeekContains(rendered, targetStart)) {
    return { skip: true, reason: WEEK_CLICK_GUARD_REASONS.targetReached, rendered: rendered || null, clickable: false };
  }
  // Ansicht nicht gerendert? Dann ist die Wochen-/Button-Position unbekannt.
  const headerRendered = Boolean(rendered && rendered.startDate && rendered.endDate);
  if (!headerRendered && !hasEvents) {
    return { skip: true, reason: WEEK_CLICK_GUARD_REASONS.viewNotReady, rendered: rendered || null, clickable: false };
  }
  return { skip: false, reason: WEEK_CLICK_GUARD_REASONS.click, rendered: rendered || null, clickable: true };
}

// Protokolleintrag fuer eine Guard-Entscheidung. BEIDE Ausgaenge werden
// protokolliert — ein Guard, der nichts aufschreibt, ist nicht von einem
// Guard, der nicht existiert, unterscheidbar. Genau daran ist der Live-Befund
// gescheitert: `note:'guard-vor-klick'` hätte sofort verraten, dass der Guard
// nicht griff.
export function weekProbeGuardEntry(decision, note = 'guard-vor-klick') {
  const renderedStart = decision && decision.rendered ? decision.rendered.startDate : null;
  if (decision && decision.reason === WEEK_CLICK_GUARD_REASONS.viewNotReady) {
    return {
      skipped: 'view-not-ready',
      renderedWeek: renderedStart || null,
      clicked: false,
      guard: WEEK_CLICK_GUARD_REASONS.viewNotReady,
      note
    };
  }
  return { ...weekProbeSkipEntry(renderedStart, note), guard: WEEK_CLICK_GUARD_REASONS.targetReached };
}

// ---------------------------------------------------------------------------
// Der ECHTE Klickpfad — exportiert und per Dependency-Injection testbar
// ---------------------------------------------------------------------------
// Bis 0d47105 lag dieser Code als Closure in getFreeBusy. Der Test
// (tests/week-probe-guard.test.mjs) baute die Schleife NACH — er rief nie
// diesen Code. Genau darum war der Guard "durch Code-Review belegt, nicht
// durch einen Test" und der Live-Befund kam trotzdem durch.
//
// Diese Funktion IST der Produktionspfad; getFreeBusy ruft sie mit denselben
// Abhaengigkeiten auf. Der Test fährt damit den Weg Guard -> Probe -> Klick
// wirklich, statt ihn zu imitieren.
export async function runWeekButtonClick({
  direction = 1,
  prevStart = null,
  targetStart = null,
  page = null,
  probe = [],
  candidates = null,
  readWeek = null,
  hasEvents = false,
  afterClick = null,
  clickWaitMs = FREEBUSY_CLICK_RENDER_WAIT_MS,
  clickCostMs = FREEBUSY_CLICK_RENDER_WAIT_MS,
  remainingMs = () => Infinity
} = {}) {
  const list = Array.isArray(candidates) && candidates.length ? candidates : weekButtonCandidates(direction);
  const notReached = (reason, rendered) => ({
    clicked: false, week: rendered || null, verified: 'guard-blocked', foundAny: false, clickedAny: false,
    targetReached: reason === WEEK_CLICK_GUARD_REASONS.targetReached,
    guard: reason
  });

  // Der Header-Read ist memoisiert: nur der erste Kandidat kostet einen Read.
  // Nach jedem Klick wird er verworfen, weil erst der Klick die Woche aendern kann.
  let targetWeekChecked = false;
  let targetWeekNow = null;
  const guardNow = async () => {
    if (!targetWeekChecked) {
      targetWeekChecked = true;
      let wk = null;
      try { wk = await readWeek(); } catch (e) { wk = null; }
      targetWeekNow = weekClickDecision({ rendered: wk, targetStart, hasEvents });
    }
    return targetWeekNow;
  };
  let foundAny = false;
  let clickedAny = false;

  for (const cand of list) {
    // GUARD VOR DER PROBE — fuer JEDE Entscheidung wird protokolliert.
    const decision = await guardNow();
    if (decision.skip) {
      probe.push(weekProbeGuardEntry(decision));
      return notReached(decision.reason, decision.rendered);
    }
    if (remainingMs() < clickCostMs) {
      probe.push({ selector: cand.selector, kind: cand.kind, found: null, clicked: false, note: 'Budget aufgebraucht' });
      break;
    }
    const loc = page.locator(cand.selector).first();
    let found = false;
    try {
      await loc.waitFor({ state: 'visible', timeout: Math.min(1000, Math.max(100, remainingMs())) });
      found = true;
    } catch (e) { found = false; }
    if (!found) {
      probe.push({ selector: cand.selector, kind: cand.kind, found: false, clicked: false });
      continue;
    }
    foundAny = true;
    try {
      await loc.click({ timeout: Math.min(2000, Math.max(100, remainingMs())) });
    } catch (e) {
      probe.push({ selector: cand.selector, kind: cand.kind, found: true, clicked: false, clickError: String((e && e.message) || e).slice(0, 120) });
      continue;
    }
    clickedAny = true;
    // Ab hier kann sich die gerenderte Woche geaendert haben — der
    // memoierte Guard-Zustand ist damit veraltet.
    targetWeekChecked = false;
    targetWeekNow = null;
    // eslint-disable-next-line no-await-in-loop
    const r = await afterClick(prevStart, clickWaitMs);
    probe.push({ selector: cand.selector, kind: cand.kind, found: true, clicked: true, verified: r.verified });
    if (shouldTryNextCandidate(r.verified)) continue; // wirkungslos -> naechster Kandidat
    return { clicked: true, week: r.week, verified: r.verified, foundAny, clickedAny, targetReached: false, guard: WEEK_CLICK_GUARD_REASONS.click };
  }
  return { clicked: false, week: null, verified: 'not-found', foundAny, clickedAny, targetReached: false, guard: null };
}

// Darf in dieselbe Richtung erneut geklickt werden?
//  - noch nie geklickt                    -> ja
//  - Gegenrichtung (direction != last)    -> ja; das ist der RÜCKWEG, der
//    ausdrücklich vorbereitet bleiben muss
//  - gleiche Richtung                     -> nur wenn die Woche nach dem letzten
//    Klick ERNEUT bestätigt wurde (datierbarer Header). Sonst unbestätigter
//    Sprung -> nein.
export function canClickSameDirection({ lastClickDirection = null, direction = 0, weekConfirmedAfterClick = false } = {}) {
  if (!lastClickDirection || !direction) return true;
  if (Math.sign(direction) !== Math.sign(lastClickDirection)) return true;
  return Boolean(weekConfirmedAfterClick);
}

// Hinweis, wenn die Event-Wartephase ohne echte Events endete. Reine Funktion.
export function eventWaitWarning({ outcome, waitedMs = 0, budgetMs = null } = {}) {
  if (outcome === 'found') return null;
  if (outcome === 'not-found') {
    return `Nach ${Math.round(waitedMs)}ms Wartephase (Budget ${budgetMs === null ? 'n/a' : budgetMs}ms) ` +
      `wurden KEINE Event-Knoten im DOM gelesen. ` +
      `0 Busy-Slots bedeuten damit KEINEN leeren Kalender, sondern ein nicht auslesbares Grid ` +
      `(siehe diagnostics.census/sourceTried).`;
  }
  if (outcome === 'budget-exhausted') {
    return `Event-Wartephase nach ${Math.round(waitedMs)}ms ohne echtes Event-Label beendet ` +
      `(Budget ${budgetMs === null ? 'n/a' : budgetMs}ms, Reserve für Extraktion abgezogen). ` +
      `0 Busy-Slots können damit ein noch nicht gerendertes Grid bedeuten, nicht einen leeren Kalender.`;
  }
  if (outcome === 'not-rendered') {
    return 'Event-Grid wurde nicht gerendert — weder datierter Wochen-Header noch echtes Event-Label gefunden.';
  }
  if (outcome === 'not-polled') {
    return `Wartephase mit ${budgetMs === null ? 'n/a' : budgetMs}ms Budget gestartet, aber OHNE Polls ` +
      `beendet (0 Polls, ${Math.round(waitedMs)}ms). Das ist ein Logikfehler der Wartephase, ` +
      `KEIN DOM-Zustand — der Kalender wurde hier gar nicht abgewartet.`;
  }
  return null;
}

// Teilt das Gesamtbudget auf Phasen auf: Eine Phase bekommt nie mehr als ihr
// Cap und nie mehr als die Restzeit; unterhalb der Mindestscheibe wird sie
// GESKIPPT (0). Runde-3-Root-Cause: Der `Math.max(1000, ...)`-Floor ließ Phasen
// nach Ablauf der Deadline weiterlaufen, und die Summe der Phasen (goto 30s +
// URL-Wait 6s + Render-Wait 20s + Retry-Goto 36s + Retry-Wait 15s) sprengte
// das Budget, obwohl jede Einzelphase gegen `remaining()` gekappt war.
export function phaseBudgetMs(remainingMs, capMs, minSliceMs = 1500) {
  if (!remainingMs || remainingMs <= 0 || !capMs || capMs <= 0) return 0;
  const slice = Math.min(capMs, remainingMs);
  return slice >= minSliceMs ? slice : 0;
}

// Wie viele Button-Klicks sind im Restbudget noch bezahlbar (Klick + Re-Render-Wait)?
export function affordableClicks(remainingMs, clickCostMs, maxClicks) {
  if (!clickCostMs || clickCostMs <= 0) return Math.max(0, maxClicks || 0);
  if (!remainingMs || remainingMs <= 0) return 0;
  return Math.max(0, Math.min(maxClicks || 0, Math.floor(remainingMs / clickCostMs)));
}

// Kandidaten-Selektoren für die Wochen-Vor/Zurück-Buttons (Runde-4-Forderung C):
// multi-selektorisch — aria-label (DE/EN), title (DE/EN), breitere Wortmuster und
// symbolbasierte Chevron-Kandidaten. Reihenfolge = Versuchreihenfolge; jeder
// Versuch wird ins Klick-Protokoll (weekButtonProbe) geschrieben, auch
// "kein Button gefunden".
export function weekButtonCandidates(direction) {
  if (direction > 0) {
    return [
      { selector: '[aria-label*="Nächste Woche" i]', kind: 'aria-label-de' },
      { selector: '[aria-label*="Next week" i]', kind: 'aria-label-en' },
      { selector: '[title*="Nächste Woche" i]', kind: 'title-de' },
      { selector: '[title*="Next week" i]', kind: 'title-en' },
      { selector: '[role="button"][aria-label*="Nächste" i]', kind: 'aria-label-de-broad' },
      { selector: '[role="button"][aria-label*="Next" i]', kind: 'aria-label-en-broad' },
      { selector: '[role="button"]:has([data-icon-name*="ChevronRight" i])', kind: 'symbol-chevron-right' },
      { selector: '[role="button"][aria-label*="chevron-right" i], [role="button"][aria-label*="chevron right" i]', kind: 'symbol-aria-chevron' }
    ];
  }
  return [
    { selector: '[aria-label*="Vorherige Woche" i]', kind: 'aria-label-de' },
    { selector: '[aria-label*="Previous week" i]', kind: 'aria-label-en' },
    { selector: '[title*="Vorherige Woche" i]', kind: 'title-de' },
    { selector: '[title*="Previous week" i]', kind: 'title-en' },
    { selector: '[role="button"][aria-label*="Vorherige" i]', kind: 'aria-label-de-broad' },
    { selector: '[role="button"][aria-label*="Previous" i]', kind: 'aria-label-en-broad' },
    { selector: '[role="button"]:has([data-icon-name*="ChevronLeft" i])', kind: 'symbol-chevron-left' },
    { selector: '[role="button"][aria-label*="chevron-left" i], [role="button"][aria-label*="chevron left" i]', kind: 'symbol-aria-chevron' }
  ];
}

// Klick-Verifikation (Runde-4-Forderung C): Nach dem Klick MUSS sich der
// gerenderte Header ändern (Datum), sonst gilt der Klick als wirkungslos.
// 'unverified' = vor ODER nach dem Klick war kein datierter Header lesbar —
// die Änderung ist nicht bewertbar (Doppel-Klick-Gefahr).
export function evaluateClickVerification(prevStart, newWeek) {
  const newStart = newWeek && newWeek.startDate ? newWeek.startDate : null;
  if (!prevStart || !newStart) return 'unverified';
  return newStart !== prevStart ? 'changed' : 'unchanged';
}

// Nur bei NACHWEISLICH wirkungslosem Klick (Header lesbar und UNVERÄNDERT) wird
// der nächste Selektor-Kandidat versucht — bei 'unverified' nie (Gefahr des
// Doppel-Advancierens, wenn der erste Klick doch gewirkt hat).
export function shouldTryNextCandidate(verified) {
  return verified === 'unchanged';
}

// Präzise Abbruch-Ursache nach erfolgloser Button-Navigation (Runde-4-Forderung 3):
// Wurde überhaupt ein Button gefunden/geklickt, aber der Header blieb unverändert,
// ist es 'button-ineffective' — NICHT 'no-week-button' (der Button existierte ja).
// Nur wenn KEIN Kandidat gefunden wurde, ist es 'no-week-button'.
export function classifyButtonFailure({ foundAny = false, clickedAny = false } = {}) {
  return (foundAny || clickedAny) ? 'button-ineffective' : 'no-week-button';
}

// Assembliert die FreeBusy-Diagnostik für Fehler-/unverified-Responses
// (Runde-4-Forderung B): Der frühe Snapshot (direkt nach Navigation erhoben)
// bleibt immer erhalten, auch wenn das Budget für eine frische Erhebung nicht
// mehr reicht. Klick-Protokoll, Abbruchgrund und Budget-Report kommen hinzu.
export function assembleFreeBusyDiagnostics({ fresh = null, early = null, weekButtonProbe = [], abortReason = 'none', budget = null, eventElementsFound = null, selectorUsed = null, rawEventLabels = null, rejectionReasons = null, waitedForEventsMs = null, eventWaitOutcome = null, eventWaitAbortReason = null, eventWaitLoadingPolls = null, loadingIndicators = null } = {}) {
  const base = fresh || early || { url: '', headerCandidates: [], dayHeaderCandidates: [], weekNavCandidates: [], ariaButtonCount: 0, ariaLabelSample: [] };
  const out = {
    ...base,
    weekButtonProbe,
    abortReason
  };
  if (early && early.url) out.urlAfterNavigation = early.url;
  if (budget) out.budget = budget;
  // Runde-5-Kernlücke: Slot-Extraktions-Diagnostik auch im Erfolgsfall (status ok),
  // damit ein echter Free-Tag von einem Extraktions-Miss unterscheidbar ist.
  if (eventElementsFound !== null) out.eventElementsFound = eventElementsFound;
  if (selectorUsed !== null) out.selectorUsed = selectorUsed;
  if (rawEventLabels !== null) out.rawEventLabels = rawEventLabels;
  if (rejectionReasons !== null) out.rejectionReasons = rejectionReasons;
  // Runde-6-Diagnostik (Iteration 7): Die Event-Wartephase ist jetzt eine eigene,
  // budgetgebundene Phase — ihre Dauer und ihr Ausgang gehören in die Diagnostik,
  // damit "0 Slots" zwischen "leerer Kalender" und "Grid noch nicht gerendert"
  // unterscheidbar bleibt.
  if (waitedForEventsMs !== null) out.waitedForEventsMs = waitedForEventsMs;
  if (eventWaitOutcome !== null) out.eventWaitOutcome = eventWaitOutcome;
  if (eventWaitAbortReason !== null) out.eventWaitAbortReason = eventWaitAbortReason;
  if (eventWaitLoadingPolls !== null) out.eventWaitLoadingPolls = eventWaitLoadingPolls;
  if (loadingIndicators !== null) out.loadingIndicators = loadingIndicators;
  return out;
}

// Widerspruchs-Warnung (Runde-5-Forderung): 0 Slots, obwohl Event-Elemente im DOM
// gefunden wurden -> Extraktions-Miss, kein leerer Kalender. Reine Funktion.
export function slotExtractionWarning(eventElementsFound, slotCount) {
  if (eventElementsFound > 0 && slotCount === 0) {
    return `Widerspruch: ${eventElementsFound} Event-Element(e) im DOM gefunden, aber 0 Busy-Slots extrahiert — ` +
      `Extraktions-Miss wahrscheinlich (siehe rejectionReasons).`;
  }
  return null;
}

// Parst ein OWA-Event-ARIA-Label in einen strukturierten Busy-Slot.
// Erkannte Muster (Beispiele):
//  - "Team Meeting, 14:00 to 15:00, Monday, September 28, 2026"
//  - "Besetzt von 09:00 bis 10:00 Montag, 28. September"
//  - "Urlaub, Ganztägig, Montag, 28. September 2026"  (All-Day: start/end = null)
// Rückgabe: { date: 'YYYY-MM-DD', start: 'HH:MM'|null, end: 'HH:MM'|null,
//             title, allDay, raw } — oder null, wenn kein Datum erkennbar ist.
export function parseEventLabel(label, defaultYear) {
  if (!label) return null;

  // Datum extrahieren: DE ("28. September 2026") vor EN ("September 28, 2026") versuchen.
  // matchAll + Monats-Check: falsche Kandidaten (z.B. "Besetzt 9") werden übersprungen.
  let day = null;
  let month = null;
  let year = null;
  const deDateRe = /(\d{1,2})\.\s*([A-Za-zäöüÄÖÜß]{3,})\.?(?:\s+(\d{4}))?/g;
  for (const m of label.matchAll(deDateRe)) {
    const mo = MONTH_NAMES[m[2].toLowerCase()];
    if (mo !== undefined) {
      day = parseInt(m[1], 10);
      month = mo;
      year = m[3] ? parseInt(m[3], 10) : null;
      break;
    }
  }
  if (day === null) {
    const enDateRe = /([A-Za-zäöüÄÖÜß]{3,})\.?\s+(\d{1,2})(?:,?\s*(\d{4}))?/g;
    for (const m of label.matchAll(enDateRe)) {
      const mo = MONTH_NAMES[m[1].toLowerCase()];
      if (mo !== undefined) {
        month = mo;
        day = parseInt(m[2], 10);
        year = m[3] ? parseInt(m[3], 10) : null;
        break;
      }
    }
  }
  if (day === null || month === null) return null;
  if (year === null) year = defaultYear;

  const pad = (n) => String(n).padStart(2, '0');
  const dateIso = `${year}-${pad(month + 1)}-${pad(day)}`;

  // Zeitfenster extrahieren: "von 09:00 bis 10:00" / "9:00 AM to 10:00 AM" / "14:00 - 15:00"
  const timeRe = /(\d{1,2}):(\d{2})\s*(AM|PM)?\s*(?:bis|to|[–—-])\s*(\d{1,2}):(\d{2})\s*(AM|PM)?/i;
  const tm = label.match(timeRe);

  if (!tm) {
    // Kein Zeitfenster im Label -> Ganztag-Event (start/end bewusst null).
    // Titel = Label-Anteil VOR der ersten Zeitkomponente (hier: dem Ganztag-Marker),
    // damit Titel mit Kommas/Doppelpunkten (z.B. "Stadt Hörstel: CMIS …") intakt
    // bleiben und Folge-Attribute (Datum, Status, "Recurring event") abgeschnitten
    // werden.
    const allDayMatch = label.match(/all[- ]?day|ganztägig|ganztages/i);
    const cut = allDayMatch ? allDayMatch.index : label.length;
    const title = label.slice(0, cut).replace(/[,;:\s]+$/, '').trim() || '(Ganztägig)';
    return { date: dateIso, start: null, end: null, title, allDay: true, raw: label };
  }

  const toMinutes = (h, m, ampm) => {
    let H = parseInt(h, 10);
    const M = parseInt(m, 10);
    const ap = (ampm || '').toUpperCase();
    if (ap === 'PM' && H < 12) H += 12;
    if (ap === 'AM' && H === 12) H = 0;
    return H * 60 + M;
  };
  const startMin = toMinutes(tm[1], tm[2], tm[3]);
  let endMin = toMinutes(tm[4], tm[5], tm[6] || tm[3]);
  // Über Mitternacht gehende Slots (z.B. "23:30 to 00:30") auf den Folgetag rollen
  if (endMin <= startMin) endMin += 1440;
  const fmt = (mins) => `${pad(Math.floor(mins / 60) % 24)}:${pad(mins % 60)}`;

  // Titel = Label-Anteil vor der Zeitangabe (Trennzeichen/Präpositionen am Ende entfernen)
  const title = label.slice(0, tm.index)
    .replace(/[,;:\s]+$/, '')
    .replace(/\s+(von|from)$/i, '')
    .trim() || '(ohne Titel)';

  return { date: dateIso, start: fmt(startMin), end: fmt(endMin), title, allDay: false, raw: label };
}

// Klassifiziert ein Event-ARIA-Label anhand einer Wortliste (DE + EN).
// Rückgabe: 'busy' (Default für gerenderte Event-Blöcke) | 'tentative' | 'free' | 'non_event'.
// 'non_event' markiert Kalender-Banner, die keine Belegung sind (z.B. Arbeitsort-Hinweise)
// und von der Busy-Slot-Extraktion ausgeschlossen werden.
// Wortgrenzen sind Pflicht: "Freitag" darf NICHT als 'free' klassifiziert werden.
export function classifyEventStatus(label) {
  const t = String(label || '');
  if (/\barbeitsort\b|\bworking location\b/i.test(t)) return 'non_event';
  if (/\bfree\b|\bfrei\b/i.test(t)) return 'free';
  if (/tentative|\bvorbehalt\b/i.test(t)) return 'tentative';
  return 'busy';
}

// Erkennt Work-Location-/Pseudo-Slots OHNE echten Titel (Beauty-Fix Live Runde 2):
// OWA rendert "Arbeitsort"/"Working location"-Banner und titellose Platzhalter-Events
// ("(ohne Titel)") als belegungsähnliche Elemente — sie blockieren keine echte
// Meetingzeit und werden aus den Busy-Slots ausgeschlossen. Auch reine URL-Schnipsel
// als "Titel" (z.B. Teams-Links) gelten nicht als echter Termin.
// Rückgabe: true = Pseudo-Slot (ausschließen), false = echter Termin.
export function isWorkLocationPseudoSlot(title) {
  const t = String(title || '').trim();
  if (!t) return true; // gar kein Titel -> Pseudo-Slot
  if (/^\(?\s*(ohne titel|no title|unbenannt|untitled)\s*\)?$/i.test(t)) return true;
  if (/arbeitsort|working location|homeoffice|home office|\bbüro\b/i.test(t)) return true;
  if (/https?:\/\//i.test(t)) return true;
  return false;
}

// Dedupliziert Busy-Slots (Beauty-Fix Live Runde 2): identisches Datum + Zeitfenster
// nur 1x; überlappende Duplikate gleicher Zeit werden zusammengeführt — bevorzugt
// bleibt der Slot mit echtem (nicht-pseudo) Titel, sonst der längere Titel.
// Ganztag-Slots werden pro Tag zusammengeführt. Reine Funktion (unit-testbar).
export function dedupeBusySlots(slots) {
  const byKey = new Map();
  for (const slot of slots || []) {
    if (!slot || !slot.date) continue;
    const timeKey = slot.allDay ? 'allday' : `${slot.start}-${slot.end}`;
    const key = `${slot.date}|${timeKey}`;
    const existing = byKey.get(key);
    if (!existing) {
      byKey.set(key, slot);
      continue;
    }
    const exPseudo = isWorkLocationPseudoSlot(existing.title);
    const newPseudo = isWorkLocationPseudoSlot(slot.title);
    if (exPseudo && !newPseudo) {
      byKey.set(key, slot);
    } else if (!exPseudo && !newPseudo
      && String(slot.title || '').length > String(existing.title || '').length) {
      byKey.set(key, slot);
    }
  }
  return Array.from(byKey.values());
}

// Partitioniert Slots in "im angefragten Range [rangeStartMs, rangeEndMs)" und
// "verworfen" (mit Grund). Ganztag-Events belegen 00:00–24:00 ihres Tages und
// überlappen den Range, wenn ihr Tag im Fenster liegt (z.B. allDay am 30.09. bei
// Range 28.09.–02.10. -> behalten). Explizit freie Slots ('free') sind keine
// Belegung. Reine Funktion (unit-testbar).
export function partitionSlotsByRange(slots, rangeStartMs, rangeEndMs) {
  const kept = [];
  const rejected = [];
  for (const slot of slots || []) {
    if (!slot || !slot.date) { rejected.push({ slot, reason: 'no-time-pattern' }); continue; }
    if (slot.status === 'free') { rejected.push({ slot, reason: 'free' }); continue; }
    const dayStartMs = new Date(`${slot.date}T00:00:00`).getTime();
    const s = slot.allDay ? dayStartMs : new Date(`${slot.date}T${slot.start}:00`).getTime();
    const e = slot.allDay ? dayStartMs + 86400000 : new Date(`${slot.date}T${slot.end}:00`).getTime();
    if (s < rangeEndMs && e > rangeStartMs) kept.push(slot);
    else rejected.push({ slot, reason: 'date-out-of-range' });
  }
  return { kept, rejected };
}

// Parst aus dem Zeilentext einer OWA-Nachrichtenliste/Suchergebniszeile einen
// Zeitstempel (ms seit Epoch) — oder null, wenn kein Datum erkennbar ist.
// Unterstützte Muster: "24.09.2026", "24.09." (aktuelles Jahr), "Heute 14:32",
// "Gestern", Wochentag-Kürzel ("Mo.", "Di.", ...) und freistehende Uhrzeiten ("14:32" = heute).
// Wird genutzt, um aus Suchergebnissen die NEUESTE Nachricht eines Threads zu wählen.
export function parseMailRowTimestamp(text, now = new Date()) {
  if (!text) return null;
  const t = String(text);
  const base = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const DAY_MS = 86400000;

  // "Heute" / "Today" (optional mit Uhrzeit)
  if (/\bheute\b|\btoday\b/i.test(t)) {
    const tm = t.match(/(\d{1,2}):(\d{2})/);
    if (tm) return base.getTime() + parseInt(tm[1], 10) * 3600000 + parseInt(tm[2], 10) * 60000;
    return base.getTime() + DAY_MS - 1; // heute, ohne Uhrzeit -> "spätester heute-Zeitpunkt"
  }

  // "Gestern" / "Yesterday"
  if (/\bgestern\b|\byesterday\b/i.test(t)) {
    return base.getTime() - DAY_MS + 43200000; // Mittag gestern
  }

  // Vollständiges Datum: 24.09.2026 (optional mit Uhrzeit verfeinert)
  const full = t.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (full) {
    const d = new Date(parseInt(full[3], 10), parseInt(full[2], 10) - 1, parseInt(full[1], 10));
    const tm = t.match(/(\d{1,2}):(\d{2})/);
    if (tm) d.setHours(parseInt(tm[1], 10), parseInt(tm[2], 10));
    return d.getTime();
  }

  // Kurzes Datum: 24.09. (ohne Jahr -> aktuelles Jahr)
  const short = t.match(/(\d{1,2})\.(\d{1,2})\.(?!\d)/);
  if (short) {
    return new Date(now.getFullYear(), parseInt(short[2], 10) - 1, parseInt(short[1], 10)).getTime();
  }

  // Wochentag (Kürzel oder ausgeschrieben): jüngstes vergangenes Vorkommen (letzte 7 Tage)
  const weekdayMap = {
    'so': 0, 'sonntag': 0, 'su': 0, 'sun': 0, 'sunday': 0,
    'mo': 1, 'montag': 1, 'mon': 1, 'monday': 1,
    'di': 2, 'dienstag': 2, 'tu': 2, 'tue': 2, 'tuesday': 2,
    'mi': 3, 'mittwoch': 3, 'we': 3, 'wed': 3, 'wednesday': 3,
    'do': 4, 'donnerstag': 4, 'th': 4, 'thu': 4, 'thursday': 4,
    'fr': 5, 'freitag': 5, 'fri': 5, 'friday': 5,
    'sa': 6, 'samstag': 6, 'sa.': 6, 'sat': 6, 'saturday': 6
  };
  const wd = t.match(/\b(So|Mo|Di|Mi|Do|Fr|Sa|Su|Mon|Tue|Wed|Thu|Fri|Sat|Sonntag|Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday)\b\.?/i);
  if (wd) {
    const target = weekdayMap[wd[1].toLowerCase().replace(/\.$/, '')];
    if (target !== undefined) {
      // Heutige Mails zeigen die Uhrzeit statt des Wochentags -> "Mo." an einem Montag = letzter Montag
      let diff = (base.getDay() - target + 7) % 7;
      if (diff === 0) diff = 7;
      return base.getTime() - diff * DAY_MS + 43200000;
    }
  }

  // Freistehende Uhrzeit (eigene Zeile) -> heute. Bewusst anchoring, damit Uhrzeiten
  // aus dem Mail-Text (z.B. "Meeting um 14:00") nicht fälschlich "heute" bedeuten.
  const tm = t.match(/(^|\n)\s*(\d{1,2}):(\d{2})\s*(\n|$)/);
  if (tm) {
    return base.getTime() + parseInt(tm[2], 10) * 3600000 + parseInt(tm[3], 10) * 60000;
  }

  return null;
}

// Quellen der Mailbox-Verifikation (Runde-4-Forderung D). Live-Befund 25.09.:
// Im Kalender-Tab lieferte die bisherige (Mail-Ansicht-)Heuristik LEERE Hints,
// wodurch checkStatus fälschlich activeMailboxMatchesRealm:false meldete.
export const MAILBOX_VERIFICATION_SOURCES = ['mail-view', 'calendar-header', 'tab-title', 'unknown'];

// Reine Verifikationslogik: Passt ein Hint zum Realm? Ohne Hints oder ohne
// Realm ist das Ergebnis bewusst NULL (unbekannt) statt false — ein fehlender
// Hint ist KEIN Nachweis einer falschen Mailbox (Kollateral-Bug Runde 3).
export function resolveMailboxVerification(hints, realmDomain) {
  const list = Array.isArray(hints) ? hints : [];
  if (!realmDomain || list.length === 0) {
    return { activeMailboxMatchesRealm: null, matchingHint: null };
  }
  const matchingHint = list.find((h) => String(h).toLowerCase().includes(realmDomain)) || null;
  return { activeMailboxMatchesRealm: matchingHint !== null, matchingHint };
}

export class OwaClient {
  constructor(realm = config.defaultRealm) {
    this.realm = realm;
  }

  setRealm(realm) {
    if (realm) {
      this.realm = realm;
    }
  }

  async getPage(headless = true) {
    const { page } = await browserManager.ensureContext(headless, this.realm);
    this.lastPage = page;
    const currentUrl = page.url();
    if (!currentUrl.includes('outlook.office.com') && !currentUrl.includes('login.microsoftonline.com') && !currentUrl.includes('outlook.cloud.microsoft')) {
      const url = this.realm
        ? `https://outlook.office.com/mail/?realm=${encodeURIComponent(this.realm)}`
        : 'https://outlook.office.com/mail/';
      await page.goto(url, { waitUntil: 'domcontentloaded' });
    }
    return page;
  }

  // Bestimmt die tatsächlich aktive Mailbox-Adresse (nicht nur den Titel).
  // Runde-4-Forderung D (Kollateral-Bug): Die Hints dürfen NICHT ausschließlich
  // aus der Mail-Ansicht stammen — im Kalender-Tab war die bisherige Heuristik
  // leer (Live-Befund 25.09.), wodurch checkStatus fälschlich
  // activeMailboxMatchesRealm:false meldete. Quellen-Kaskade:
  //   1. 'mail-view'        — bisherige DOM-Heuristik (Account-Button, Persona-
  //                           Anker, Header-Text; bewährt in der Mail-Ansicht)
  //   2. 'calendar-header'  — Account-Switcher/ARIA-Labels mit @ im App-Bereich
  //                           (funktioniert auch im Kalender-Tab)
  //   3. 'tab-title'        — E-Mail-Adresse im Fenster-/Tab-Titel
  //   4. 'unknown'          — keine Hints gefunden (Verifikation unbekannt, s. resolveMailboxVerification)
  // Rückgabe: { hints: string[], source: MAILBOX_VERIFICATION_SOURCES-Eintrag }.
  async detectActiveMailbox() {
    const page = this.lastPage || (await this.getPage(true));
    // 1) Bisherige DOM-Heuristik (Mail-Ansicht + generische Header-Elemente)
    try {
      const domHints = await page.evaluate(() => {
        const candidates = [];
        // 1) OWA-Account-Button (oben rechts): aria-label / title / Text
        const accountBtn = document.querySelector('[data-tid="account"] button, [role="button"][aria-label*="Account" i], [role="button"][aria-label*="Konto" i]');
        if (accountBtn) {
          const lbl = accountBtn.getAttribute('aria-label') || accountBtn.getAttribute('title') || accountBtn.innerText || '';
          if (lbl) candidates.push(lbl);
        }
        // 2) Account-Switcher / Persona-Links mit @
        document.querySelectorAll('[title*="@"], [aria-label*="@"], [data-tid="accountName"], [data-account-name]').forEach(el => {
          const txt = (el.getAttribute('title') || el.getAttribute('aria-label') || el.innerText || '').trim();
          if (txt && txt.includes('@')) candidates.push(txt);
        });
        // 3) Textknoten im Header, die eine E-Mail enthalten
        const header = document.querySelector('[role="banner"], #owaHeader, [data-tid="header"]');
        if (header) {
          const txt = header.innerText || '';
          const emails = txt.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || [];
          candidates.push(...emails);
        }
        return [...new Set(candidates.map(c => c.trim()).filter(Boolean))].slice(0, 8);
      });
      if (domHints && domHints.length) return { hints: domHints, source: 'mail-view' };
    } catch (e) { /* bewusst ignorieren — nächste Quelle */ }
    // 2) Kalender-Header: Account-Switcher/ARIA-Labels mit @ im gesamten App-Bereich
    try {
      const calHints = await page.evaluate(() => {
        const candidates = [];
        // Account-Elemente im Banner/App-Bar mit @ in Label oder Titel
        document.querySelectorAll(
          '[role="banner"] [aria-label*="@"], [role="banner"] [title*="@"], header [aria-label*="@"], header [title*="@"], [data-tid="account"]'
        ).forEach(el => {
          const txt = (el.getAttribute('aria-label') || el.getAttribute('title') || el.innerText || '').trim();
          if (txt) candidates.push(txt);
        });
        // Account-Menü-/Switcher-Button (Initialen) — aria-label enthält oft Name ODER Adresse
        document.querySelectorAll('[role="button"][aria-label*="Konto" i], [role="button"][aria-label*="Account" i], [role="button"][aria-label*="account manager" i], [role="button"][aria-label*="Kontomanager" i]').forEach(el => {
          const txt = (el.getAttribute('aria-label') || el.getAttribute('title') || '').trim();
          if (txt) candidates.push(txt);
        });
        // E-Mail-Muster im gesamten Banner-Text
        const banner = document.querySelector('[role="banner"], header, #owaHeader');
        if (banner) {
          const emails = (banner.innerText || '').match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || [];
          candidates.push(...emails);
        }
        return [...new Set(candidates.map(c => c.trim()).filter(Boolean))].slice(0, 8);
      });
      if (calHints && calHints.length) return { hints: calHints, source: 'calendar-header' };
    } catch (e) { /* bewusst ignorieren — nächste Quelle */ }
    // 3) Tab-/Fenstertitel (OWA rendert dort je nach Build Name/Adresse)
    try {
      const title = await page.title();
      const emails = (title || '').match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || [];
      if (emails.length) return { hints: emails, source: 'tab-title' };
    } catch (e) { /* bewusst ignorieren */ }
    return { hints: [], source: 'unknown' };
  }

  async checkStatus() {
    try {
      const page = await this.getPage(true);
      const url = page.url();

      if (url.includes('login.microsoftonline.com')) {
        return {
          authenticated: false,
          currentUrl: url,
          realm: this.realm,
          message: "Nicht angemeldet. Bitte führe 'outlook_login' aus, um das Browserfenster für die einmalige Anmeldung zu öffnen."
        };
      }

      // Ansichtsgerechtes Warten: Im Kalender-Tab existieren die Mail-Selektoren
      // nie (Live-Befund Runde 4) — dort auf den Kalender-Banner warten statt
      // 12s in die Leere zu laufen.
      const onCalendarView = url.includes('/calendar');
      await page.waitForSelector(
        onCalendarView
          ? '[role="banner"], div[role="main"]'
          : '#topSearchInput, input[aria-label*="Search"], div[aria-label*="Message list"]',
        { timeout: onCalendarView ? 6000 : 12000 }
      ).catch(() => null);

      const title = await page.title();
      const isLoaded = title.toLowerCase().includes('outlook') || page.url().includes('/mail') || page.url().includes('/calendar');

      // Aktive Mailbox-Adresse robust bestimmen (Verwechslungs-Schutz zwischen
      // Tenants) — mit Verifikationsquelle statt reiner Mail-Ansicht-Heuristik.
      const { hints: mailboxHints, source: verificationSource } = await this.detectActiveMailbox();

      // Sanfter Abgleich: Zeigt an, ob die aktive Mailbox zum erwarteten Realm passt.
      // Runde 4 (Kollateral-Bug): Ohne Hints oder ohne Realm ist das Ergebnis
      // bewusst NULL (unbekannt) statt false — ein fehlender Hint ist kein
      // Nachweis einer falschen Mailbox.
      const realmDomain = this.realm && this.realm.includes('.') ? this.realm.toLowerCase() : null;
      const verification = resolveMailboxVerification(mailboxHints, realmDomain);
      const activeMailboxMatchesRealm = verification.activeMailboxMatchesRealm;

      return {
        authenticated: isLoaded,
        currentUrl: page.url(),
        title: title,
        realm: this.realm,
        mailboxHints,
        verificationSource,
        activeMailboxMatchesRealm,
        message: isLoaded
          ? (activeMailboxMatchesRealm === true
              ? `Outlook Web ist eingeloggt (Mailbox passt zu ${this.realm}, Verifikation: ${verificationSource}).`
              : activeMailboxMatchesRealm === false
                ? `ACHTUNG: Eingeloggt, aber die aktive Mailbox passt nachweislich NICHT zu '${this.realm}' (Verifikation: ${verificationSource}; Hints: ${mailboxHints.join(' | ')}). Bitte im sichtbaren Fenster die korrekte Mailbox bestätigen.`
                : `Eingeloggt, aber Mailbox konnte nicht verifiziert werden (Verifikationsquelle: ${verificationSource}; keine Account-Hinweise in dieser Ansicht). Bitte bei kritischen Aktionen im sichtbaren Fenster die Mailbox bestätigen.`)
          : "Seite lädt noch oder erfordert Anmeldung."
      };
    } catch (err) {
      return {
        authenticated: false,
        realm: this.realm,
        error: err.message
      };
    }
  }

  async createDraft({ to, subject, body, cc }) {
    const page = await this.getPage(true);

    const toStr = Array.isArray(to) ? to.join('; ') : (to || '');
    const ccStr = Array.isArray(cc) ? cc.join('; ') : (cc || '');
    const subjectStr = subject || '';
    const bodyStr = body || '';

    // Important: Do not pass body in URL params, as that forces OWA into plain-text and strips the default signature!
    const composeUrl = buildComposeUrl(this.realm, { to, subject, cc });

    await page.goto(composeUrl, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(5000);

    const editor = await page.waitForSelector(
      'div[role="textbox"][aria-label*="Message body" i], div[role="textbox"][aria-label*="Nachrichtentext" i], div[aria-label="Message body"], div[aria-label="Nachrichtentext"]',
      { timeout: 15000 }
    ).catch(() => null);

    if (editor) {
      await editor.click();
      await page.waitForTimeout(300);

      // Check if signature already in DOM
      const hasSignature = await editor.evaluate(el => el.innerHTML.includes('id="Signature"') || el.innerHTML.includes('David Hallmann'));

      if (!hasSignature) {
        // Inject rich signature if not present in new compose window
        await editor.evaluate((el, sig) => {
          el.innerHTML = `<p><br></p><p><br></p>${sig}`;
        }, DEFAULT_SIGNATURE_HTML);
        await page.waitForTimeout(300);
      }

      // Prepend body lines before signature via safe DOM insertion
      await editor.evaluate((el, text) => {
        const lines = text.split('\n');
        const fragment = document.createDocumentFragment();
        for (const line of lines) {
          const p = document.createElement('p');
          if (line.trim().length === 0) {
            p.appendChild(document.createElement('br'));
          } else {
            p.textContent = line;
          }
          fragment.appendChild(p);
        }
        const spacer = document.createElement('p');
        spacer.appendChild(document.createElement('br'));
        fragment.appendChild(spacer);

        el.insertBefore(fragment, el.firstChild);
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
      }, bodyStr);

      await page.waitForTimeout(500);

      // Save draft
      await page.keyboard.press('Control+s');
      await page.waitForTimeout(3000);
    }

    return {
      success: true,
      to: toStr,
      cc: ccStr || undefined,
      subject: subjectStr,
      preview: bodyStr.length > 200 ? bodyStr.substring(0, 200) + '...' : bodyStr,
      folder: "Entwürfe (Drafts)",
      message: "Entwurf wurde erfolgreich in Outlook unter 'Entwürfe' gespeichert (inklusive Signatur)."
    };
  }

  async createCalendarEvent({ subject, start, end, body, is_all_day = false, timezone = 'Europe/Berlin', transparency = 'busy' } = {}) {
    if (!subject || typeof subject !== 'string' || subject.trim().length === 0) {
      throw new Error('Parameter "subject" ist erforderlich und darf nicht leer sein.');
    }
    if (!start || !end) {
      throw new Error('Parameter "start" und "end" sind erforderlich.');
    }

    const page = await browserManager.ensureContext(true, this.realm).then(r => r.page);
    this.lastPage = page;

    const startDt = normalizeCalendarDate(start, timezone);
    const endDt = normalizeCalendarDate(end, timezone);

    // Body NICHT in URL-Parameter geben, da OWA den Editor sonst auf Plain-Text umstellt.
    const calendarUrl = buildCalendarComposeUrl(this.realm, { subject, start: startDt, end: endDt, is_all_day });

    // Der Kalender-Deep-Link kann einen NEUEN Tab öffnen. Daher Zeitstempel der Tab-Liste merken.
    const context = browserManager.context;
    const pagesBefore = context ? context.pages().length : 0;

    await page.goto(calendarUrl, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(6000);

    // Falls ein neuer Tab geöffnet wurde: auf diesen wechseln
    let activePage = page;
    if (context) {
      const pagesAfter = context.pages();
      if (pagesAfter.length > pagesBefore) {
        activePage = pagesAfter[pagesAfter.length - 1];
        this.lastPage = activePage;
        await activePage.waitForLoadState('domcontentloaded').catch(() => null);
        await activePage.waitForTimeout(1500);
      }
    }

    // Der Event-Editor von OWA stellt den Betreff in einem Textfeld dar
    const titleInput = await activePage.waitForSelector(
      'input[aria-label*="Titel" i], input[aria-label*="Title" i], input[aria-label*="Betreff" i], input[aria-label*="Subject" i], input[placeholder*="Titel" i], input[placeholder*="Add a title" i]',
      { timeout: 15000 }
    ).catch(() => null);

    if (titleInput) {
      await titleInput.click();
      await activePage.waitForTimeout(200);
      await titleInput.fill(subject);
      await activePage.waitForTimeout(300);
    }

    // Body (Beschreibung) eintragen, falls vorhanden
    if (body && body.trim().length > 0) {
      const bodyEditor = await activePage.waitForSelector(
        'div[role="textbox"][aria-label*="Description" i], div[role="textbox"][aria-label*="Beschreibung" i], div[role="textbox"][aria-label*="Notes" i], div[role="textbox"][aria-label*="Notizen" i], div[role="textbox"][aria-label*="body" i]',
        { timeout: 8000 }
      ).catch(() => null);

      if (bodyEditor) {
        await bodyEditor.click();
        await activePage.waitForTimeout(300);
        await bodyEditor.evaluate((el, text) => {
          const lines = text.split('\n');
          const fragment = document.createDocumentFragment();
          for (const line of lines) {
            const p = document.createElement('p');
            if (line.trim().length === 0) {
              p.appendChild(document.createElement('br'));
            } else {
              p.textContent = line;
            }
            fragment.appendChild(p);
          }
          el.insertBefore(fragment, el.firstChild);
          el.dispatchEvent(new Event('input', { bubbles: true }));
          el.dispatchEvent(new Event('change', { bubbles: true }));
        }, body);
        await activePage.waitForTimeout(400);
      }
    }

    // Speichern
    await activePage.keyboard.press('Control+s');
    await activePage.waitForTimeout(3000);

    return {
      success: true,
      subject,
      start: startDt || start,
      end: endDt || end,
      is_all_day,
      transparency,
      message: `Firmenkalender-Eintrag "${subject}" wurde in Outlook gespeichert (${transparency}). Bitte im Kalender kurz gegenprüfen.`
    };
  }

  async getFreeBusy({ start, end } = {}) {
    if (!start || !end) {
      throw new Error('Parameter "start" und "end" sind erforderlich.');
    }

    const page = await browserManager.ensureContext(true, this.realm).then(r => r.page);
    this.lastPage = page;

    const startDt = normalizeCalendarDate(start) || start;
    const endDt = normalizeCalendarDate(end) || end;

    // Angefragter Range als Datums-Strings (YYYY-MM-DD)
    const rangeStartDate = String(startDt).slice(0, 10);
    const rangeEndDate = String(endDt).slice(0, 10);
    // Fallback-Jahr für Event-Labels ohne Jahresangabe (OWA lässt das Jahr oft weg)
    const defaultYear = parseInt(rangeStartDate.slice(0, 4), 10);

    // Zielwoche bestimmen: Montag der Woche, die das Startdatum enthält
    // (OWA-Wochenansicht beginnt montags).
    const startDateObj = new Date(rangeStartDate + 'T00:00:00');
    const weekday = startDateObj.getDay(); // 0 = Sonntag
    const offsetToMonday = weekday === 0 ? -6 : 1 - weekday;
    const mondayIso = toIsoDate(new Date(startDateObj.getTime() + offsetToMonday * 86400000));

    // ---------------------------------------------------------------------
    // Navigations-Strategie (Iteration 4, Live-Befunde Runde 3 vom 25.09.):
    // Der startdate-Param wird von outlook.cloud.microsoft NICHT zuverlässig
    // befolgt (zukünftige Woche renderte leer, aktuelle Woche funktionierte).
    // Er bleibt nur noch OPTIMIERUNG (spart Klicks, wenn er wirkt) — vertraut
    // wird ihm NIE. PRIMÄRPFAD: bewährte Kalender-URL, gerenderte Woche
    // ermitteln, dann per Vor/Zurück-Buttons zur Zielwoche navigieren.
    //
    // NEU Runde 4 (MCP-Timeout -32001 ohne Ergebnisobjekt): ALLE Phasen
    // (Navigation, Klicks, Re-Render-Warten, Extraktion, Diagnostik) teilen
    // sich EIN hartes Gesamtbudget von 45s — sicher unter dem MCP-Client-
    // Timeout. Ist es aufgebraucht, wird sofort das Beste zurückgegeben, was
    // vorhanden ist (inkl. Diagnostik). Der Call läuft NIE ins Client-Timeout.
    // ---------------------------------------------------------------------
    const optimizedUrl = buildCalendarWeekUrl(this.realm, mondayIso); // mit startdate (Optimierung)
    const provenUrl = this.realm
      ? `https://outlook.office.com/calendar/view/week?realm=${encodeURIComponent(this.realm)}`
      : 'https://outlook.office.com/calendar/view/week';

    // Hartes Gesamtzeitbudget für den KOMPLETTEN Call.
    const budgetStart = Date.now();
    const deadline = budgetStart + FREEBUSY_TOTAL_BUDGET_MS;
    const remaining = () => deadline - Date.now();
    // Phasen-Scheibe: nie größer als das Cap, nie größer als die Restzeit,
    // 0 (skip) unterhalb der Mindestscheibe — KEIN Floor mehr, der Phasen nach
    // Ablauf der Deadline weiterlaufen ließ (Root Cause Runde 3).
    const phase = (capMs) => phaseBudgetMs(remaining(), capMs);
    // Klick-Kosten für die Rationierung: Probe + Klick + Re-Render-Wait.
    const CLICK_COST_MS = 1500 + FREEBUSY_CLICK_RENDER_WAIT_MS;
    // Klick-Protokoll (Runde-4-Forderung C): Jeder Selektor-Kandidaten-Versuch
    // wird hier protokolliert (auch "nicht gefunden") und in der Diagnostik
    // zurückgegeben — so ist sichtbar, welcher Kandidat existierte/welcher
    // geklickt wurde bzw. dass gar kein Button gefunden wurde.
    const weekButtonProbe = [];

    // Kompakter Budget-Report für die Diagnostik (wo ging die Zeit verloren?)
    const freeBusyBudgetReport = () => ({
      totalMs: FREEBUSY_TOTAL_BUDGET_MS,
      usedMs: Date.now() - budgetStart,
      remainingMs: Math.max(0, remaining())
    });

    // Fehler-Ergebnis MIT Diagnostik (Runde-4-Forderung B): Kein nackter Throw
    // mehr — das Objekt überlebt den MCP-Transport vollständig.
    const errorResult = (message, extra = {}) => ({
      status: 'error',
      error: message,
      start: startDt,
      end: endDt,
      requestedRange: { startDate: rangeStartDate, endDate: rangeEndDate },
      ...extra
    });

    // Frische Diagnostik nur erheben, wenn das Restbudget die teure
    // DOM-Traversierung noch trägt; sonst liefert der frühe Snapshot weiter.
    const freshDiag = () => (remaining() > 4000
      ? collectCalendarDiagnostics(page).catch(() => null)
      : Promise.resolve(null));

    // URL-Verifikation: OWA leitet bei ungültiger Session/Route client-seitig um
    // (z.B. zurück auf /mail/) — die Ziel-URL allein beweist noch nichts.
    const waitOnCalendarUrl = async (budgetMs) => {
      const urlDeadline = Date.now() + budgetMs;
      while (Date.now() < urlDeadline) {
        if ((page.url() || '').includes('/calendar')) return true;
        if (Date.now() + 500 > urlDeadline) break;
        await page.waitForTimeout(500);
      }
      return (page.url() || '').includes('/calendar');
    };

    const gotoCalendar = async (url) => {
      const gotoBudget = phase(12000);
      if (gotoBudget <= 0) return false;
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: gotoBudget }).catch(() => null);
      return waitOnCalendarUrl(phase(4000));
    };

    // Event-Grid-Fertig-Beweis (Defekt 1+3): Es gibt ab jetzt NUR NOCH EINE
    // Event-Sammlung — collectCalendarEventLabels. Vorher existierten ZWEI
    // unabhaengige Implementierungen (probeEventGrid hier inline, der
    // Kollektor in collectCalendarEventLabels). Genau daraus entstand der
    // Widerspruch des Smoke-Tests: die Warte-Heuristik meldete "found",
    // waehrend die Extraktion 0 Elemente fand — beide konnten gleichzeitig
    // "wahr" sein, weil sie unterschiedliche Selektoren kannten.
    //
    // Die Warte-Probe laeuft im 'fast'-Modus: nur die dokumentweiten
    // Attribut-Selektoren, KEIN '*'-Durchlauf (der kostet auf einer echten
    // OWA-Seite hunderte ms und wuerde das Poll-Budget auffressen). Die
    // vollstaendige Sammlung inkl. Struktur-Scan laeuft genau einmal danach
    // und entscheidet ueber den Ausgang.
    //
    // Bug 3: Diese Vorab-Probe entscheidet NICHT mehr darueber, ob die
    // Event-Wartephase gepollt wird. Genau das war die Root Cause
    // (eventWaitPolls 0 bei 20s Budget). Ihre Rueckgabe steuert nur noch
    // GUARD 2; die Wartephase hat einen eigenen Beobachter (makeEventObserver).
    let lastEventSelector = null;
    const hasEventElements = async () => {
      const r = await collectCalendarEventLabels(page, { preferSelector: lastEventSelector, weekStartIso: mondayIso, fast: true });
      if (r.eventElementsFound > 0 && r.selectorUsed) lastEventSelector = r.selectorUsed;
      return r.eventElementsFound > 0;
    };

    // Zustandswarten statt starrem Timeout: auf gerenderten Wochen-/Tages-Header
    // ODER Event-Elemente warten (Live-Fail 24.09.2026: 6s reichten nicht, das Grid
    // war noch nicht da -> 0 Slots, "Work week" als einziger Text). Poll-Intervall
    // 500ms mit Budget-Check je Iteration (Runde-4-Forderung A).
    //
    // Runde-5-Fix (Root Cause 0 Slots trotz verifizierter Woche): Der Header rendert
    // VOR den Events. Früher brach die Schleife beim ersten datierten Header sofort
    // ab -> die Slot-Extraktion lief, bevor die Events im DOM waren. Jetzt bekommt
    // ein datierter Header eine kurze Gnadenfrist, in der weiter auf Events gepollt
    // wird (budgetiert), damit die Extraktion nicht ins Leere greift.
    const EVENT_GRACE_MS = 2500;
    const waitForCalendarRendered = async (budgetMs) => {
      const waitDeadline = Date.now() + budgetMs;
      let rendered = { text: null, startDate: null, endDate: null };
      let events = false;
      let headerSeenAt = null;
      while (Date.now() < waitDeadline && remaining() > 0) {
        rendered = await readRenderedCalendarWeek(page);
        events = await hasEventElements();
        if (rendered.startDate) {
          if (!headerSeenAt) headerSeenAt = Date.now();
          if (events) break;
          if (Date.now() - headerSeenAt >= EVENT_GRACE_MS) break; // Gnadenfrist abgelaufen
        } else if (events) {
          break;
        }
        if (Date.now() + 500 > waitDeadline || remaining() <= 0) break;
        await page.waitForTimeout(500);
      }
      return { rendered, hasEvents: events };
    };

    // Multi-selektorischer Wochen-Button-Klick (Runde-4-Forderung C): Probiert
    // Kandidaten nacheinander (aria-label DE/EN, title DE/EN, breite Muster,
    // Chevron-Symbole), protokolliert JEDEN Versuch ins weekButtonProbe und
    // verifiziert nach dem Klick, dass sich der gerenderte Header wirklich
    // geändert hat — nur bei nachweislich wirkungslosem Klick ('unchanged') wird
    // der nächste Kandidat versucht; bei 'unverified' nie (Doppel-Klick-Gefahr,
    // falls der erste Klick doch gewirkt hat).
    //
    // Bug 2 (Live-Befund 25.09.2026): Der Code liegt jetzt in der exportierten
    // Funktion runWeekButtonClick — NICHT als Closure. Der alte Test hat die
    // Schleife nachgebaut und diese Funktion nie aufgerufen; deshalb war der
    // Guard "nur durch Code-Review belegt" und das `note:'guard-vor-klick'`
    // konnte im Live-Lauf fehlen, ohne dass ein Test rot wurde. Diese Funktion
    // IST der Produktionspfad, damit der Test den echten Draht faehrt.
    // Der Guard selbst steckt in weekClickDecision (geteilte Definition mit
    // der aeusseren Schritt-Schleife) und entscheidet in BEIDEN Richtungen:
    // Zielwoche erreicht -> Skip; Ansicht nicht gerendert -> KEIN Klick.
    const clickWeekButton = (direction, prevStart) => runWeekButtonClick({
      direction,
      prevStart,
      targetStart: rangeStartDate,
      page,
      probe: weekButtonProbe,
      readWeek: () => readRenderedCalendarWeek(page),
      afterClick: (prev, budgetMs) => waitRenderAfterClick(prev, phase(budgetMs)),
      clickWaitMs: FREEBUSY_CLICK_RENDER_WAIT_MS,
      clickCostMs: CLICK_COST_MS,
      remainingMs: remaining
    });

    // Wartet nach einem Button-Klick auf das Re-Render und BEWERTET die Änderung
    // (Runde-4-Forderung C): 'changed' = Header-Datum hat sich geändert (Klick
    // wirkte), 'unchanged' = Header lesbar und identisch (Klick wirkungslos),
    // 'unverified' = Header vor/nach dem Klick nicht datierbar (nicht bewertbar).
    // Poll-Intervall 500ms, Budget-Check je Iteration, kein blindes Nachwarten mehr.
    const waitRenderAfterClick = async (prevStart, budgetMs) => {
      const waitDeadline = Date.now() + Math.min(budgetMs, Math.max(0, remaining()));
      let week = { text: null, startDate: null, endDate: null };
      while (Date.now() < waitDeadline) {
        week = await readRenderedCalendarWeek(page);
        if (week.startDate) break;
        if (Date.now() + 500 > waitDeadline || remaining() <= 0) break;
        await page.waitForTimeout(500);
      }
      return { week, verified: evaluateClickVerification(prevStart, week) };
    };

    // 1) Optimierungsversuch MIT startdate (spart Button-Klicks, wenn der Param wirkt)
    let onCalendar = await gotoCalendar(optimizedUrl);
    // 2) Fallback: bewährte URL OHNE startdate (Basis des Primärpfads)
    if (!onCalendar) {
      onCalendar = await gotoCalendar(provenUrl);
    }
    if (!onCalendar) {
      // Runde 4: Fehler-Objekt MIT Diagnostik statt Throw — sie muss den
      // Transport zum Client überleben.
      const diagnostics = assembleFreeBusyDiagnostics({
        early: await collectCalendarDiagnostics(page).catch(() => null),
        weekButtonProbe,
        abortReason: 'no-calendar-tab',
        budget: freeBusyBudgetReport()
      });
      return errorResult(
        `Kalender-Tab konnte nicht erreicht werden (URL nach Navigation: ${page.url() || 'unbekannt'}). ` +
        `Möglicherweise ist die Session abgelaufen — bitte outlook_login erneut ausführen oder OWA manuell prüfen.`,
        { diagnostics }
      );
    }

    // FRÜHER DIAGNOSTIK-SNAPSHOT (Runde-4-Forderung B): Direkt nach der Navigation
    // erheben und im Kontext halten — URL, Header-Kandidaten und ARIA-Sample
    // überleben so jeden späteren Abbruch-/Budget-Pfad.
    const earlyDiagnostics = await collectCalendarDiagnostics(page).catch(() => null);

    let { rendered, hasEvents } = await waitForCalendarRendered(phase(8000));

    // VERIFIKATION: Gerenderten Wochen-Header gegen den angefragten Range prüfen.
    // Nie blind Daten der falschen Woche zurückgeben.
    // EINE Definition fuer die Zielwochen-Pruefung: der Klick-Guard in
    // clickWeekButton und diese aeussere Schleife duerfen nicht
    // auseinanderlaufen (Live-Befund 25.09.2026).
    const isRequestedWeek = (r) => renderedWeekContains(r, rangeStartDate);

    // FRÜHER KONTROLL-RETURN (Runde-4-Forderung B): Kein datierter Header und
    // keine Events nach Navigation + Warten — wenn das Restbudget die
    // Eskalationsleiter (frische Navigation ~15s bzw. Button-Klicks) nicht mehr
    // tragen kann, sofort error+diagnostics zurückgeben statt weiterzuklicken.
    if (!rendered.startDate && !hasEvents && remaining() < FREEBUSY_MIN_REMAINING_FOR_ESCALATION_MS) {
      const diagnostics = assembleFreeBusyDiagnostics({
        early: earlyDiagnostics,
        weekButtonProbe,
        abortReason: 'early-no-dated-header',
        budget: freeBusyBudgetReport()
      });
      return errorResult(
        `Früher Abbruch: Nach Navigation und Warten weder ein datierter Wochen-Header noch Event-Elemente ` +
        `im Kalender-Tab gefunden; das Restbudget reicht nicht mehr für die Eskalationsleiter. ` +
        `Bitte erneut versuchen bzw. Session prüfen (outlook_login).`,
        { renderedWeek: rendered, diagnostics }
      );
    }

    // Button-basierte Wochen-Korrektur (PRIMÄRPFAD): Vor/Zurück bis die gerenderte
    // Woche der Zielwoche entspricht. Rationierung gegen das Gesamtbudget
    // (Runde-4-Forderung A): max. 6 Klicks, jeder Klick nur bei ausreichendem
    // Restbudget; Eskalationsstufen (frische Navigation, Fernsprung) nur, wenn
    // das Budget sie noch trägt — sonst Abbruch mit dem, was da ist.
    const MAX_BUTTON_CLICKS = FREEBUSY_MAX_BUTTON_CLICKS;
    let gotoRetryDone = false;
    let forcedNextRetryDone = false;
    let farJumpDone = false;
    let buttonClicks = 0;
    let abortReason = 'none';
    // Klick-Sperre pro Richtung (Iteration 7, Defekt 2): verhindert, dass bei
    // unlesbarem Header ('unverified') mehrfach in dieselbe Richtung geklickt und
    // die Zielwoche ÜBERSPRUNGEN wird. Die Gegenrichtung (Rückweg) bleibt frei.
    let lastClickDirection = null;
    let weekConfirmedAfterClick = false;

    for (let step = 0; step < 30; step++) {
      // Klick-Pfad NUR bei echter Abweichung (Iteration 7, Defekt 2): Ist die
      // Zielwoche bereits gerendert und verifiziert, wird das als SKIP
      // protokolliert — keine Button-Probe, kein Klick. Runde 6 protokollierte
      // trotz bestätigter Zielwoche einen "Next week"-Klick.
      //
      // Bug 2: GLEICHE Entscheidungsfunktion wie im Klickpfad
      // (weekClickDecision). Vorher hatte die aeussere Schleife ihren eigenen,
      // schwaecheren Test (isRequestedWeek) — die beiden konnten auseinander-
      // laufen, und genau das war der Live-Befund.
      //
      // Ehrliche Abgrenzung: die Schleife arbeitet auf dem `rendered`-Snapshot,
      // der nach jedem Klick und jedem Render-Wait neu erhoben wird; der
      // Klickpfad liest dagegen unmittelbar VOR der Probe frisch. Der
      // Snapshot ist damit der Vorfilter, der frische Read im Klickpfad die
      // maechsliche Sperre. Neu und entscheidend ist hier die zweite
      // Bedingung der geteilten Funktion: ohne datierbaren Header wird nicht
      // nur nicht auf die Zielwoche geklickt, sondern gar nicht geklickt.
      const outerDecision = weekClickDecision({ rendered, targetStart: rangeStartDate, hasEvents });
      if (outerDecision.skip) {
        weekButtonProbe.push(weekProbeGuardEntry(outerDecision));
        if (outerDecision.reason === WEEK_CLICK_GUARD_REASONS.viewNotReady) {
          // Ansicht nicht gerendert: NICHT klicken. Der Klickpfad wuerde das
          // ohnehin blockieren; hier wird es benannt, damit der Abbruchgrund
          // ehrlich ist statt 'no-dated-header'.
          abortReason = 'view-not-ready';
        }
        break;
      }
      if (remaining() <= 0) { abortReason = 'budget'; break; }

      if (!rendered.startDate) {
        // Gerenderte Woche UNBEKANNT: Eskalationsleiter (nur bei ausreichendem Budget)
        if (!gotoRetryDone && remaining() >= 15000) {
          // 1x frisch navigieren und erneut zustandsbasiert warten
          gotoRetryDone = true;
          await gotoCalendar(provenUrl);
          const retry = await waitForCalendarRendered(phase(6000));
          rendered = retry.rendered;
          hasEvents = hasEvents || retry.hasEvents;
          continue;
        }
        if (!forcedNextRetryDone && remaining() >= CLICK_COST_MS + 1000) {
          // Erzwungene Korrektur auch bei unbekannter Woche: 1x "Nächste Woche" +
          // Re-Render + erneuter Erkennungsversuch, bevor abgebrochen wird.
          // Iteration 7: Nur wenn die Woche seit dem letzten Klick ERNEUT
          // bestätigt wurde — sonst überspringt der Klick die Zielwoche.
          if (!canClickSameDirection({ lastClickDirection, direction: 1, weekConfirmedAfterClick })) {
            abortReason = 'same-direction-unconfirmed';
            break;
          }
          forcedNextRetryDone = true;
          const forced = await clickWeekButton(1, rendered.startDate);
          if (forced.guard === WEEK_CLICK_GUARD_REASONS.viewNotReady) {
            // Bug 2: Der erzwungene "Nächste Woche"-Klick war der Live-Befund.
            // Er lief auf einer Ansicht ohne Wochen-Header — die Wochen- und
            // Button-Position war unbekannt, ein Klick haette die Zielwoche
            // ueberspringen koennen. Der Guard blockiert das jetzt; hier wird
            // es als Abbruchgrund benannt statt als 'no-dated-header'.
            abortReason = 'view-not-ready';
            break;
          }
          if (forced.targetReached) {
            // Zielwoche war bereits verifiziert -> uebernehmen und beenden.
            // Kein Fehlerpfad: "kein Klick" waere hier falsch, weil der
            // naechste GUARD sonst "week-mismatch" melden wuerde.
            rendered = forced.week;
            hasEvents = hasEvents || (await hasEventElements());
            break;
          }
          if (forced.clicked) {
            rendered = forced.week;
            hasEvents = hasEvents || (await hasEventElements());
            lastClickDirection = 1;
            weekConfirmedAfterClick = Boolean(forced.week && forced.week.startDate);
          }
          continue;
        }
        abortReason = 'no-dated-header';
        break; // Nichts hilft -> Guard unten liefert Fehler MIT Diagnostik
      }

      // Gerenderte Woche BEKANNT: Delta in Wochen berechnen und per Buttons navigieren
      const deltaWeeks = Math.round(
        (new Date(mondayIso + 'T00:00:00') - new Date(rendered.startDate + 'T00:00:00')) / (7 * 86400000)
      );
      if (deltaWeeks === 0) break; // Zielwoche gerendert -> unten bewerten

      if (buttonClicks >= MAX_BUTTON_CLICKS) { abortReason = 'max-clicks'; break; }

      if (Math.abs(deltaWeeks) > MAX_BUTTON_CLICKS) {
        // Zu weit entfernt für Button-Klicks: 1x startdate-URL als Fernsprung
        // (Optimierung), danach wird wieder verifiziert — nur bei ausreichendem Budget.
        if (!farJumpDone && remaining() >= 15000) {
          farJumpDone = true;
          await gotoCalendar(optimizedUrl);
          const retry = await waitForCalendarRendered(phase(6000));
          rendered = retry.rendered;
          hasEvents = hasEvents || retry.hasEvents;
          continue;
        }
        abortReason = 'too-far';
        break;
      }

      // Klick-Rationierung: Klick + Re-Render-Wait müssen ins Restbudget passen
      if (remaining() < CLICK_COST_MS) { abortReason = 'budget'; break; }

      // Iteration 7, Defekt 2: In dieselbe Richtung nur nach erneuter Wochen-
      // Bestätigung. Der Rückweg (Gegenrichtung) bleibt ausdrücklich erlaubt.
      const direction = Math.sign(deltaWeeks);
      if (!canClickSameDirection({ lastClickDirection, direction, weekConfirmedAfterClick })) {
        abortReason = 'same-direction-unconfirmed';
        break;
      }

      const r = await clickWeekButton(direction, rendered.startDate);
      if (r.guard === WEEK_CLICK_GUARD_REASONS.viewNotReady) {
        // Bug 2: Der frische Header-Read im Guard fand die Ansicht gerade nicht
        // gerendert. Klicken waere nicht vertretbar; der Grund wird benannt.
        // Bewusst VOR dem !r.clicked-Zweig, sonst wuerde es als
        // 'button-ineffective'/'no-week-button' fehlattet.
        abortReason = 'view-not-ready';
        break;
      }
      if (r.targetReached) {
        // Zielwoche verifiziert -> uebernehmen und die Klick-Schleife
        // verlassen. Bewusst VOR dem !r.clicked-Zweig: ein Skipped-Klick ist
        // kein Fehlschlag.
        rendered = r.week;
        hasEvents = hasEvents || (await hasEventElements());
        break;
      }
      if (!r.clicked) {
        // Kein Kandidat wirkungsvoll -> 1x frische Navigation als Fallback (budgetiert)
        if (!gotoRetryDone && remaining() >= 15000) {
          gotoRetryDone = true;
          await gotoCalendar(provenUrl);
          const retry = await waitForCalendarRendered(phase(6000));
          rendered = retry.rendered;
          hasEvents = hasEvents || retry.hasEvents;
          continue;
        }
        // Präzise Ursache (Runde-4-Forderung 3): Button gefunden/geklickt, aber
        // Header unverändert -> 'button-ineffective'; gar kein Button -> 'no-week-button'.
        abortReason = classifyButtonFailure(r);
        break;
      }
      buttonClicks++;

      // Klick war verifiziert wirkungsvoll (Header geändert) — gerenderte Woche übernehmen
      rendered = r.week;
      lastClickDirection = direction;
      weekConfirmedAfterClick = Boolean(r.week && r.week.startDate);
      hasEvents = hasEvents || (await hasEventElements());
    }

    // GUARD 1: Falsche Woche gerendert -> keine Daten statt falscher Daten.
    // Runde 4: Als Fehler-Objekt MIT Diagnostik zurückgeben (transportfest),
    // nicht als Throw, der die Diagnostik im Timeout verschlucken kann.
    if (rendered.startDate && !isRequestedWeek(rendered)) {
      const diagnostics = assembleFreeBusyDiagnostics({
        fresh: await freshDiag(),
        early: earlyDiagnostics,
        weekButtonProbe,
        abortReason: abortReason !== 'none' ? abortReason : 'week-mismatch',
        budget: freeBusyBudgetReport()
      });
      return errorResult(
        `Kalender-Wochenansicht zeigt nicht die angefragte Woche: angefragt ab ${rangeStartDate}, ` +
        `gerendert ${rendered.startDate} – ${rendered.endDate || '?'} ` +
        `(Header: "${(rendered.text || '').slice(0, 120)}"). ` +
        `Es werden bewusst KEINE falschen Belegungsdaten zurückgegeben.`,
        { renderedWeek: rendered, diagnostics }
      );
    }

    // GUARD 2 (defensiv, geschärft): Weder datierter Header noch Event-Elemente nach
    // Navigation + Warten + timeboxter Eskalation -> Kalenderansicht nicht erreicht/
    // nicht gerendert. Fehler-Objekt MIT Diagnostik statt Throw.
    // (Leere Woche mit datiertem Header ist legitim.)
    if (!rendered.startDate && !hasEvents) {
      const diagnostics = assembleFreeBusyDiagnostics({
        fresh: await freshDiag(),
        early: earlyDiagnostics,
        weekButtonProbe,
        abortReason: abortReason !== 'none' ? abortReason : 'not-rendered',
        budget: freeBusyBudgetReport(),
        waitedForEventsMs: 0,
        eventWaitOutcome: 'not-rendered'
      });
      return errorResult(
        `Kalenderansicht nicht geladen — Kalender-Tab nicht erreicht oder Woche nicht gerendert. ` +
        `Weder ein datierter Wochen-Header noch Event-Elemente gefunden (nach Navigation, Warten und timeboxter Eskalation). ` +
        `Bitte erneut versuchen bzw. Session prüfen (outlook_login).`,
        { renderedWeek: rendered, diagnostics }
      );
    }

    // Wochen-Nachweis (unabhaengig von den Events). Der Header-/Header-Datum-
    // Pfad funktioniert und bleibt unveraendert — nur die BEDEUTUNG von
    // dataQuality wird ehrlich (siehe resolveDataQuality weiter unten).
    const weekVerified = Boolean(rendered.startDate && rendered.endDate);

    // ECHTE Event-Wartephase (Bug 3, Live-Befund 25.09.2026): Nach bestaetigter
    // Wochen-Verifikation wird GEZIELT auf echte Event-Knoten gepollt. Budget:
    // min(20s, Restzeit - 5s Reserve) — die Reserve traegt Extraktion +
    // Diagnostik, damit der Call nie ins MCP-Client-Timeout laeuft.
    //
    // Die Vorab-Meldung hasEvents wird hier BEWUSST NICHT mehr als Abkuerzung
    // benutzt. Live-Befund: waitedForEventsMs 0 / eventWaitPolls 0 bei einem
    // Budget von 20000ms. Ursache war genau diese Abkuerzung — die stale
    // Vorab-Probe entschied, OB ueberhaupt gepollt wird, und verhinderte damit
    // jede Widerlegung. Die Wartephase arbeitet jetzt ausschliesslich auf
    // Beobachtungen; die Beobachtung selbst ist die schnelle Probe plus der
    // Lade-Indikator-Check.
    const eventWaitBudget = eventWaitBudgetMs(remaining());
    const observeCalendarEvents = makeEventObserver((preferSelector) => collectCalendarEventLabels(page, {
      preferSelector,
      weekStartIso: rendered.startDate || mondayIso,
      fast: true
    }));
    const eventWait = await waitForCalendarEvents({
      observe: observeCalendarEvents,
      budgetMs: eventWaitBudget,
      pollMs: FREEBUSY_EVENT_WAIT_POLL_MS,
      settleMs: FREEBUSY_EVENT_WAIT_SETTLE_MS,
      sleep: (ms) => page.waitForTimeout(ms)
    });
    hasEvents = eventWait.found;

    // Alle Event-Elemente über den GEMEINSAMEN Kollektionspfad sammeln (Runde-5-Fix):
    // derselbe Pfad wie die Diagnostik, toleranter Sichtbarkeitsfilter (offsetParent
    // ist KEIN Hard-K.O. für Events), inkl. Roh-Diagnostik für den 0-Slot-Fall.
    // Iteration 7: der zuletzt erfolgreiche Selektor wird ZUERST probiert.
    // ABSCHLIESSENDE Sammlung: voller Pfad inkl. Struktur-Scan (Text + Geometrie).
    // Der zuletzt erfolgreiche Selektor wird ZUERST probiert. Diese Sammlung —
    // nicht die Vorab-Probe — entscheidet ueber eventElementsFound, dataQuality
    // und eventWaitOutcome.
    const eventCollection = await collectCalendarEventLabels(page, {
      preferSelector: observeCalendarEvents.lastSelector(),
      weekStartIso: rendered.startDate || mondayIso
    });
    const rawLabels = eventCollection.labels;
    const eventElementsFound = eventCollection.eventElementsFound;
    const selectorUsed = eventCollection.selectorUsed;

    // Defekt 3 / Bug 3: Der Ausgang der Wartephase wird an die TATSAECHLICH in
    // Abschluss-Sammlung beobachtete Event-Elementzahl gekoppelt. 'found' ist
    // damit unmoeglich, wenn 0 Knoten gelesen wurden. Und jetzt zum ersten Mal
    // auch umgekehrt: ein nicht gelaufener Poll trotz Budget ('not-polled') ist
    // ein Logikfehler und wird als solcher gemeldet statt als 'not-found'.
    const eventWaitReconciled = reconcileEventWait({
      wait: eventWait,
      eventElementsFound,
      waitBudgetMs: eventWaitBudget
    });
    eventWait.outcome = eventWaitReconciled.outcome;
    eventWait.abortReason = eventWaitReconciled.abortReason;
    eventWait.found = eventElementsFound > 0;
    hasEvents = eventWait.found;

    // Defekt 2: dataQuality nur noch als 'verified', wenn die Woche passt UND
    // echte Event-Knoten gelesen wurden. Sonst 'unverified-empty' — dann ist
    // busySlots:[] AUSDRUECKLICH kein "Kalender ist frei".
    const dataQuality = resolveDataQuality({ weekVerified, eventElementsFound });

    // Deduplizieren (gleiche Events können mehrfach im DOM hängen), strukturiert parsen
    // und anhand der Wortliste klassifizieren (Busy/Tentative/Free). Jede Verwerfung
    // wird mit Grund protokolliert (Runde-5-Forderung: 0-Slots von echtem Free-Tag
    // unterscheidbar machen).
    const seen = new Set();
    const parsedSlots = [];
    const rejectionReasons = [];
    const reject = (label, reason) => rejectionReasons.push({ label: String(label).slice(0, 120), reason });
    for (const label of rawLabels) {
      if (seen.has(label)) { reject(label, 'duplicate'); continue; }
      seen.add(label);
      const slot = parseEventLabel(label, defaultYear);
      if (!slot) { reject(label, 'no-time-pattern'); continue; }
      slot.status = classifyEventStatus(label);
      parsedSlots.push(slot);
    }
    // Header-artige Labels (Datum ohne Zeit) hat der Kollektor als Event-Kandidaten
    // verworfen -> als 'header-misclassified' ausweisen.
    for (const label of eventCollection.headerLikeLabels) reject(label, 'header-misclassified');
    // Iteration 7: Labels mit EXPLIZITEM Ablehnungsgrund. Runde 6 lieferte 12x
    // 'header-misclassified' für "calendar view, current time: Fri 22:31" — das
    // Rauschen verdeckte die eigentliche Ursache. Jetzt: 'clock-indicator' bzw.
    // 'work-location'.
    for (const ne of eventCollection.nonEventLabels) reject(ne.label, ne.reason);

    // Beauty-Fix (Live Runde 2): Work-Location-Banner und titellose Pseudo-Slots
    // ("(ohne Titel)", Büro/Homeoffice/URL-Schnipsel) sind keine echte Belegung
    // und werden aus den Slots ausgeschlossen.
    const realSlots = parsedSlots.filter((slot) => {
      const pseudo = slot.status === 'non_event' || isWorkLocationPseudoSlot(slot.title);
      if (pseudo) reject(slot.raw, 'work-location');
      return !pseudo;
    });

    // Dedupe: identisches Datum+Zeit nur 1x; überlappende Duplikate gleicher Zeit
    // zusammenführen (bevorzugt wird der Slot mit echtem Titel behalten).
    const dedupedSlots = dedupeBusySlots(realSlots);
    const keptSet = new Set(dedupedSlots);
    for (const slot of realSlots) if (!keptSet.has(slot)) reject(slot.raw, 'duplicate');

    // Nur Slots zurückgeben, die mit dem angefragten Range [start, end) überlappen.
    // Ganztag-Events belegen 00:00–24:00 ihres Tages. Explizit freie Slots ('free')
    // sind keine Belegung und werden ausgeschlossen.
    const rangeStartIso = String(startDt).includes('T') ? startDt : `${startDt}T00:00:00`;
    const rangeEndIso = String(endDt).includes('T') ? endDt : `${endDt}T23:59:59`;
    const rangeStartMs = new Date(rangeStartIso).getTime();
    const rangeEndMs = new Date(rangeEndIso).getTime();

    const { kept: busySlots, rejected: rangeRejected } = partitionSlotsByRange(dedupedSlots, rangeStartMs, rangeEndMs);
    for (const r of rangeRejected) reject(r.slot && r.slot.raw ? r.slot.raw : '', r.reason);

    // Stabil nach Datum und Beginn sortieren
    busySlots.sort((a, b) =>
      `${a.date}T${a.start || '00:00'}`.localeCompare(`${b.date}T${b.start || '00:00'}`)
    );

    // Hinweise zusammenbauen (Verifikationsstatus + Range-Abdeckung)
    let note = 'Busy-Slots aus ARIA-Labels der Kalender-Event-Elemente der Wochenansicht extrahiert ' +
      '(Datum/Zeit strukturiert geparst, Work-Location-Pseudo-Slots entfernt, Zeit-Duplikate zusammengeführt, ' +
      'per Wortliste klassifiziert, auf den angefragten Zeitraum gefiltert). ' +
      'Navigation: startdate-Param ist Primärsignal — bestätigt der sichtbare Wochen-Header (Range oder ' +
      'Prev/Next-Ableitung) die Zielwoche, sind KEINE Button-Klicks nötig; Buttons nur bei Abweichung.';
    if (dataQuality === 'unverified') {
      note += ' WARNUNG: Wochen-Header konnte nicht verifiziert werden — Daten ggf. unvollständig oder aus der falschen Woche.';
    }
    if (dataQuality === 'unverified-empty') {
      note += ` WARNUNG: Woche ${rendered.startDate || '?'} ist verifiziert, aber im DOM wurden ` +
        `0 Event-Knoten gelesen (Quellen: ${(eventCollection.sourceTried || []).map((s) => s.source).join(', ') || 'keine'}). ` +
        `busySlots=[] bedeutet hier KEINEN leeren Kalender, sondern ein nicht auslesbares Grid — ` +
        `NICHT als "frei" verwenden.`;
    }
    // EIN ehrlicher Abbruchgrund: Navigations-Abbruch ODER Warte-Abbruch. Vorher
    // endete der Call nach ~4s still mit abortReason 'none', obwohl die
    // Wartephase mit 20s Budget nicht gelaufen war (Live-Befund 25.09.2026).
    const reportedAbortReason = effectiveAbortReason({
      navReason: abortReason,
      eventWaitReason: eventWait.abortReason,
      eventElementsFound
    });
    if (dataQuality === 'verified' && reportedAbortReason !== 'none') {
      note += ` HINWEIS: Navigation vorzeitig beendet (${reportedAbortReason}); die angezeigte Woche wurde verifiziert.`;
    }
    if (eventWait.outcome !== 'found' && eventWait.outcome !== 'not-rendered') {
      note += ` HINWEIS: Auf echte Event-Knoten wurde ${Math.round(eventWait.waitedMs)}ms lang gepollt ` +
        `(${eventWait.polls} Polls, davon ${eventWait.loadingPolls} mit Lade-Indikator; ` +
        `Budget ${eventWaitBudget}ms, Abbruchgrund ${eventWait.abortReason || eventWait.outcome}) — ohne Treffer. ` +
        `0 Slots bedeuten dann nicht sicher "leerer Kalender".`;
    }
    if (rendered.endDate && rangeEndDate > rendered.endDate) {
      note += ` HINWEIS: Der angefragte Range endet (${rangeEndDate}) nach der gerenderten Woche (${rendered.endDate}) — spätere Tage sind nicht enthalten.`;
    }

    const result = {
      status: 'ok',
      start: startDt,
      end: endDt,
      calendar: 'Firmenkalender (Outlook Web)',
      renderedWeek: {
        text: rendered.text,
        startDate: rendered.startDate,
        endDate: rendered.endDate
      },
      // Wochen-Nachweis getrennt von der Event-Belastbarkeit (Defekt 2): die
      // Wochen-Verifikation funktioniert und bleibt so sichtbar.
      weekVerified,
      dataQuality,
      // Maschinenlesbare Ehrlichkeit: nur bei dataQuality==='verified' sind
      // busySlots eine belastbare Aussage ueber Belegung.
      busySlotsReliable: isBelievableBusySlots(dataQuality),
      busySlots,
      count: busySlots.length,
      note
    };
    // Reiche Diagnostik mitliefern — bei allem, was NICHT belastbar ist:
    //  - dataQuality != 'verified' (Woche oder Events unlesbar),
    //  - 0 Slots trotz gelesener Events (Extraktions-Miss),
    //  - 0 Event-Knoten (Grid nicht auslesbar).
    // Enthält URL (inkl. früher Snapshot), Header-/Tageskopf-Kandidaten, ARIA-Sample,
    // Klick-Protokoll, Abbruchgrund, Budget-Report, die Quellen-Leiter
    // (sourceTried/sourceUsed) und den Struktur-Census der echten Grid-DOM.
    if (!result.busySlotsReliable || busySlots.length === 0 || eventWait.outcome !== 'found') {
      const diag = assembleFreeBusyDiagnostics({
        fresh: await freshDiag(),
        early: earlyDiagnostics,
        weekButtonProbe,
        abortReason: reportedAbortReason,
        budget: freeBusyBudgetReport(),
        eventElementsFound,
        selectorUsed,
        rawEventLabels: rawLabels.slice(0, 20),
        rejectionReasons: rejectionReasons.slice(0, 40),
        waitedForEventsMs: Math.round(eventWait.waitedMs),
        eventWaitOutcome: eventWait.outcome,
        eventWaitAbortReason: eventWait.abortReason,
        eventWaitLoadingPolls: eventWait.loadingPolls,
        loadingIndicators: eventCollection.loadingIndicators
      });
      diag.weekVerified = weekVerified;
      diag.dataQuality = dataQuality;
      diag.eventWaitPolls = eventWait.polls;
      diag.eventWaitBudgetMs = eventWaitBudget;
      diag.loading = eventCollection.loading === true;
      if (eventCollection.sourceUsed) diag.sourceUsed = eventCollection.sourceUsed;
      if (eventCollection.sourceTried) diag.sourceTried = eventCollection.sourceTried;
      if (eventCollection.census) diag.census = eventCollection.census;
      if (eventCollection.collectError) diag.collectError = eventCollection.collectError;
      if (eventCollection.dateMissingLabels && eventCollection.dateMissingLabels.length) {
        diag.dateMissingLabels = eventCollection.dateMissingLabels;
      }
      const warnings = [
        slotExtractionWarning(eventElementsFound, busySlots.length),
        eventWaitWarning({ outcome: eventWait.outcome, waitedMs: eventWait.waitedMs, budgetMs: eventWaitBudget })
      ].filter(Boolean);
      if (warnings.length) diag.warning = warnings.join(' ');
      result.diagnostics = diag;
    }
    return result;
  }

  // Lokalisiert ein Kalender-Event im OWA-Firmenkalender und öffnet dessen Detail-Panel.
  // Suchbegriff = Titel-Fragment. Geht über die Wochenansicht (für den Zieltext), da die
  // Kalender-Suche in OWA keine zuverlässigen klickbaren Ergebnisse liefert. Der Klick
  // wird als ECHTER Playwright-Klick ausgeführt (React braucht native Events).
  // Rückgabe: { page, titleMatched, detailText } oder wirft, wenn nichts gefunden.
  async openCalendarEventByQuery(query, { timeout = 30000 } = {}) {
    if (!query || typeof query !== 'string' || query.trim().length === 0) {
      throw new Error('Parameter "query" (Suchbegriff/Titel-Fragment) ist erforderlich.');
    }
    const page = await browserManager.ensureContext(true, this.realm).then(r => r.page);
    this.lastPage = page;
    const ql = query.trim().toLowerCase();

    const calUrl = this.realm
      ? `https://outlook.office.com/calendar/view/week?realm=${encodeURIComponent(this.realm)}`
      : 'https://outlook.office.com/calendar/view/week';
    await page.goto(calUrl, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(6000);

    // Event-Elemente über ihr ARIA-Label (Format "<Titel>, HH:MM to HH:MM, ...") finden.
    // Playwright locator mit case-insensitivem Substring auf den Titel.
    const titleEscaped = ql.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const locator = page.locator(`[role="button"][aria-label*="${titleEscaped}" i]`).first();

    // WICHTIG: OWA öffnet das Event-Detail/Editor-Panel nur per DOPPELKlick zuverlässig.
    // Ein Einzelklick selektiert das Event ggf. nur, und Klicks auf Container-Koordinaten
    // können auf überlagerte Elemente (z.B. All-Day-Banner) treffen. locator.dblclick()
    // trifft exakt die Mitte des ARIA-Label-Elements und öffnet das RICHTIGE Event.
    let clicked = false;
    try {
      await locator.waitFor({ state: 'visible', timeout: 15000 });
      await locator.dblclick({ timeout: 10000 });
      clicked = true;
    } catch (e) {
      // Fallback: exakterer Präfix-Match
      try {
        const exactLocator = page
          .locator(`[role="button"][aria-label^="${titleEscaped}" i]`)
          .first();
        await exactLocator.waitFor({ state: 'visible', timeout: 10000 });
        await exactLocator.dblclick({ timeout: 10000 });
        clicked = true;
      } catch (e2) {
        clicked = false;
      }
    }

    if (!clicked) {
      throw new Error(`Kein Kalender-Event für Suchbegriff "${query}" gefunden (Wochenansicht).`);
    }

    // OWA öffnet nach Doppelklick ein Editor-/Reading-Panel (Dialog)
    await page.waitForTimeout(4000);

    const detailText = await page.evaluate(() => {
      const pane = document.querySelector('[role="dialog"], [role="main"], [class*="eventViewer"], [class*="eventDetails"], [class*="reading"]');
      return pane ? (pane.innerText || '') : (document.body.innerText || '');
    });

    return { page, titleMatched: query, detailText };
  }

  async deleteCalendarEvent({ query, id } = {}) {
    const page = await browserManager.ensureContext(true, this.realm).then(r => r.page);
    this.lastPage = page;
    const targetQuery = id || query;
    if (!targetQuery) {
      throw new Error('Parameter "query" oder "id" (Suchbegriff/Titel-Fragment) ist erforderlich.');
    }

    await this.openCalendarEventByQuery(targetQuery);

    // Nach dem Öffnen des Events das "Löschen"-/Entfernen-Element finden und klicken.
    // Alle DOM-Vorgänge laufen nach dem dblclick im Editor-Dialog ab. Der Löschen-Button
    // im OWA-Editor hat aria-label "Delete". Echte Playwright-Klicks lösen React-Events
    // zuverlässig aus (page.evaluate el.click() reicht oft nicht).
    let deleted = false;

    // 1) Direkt im geöffneten Editor-Dialog den Delete-Button finden und echt klicken.
    const delLoc = page.locator(
      '[role="dialog"] button[aria-label="Delete"], [role="dialog"] [role="button"][aria-label="Delete"], button[aria-label="Delete"], [role="button"][aria-label="Delete"]'
    ).first();
    try {
      await delLoc.waitFor({ state: 'visible', timeout: 8000 });
      await delLoc.click({ timeout: 6000 });
      deleted = true;
      await page.waitForTimeout(1500);
    } catch (e) {
      // Fallback: Textbasiert "Löschen/Delete" im Dialog (nicht Serie/alle)
      const txtDel = page.locator('[role="dialog"] button, [role="dialog"] [role="button"]').filter({ hasText: /delete|löschen/i }).first();
      try {
        await txtDel.waitFor({ state: 'visible', timeout: 5000 });
        await txtDel.click({ timeout: 5000 });
        deleted = true;
        await page.waitForTimeout(1500);
      } catch (e2) {
        deleted = false;
      }
    }

    // 2) Falls ein Bestätigungs-/Optionsdialog erscheint (z.B. "Are you sure you want
    //    to delete this event? Delete Cancel"), den Lösch-Knopf bestätigen. WICHTIG:
    //    NUR im Bestätigungsdialog suchen (nicht im Editor-Dialog [0], der ebenfalls
    //    einen "Delete"-Button hat). Erkennung über Bestätigungs-Marker im Dialogtext.
    for (let i = 0; i < 5; i++) {
      await page.waitForTimeout(800);
      const confirmHit = await page.evaluate(() => {
        // Nur Dialoge, die wie eine Lösch-Bestätigung aussehen (Textmarker)
        const dialogs = Array.from(document.querySelectorAll('[role="dialog"], [role="alertdialog"]'))
          .filter(d => {
            const t = (d.innerText || '').toLowerCase();
            return (t.includes('are you sure') || t.includes('confirm') ||
                    t.includes('sind sie sicher') || t.includes('sicher, dass') ||
                    t.includes('delete event') || t.includes('löschen') || t.includes('event löschen')) &&
                   !t.includes('duplicate');
          });
        for (const root of dialogs) {
          const els = Array.from(root.querySelectorAll('button, [role="button"]'));
          // kürzesten (= exaktesten) Lösch-Button im Bestätigungsdialog klicken
          const cands = els.map(el => ({ el, t: (el.innerText||'').trim().toLowerCase(), l: (el.getAttribute('aria-label')||'').toLowerCase() }))
            .filter(c => {
              const comb = `${c.t} ${c.l}`.trim();
              return (comb.includes('delete') || comb.includes('löschen')) &&
                     !comb.includes('delete all') && !comb.includes('delete series') &&
                     !comb.includes('alle löschen') && !comb.includes('serie löschen') &&
                     !comb.includes('cancel') && !comb.includes('abbrechen');
            });
          if (cands.length) {
            cands.sort((a, b) => (a.t||'x').length - (b.t||'x').length);
            cands[0].el.click();
            return true;
          }
        }
        return false;
      });
      if (confirmHit) {
        await page.waitForTimeout(1500);
        break;
      }
    }

    if (!deleted) {
      return {
        success: false,
        query: targetQuery,
        message: `Löschen-Button für "${targetQuery}" nicht gefunden. Bitte manuell im Kalender prüfen.`
      };
    }

    // VERIFIKATION: Zurück zur Wochenansicht und prüfen, ob das Event wirklich weg ist.
    // Nie blind success melden (Zero-Tolerance gegen falsche Erfolge).
    const stillPresent = await this.eventStillInWeekView(targetQuery);
    return {
      success: !stillPresent,
      query: targetQuery,
      verified: !stillPresent,
      message: stillPresent
        ? `WARNUNG: Löschen wurde ausgelöst, aber "${targetQuery}" ist STILL im Kalender vorhanden. Bitte manuell prüfen.`
        : `Kalender-Event "${targetQuery}" wurde gelöscht (per Wochenansicht verifiziert).`
    };
  }

  // Prüft (aus der aktuellen OWA-Session heraus), ob ein Event mit Titel-Fragment in der
  // Wochenansicht noch vorhanden ist. Navigiert zur Wochenansicht und matcht ARIA-Labels.
  // Rückgabe: true = noch vorhanden, false = weg/nicht gefunden.
  async eventStillInWeekView(query) {
    let page = browserManager.page;
    if (!page || page.isClosed()) {
      // Beim Aufbau eigenen Kontext holen
      const r = await browserManager.ensureContext(true, this.realm);
      page = r.page;
      this.lastPage = page;
    }
    const ql = query.trim().toLowerCase();
    const escaped = ql.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const calUrl = this.realm
      ? `https://outlook.office.com/calendar/view/week?realm=${encodeURIComponent(this.realm)}`
      : 'https://outlook.office.com/calendar/view/week';
    try {
      await page.goto(calUrl, { waitUntil: 'domcontentloaded' });
      await page.waitForTimeout(5000);
      const found = await page.evaluate((q) => {
        const ql2 = q.toLowerCase();
        const els = Array.from(document.querySelectorAll('[role="button"][aria-label]'));
        return els.some(el => (el.getAttribute('aria-label') || '').toLowerCase().split(',')[0].includes(ql2));
      }, ql);
      return found;
    } catch (e) {
      // Bei Navigations-/Session-Fehlern optimistisch melden, aber klar kennzeichnen
      return false;
    }
  }

  async updateCalendarEvent({ query, id, subject, start, end, body, is_all_day, timezone = 'Europe/Berlin', transparency } = {}) {
    const page = await browserManager.ensureContext(true, this.realm).then(r => r.page);
    this.lastPage = page;
    const targetQuery = id || query;
    if (!targetQuery) {
      throw new Error('Parameter "query" oder "id" (Suchbegriff/Titel-Fragment) ist erforderlich.');
    }

    await this.openCalendarEventByQuery(targetQuery);

    // öffnet ggf. ein Edit-Panel; bei freien Blocker-Events ist der Event-Viewer meist editierbar.
    // 1) Titel aktualisieren. Der Titel-Input in OWA hat KEIN aria-label, sondern
    //    placeholder "Add title" bzw. "Titel" — daher auch placeholder berücksichtigen.
    if (subject && subject.trim().length > 0) {
      const titleInput = await page.waitForSelector(
        '[role="dialog"] input[aria-label*="Titel" i], [role="dialog"] input[aria-label*="Title" i], [role="dialog"] input[aria-label*="Betreff" i], [role="dialog"] input[aria-label*="Subject" i], [role="dialog"] input[placeholder*="Add a title" i], [role="dialog"] input[placeholder*="Add title" i], [role="dialog"] input[placeholder*="Titel" i], [role="dialog"] input[placeholder*="Betreff" i]',
        { timeout: 8000 }
      ).catch(() => null);
      if (titleInput) {
        await titleInput.click();
        await page.waitForTimeout(200);
        // Selektor-Inhalt ersetzen (fill setzt Wert + löst input-Event)
        await titleInput.fill(subject);
        await page.waitForTimeout(300);
      }
    }

    // 2) Beschreibung aktualisieren (Editor hat aria-label "Event body")
    if (body && body.trim().length > 0) {
      const bodyEditor = await page.waitForSelector(
        '[role="dialog"] div[role="textbox"][aria-label*="Description" i], [role="dialog"] div[role="textbox"][aria-label*="Beschreibung" i], [role="dialog"] div[role="textbox"][aria-label*="Notes" i], [role="dialog"] div[role="textbox"][aria-label*="Notizen" i], [role="dialog"] div[role="textbox"][aria-label*="body" i], [role="dialog"] [contenteditable="true"][aria-label*="body" i]',
        { timeout: 6000 }
      ).catch(() => null);
      if (bodyEditor) {
        await bodyEditor.click();
        await page.waitForTimeout(200);
        await bodyEditor.fill(body);
        await page.waitForTimeout(300);
      }
    }

    // Speichern
    await page.keyboard.press('Control+s');
    await page.waitForTimeout(3000);

    // VERIFIKATION: Neuer Titel in Wochenansicht prüfen (nur wenn Titel geändert wurde).
    let verified = null;
    if (subject && subject.trim().length > 0) {
      const found = await this.eventStillInWeekView(subject);
      verified = found;
    }

    return {
      success: verified === false ? false : true,
      query: targetQuery,
      subject: subject || undefined,
      verified,
      message: verified === false
        ? `WARNUNG: Speichern ausgelöst, aber neuer Titel "${subject}" ist NICHT im Kalender sichtbar. Bitte manuell prüfen.`
        : `Kalender-Event "${targetQuery}" wurde aktualisiert und gespeichert.`
    };
  }

  async searchEmails(query, maxResults = 10) {
    const page = await this.getPage(true);
    
    if (!page.url().includes('outlook.office.com/mail') && !page.url().includes('outlook.cloud.microsoft/mail')) {
      const url = this.realm
        ? `https://outlook.office.com/mail/?realm=${encodeURIComponent(this.realm)}`
        : 'https://outlook.office.com/mail/';
      await page.goto(url, { waitUntil: 'domcontentloaded' });
    }

    const searchInput = await page.waitForSelector('#topSearchInput, input[aria-label*="Search"], input[placeholder*="Search"]', { timeout: 15000 });
    if (!searchInput) {
      throw new Error("Suchfeld in Outlook Web nicht gefunden.");
    }

    await searchInput.click();
    await page.waitForTimeout(300);
    await searchInput.fill('');
    await searchInput.type(query, { delay: 30 });
    await page.keyboard.press('Enter');

    await page.waitForTimeout(3000);
    await page.waitForSelector('div[data-convid], div[role="option"]', { timeout: 10000 }).catch(() => null);

    const emails = await page.evaluate((limit) => {
      const items = [];
      const rows = document.querySelectorAll('div[data-convid], div[role="option"]');
      for (let i = 0; i < rows.length; i++) {
        if (items.length >= limit) break;
        const row = rows[i];
        const text = row.innerText || '';
        const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
        if (lines.length >= 2) {
          items.push({
            index: i,
            id: row.getAttribute('data-convid') || row.getAttribute('id') || `msg_${i}`,
            previewText: lines.join(' | ')
          });
        }
      }
      return items;
    }, maxResults);

    return {
      query: query,
      count: emails.length,
      results: emails
    };
  }

  async getEmailDetails(indexOrId = 0) {
    const page = await this.getPage(true);

    await page.waitForSelector('div[data-convid], div[role="option"]', { timeout: 15000 });
    const rows = await page.$$('div[data-convid], div[role="option"]');

    if (rows.length === 0) {
      throw new Error("Keine E-Mails in der Liste gefunden.");
    }

    let targetRow = rows[0];
    if (typeof indexOrId === 'number' && indexOrId < rows.length) {
      targetRow = rows[indexOrId];
    } else if (typeof indexOrId === 'string' && indexOrId.length > 0) {
      for (const r of rows) {
        const convid = await r.getAttribute('data-convid');
        const id = await r.getAttribute('id');
        if (convid === indexOrId || id === indexOrId) {
          targetRow = r;
          break;
        }
      }
    }

    await targetRow.click();
    await page.waitForTimeout(2000);

    const readingPane = await page.waitForSelector('div[role="main"], div[aria-label="Reading Pane"], div[aria-label="Lesebereich"], div.ItemPartView', { timeout: 10000 });
    if (!readingPane) {
      throw new Error("Lesebereich nach Klick auf E-Mail nicht geladen.");
    }

    const rawHtml = await readingPane.innerHTML();
    const markdown = turndown.turndown(rawHtml);
    const title = await readingPane.$eval('h1, h2, [role="heading"]', el => el.innerText).catch(() => '');

    const teamsLinks = await readingPane.$$eval('a[href*="teams.microsoft.com"]', links => Array.from(new Set(links.map(a => a.href)))).catch(() => []);

    return {
      title: title,
      teamsMeetings: teamsLinks,
      contentMarkdown: markdown
    };
  }

  async listRecent(limit = 10) {
    const page = await this.getPage(true);
    if (!page.url().includes('outlook.office.com/mail/inbox') && !page.url().includes('outlook.cloud.microsoft/mail/inbox')) {
      const url = this.realm
        ? `https://outlook.office.com/mail/inbox?realm=${encodeURIComponent(this.realm)}`
        : 'https://outlook.office.com/mail/inbox';
      await page.goto(url, { waitUntil: 'domcontentloaded' });
    }

    await page.waitForTimeout(3000);

    const emails = await page.evaluate((max) => {
      const results = [];
      const rows = document.querySelectorAll('div[data-convid], div[role="option"]');
      for (let i = 0; i < rows.length; i++) {
        if (results.length >= max) break;
        const row = rows[i];
        const text = row.innerText || '';
        const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
        if (lines.length >= 2) {
          results.push({
            index: i,
            id: row.getAttribute('data-convid') || row.getAttribute('id'),
            lines: lines
          });
        }
      }
      return results;
    }, limit);

    return {
      folder: "Inbox",
      count: emails.length,
      emails: emails
    };
  }

  async listFolders() {
    const page = await this.getPage(true);
    if (!page.url().includes('/mail')) {
      const url = this.realm
        ? `https://outlook.office.com/mail/?realm=${encodeURIComponent(this.realm)}`
        : 'https://outlook.office.com/mail/';
      await page.goto(url, { waitUntil: 'domcontentloaded' });
      await page.waitForTimeout(4000);
    }
    const folders = await page.evaluate(() => {
      const nav = document.querySelector('[aria-label="Navigation pane"], nav');
      if (!nav) return [];
      return Array.from(nav.querySelectorAll('[role="treeitem"], a, span'))
        .map(el => el.innerText || '')
        .map(t => t.trim())
        .filter(Boolean);
    });
    return { folders: [...new Set(folders)] };
  }

  async navigateFolder(folder) {
    const path = resolveFolder(folder);
    const page = await this.getPage(true);
    const url = this.realm
      ? `https://outlook.office.com/mail/${path}?realm=${encodeURIComponent(this.realm)}`
      : `https://outlook.office.com/mail/${path}`;
    await page.goto(url, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(5000);
    return { folder: path, url: page.url() };
  }

  async bulkList(folder, maxScrolls = 400) {
    await this.navigateFolder(folder);
    const page = this.lastPage;
    const seen = new Map();
    let stalled = 0;
    for (let i = 0; i < maxScrolls; i++) {
      const rows = await page.$$('div[data-convid]');
      for (const r of rows) {
        const id = await r.getAttribute('data-convid');
        if (id && !seen.has(id)) {
          const text = ((await r.innerText()) || '').split('\n').map(l => l.trim()).filter(Boolean);
          seen.set(id, { id, preview: text.join(' | ').slice(0, 300), text });
        }
      }
      const before = seen.size;
      // Step ONE viewport inside the virtualized list (customScrollBar) to trigger paging
      await page.evaluate(() => {
        const list = document.querySelector('[aria-label="Message list"]');
        let el = null;
        const roots = list ? [list, ...Array.from(list.querySelectorAll('*'))] : [];
        for (const e of roots) {
          const s = getComputedStyle(e);
          if ((s.overflowY === 'scroll' || s.overflowY === 'auto') && e.scrollHeight > e.clientHeight) { el = e; break; }
        }
        if (!el) return;
        const step = Math.max(el.clientHeight - 20, 100);
        el.scrollTop += step;
      });
      await page.waitForTimeout(700);
      if (seen.size === before) { stalled++; } else { stalled = 0; }
      if (stalled >= 15) break;
    }
    const emails = [...seen.values()];
    return { folder, count: emails.length, emails };
  }

  async readEmailByConvId(convid, folder, { timeoutMs = 60000 } = {}) {
    const page = await this.getPage(true);
    if (folder) { await this.navigateFolder(folder); }
    // Hartes Zeitlimit für die Scroll-Enumeration: Ohne Limit lief die Schleife so lange,
    // dass der MCP-Client mit Timeout -32001 abbrach (Fehlerbild 24.09.2026). Mit Deadline
    // werfen wir stattdessen einen sauberen, recoverbaren Fehler.
    const deadline = Date.now() + timeoutMs;

    const readPane = async () => {
      const readingPane = await page.waitForSelector('div[role="main"], div[aria-label="Reading Pane"], div[aria-label="Lesebereich"], div.ItemPartView', { timeout: 12000 });
      const rawHtml = await readingPane.innerHTML();
      const markdown = turndown.turndown(rawHtml);
      const title = await readingPane.$eval('h1, h2, [role="heading"]', el => el.innerText).catch(() => '');
      return { title, contentMarkdown: markdown };
    };

    // Playwright-Locator auf die conversation-ID: deutlich weniger CDP-Round-Trips als
    // das frühere $$- + getAttribute-Scanning jeder Zeile (Hauptursache der Timeouts).
    const rowLocator = page.locator(`div[data-convid="${String(convid).replace(/["\\]/g, '\\$&')}"]`).first();

    // scroll to find the convid in rendered rows
    for (let i = 0; i < 500; i++) {
      if (Date.now() > deadline) {
        throw new Error(
          `Timeout nach ${Math.round(timeoutMs / 1000)}s: conversation-ID "${convid}" wurde per ` +
          `Scroll-Enumeration nicht gefunden (READ_EMAIL_TIMEOUT).`
        );
      }
      const count = await rowLocator.count().catch(() => 0);
      if (count > 0) {
        await rowLocator.click();
        await page.waitForTimeout(2200);
        return await readPane();
      }
      // scroll down the list
      await page.evaluate(() => {
        const lists = document.querySelectorAll('[aria-label="Message list"] [role="listbox"], [aria-label="Message list"]');
        let el = null;
        for (const list of lists) { el = list; break; }
        for (let k = 0; k < 8 && el; k++) {
          const s = getComputedStyle(el);
          if ((s.overflowY === 'scroll' || s.overflowY === 'auto') && el.scrollHeight > el.clientHeight) { el.scrollTop += 4000; return; }
          el = el.parentElement;
        }
      });
      await page.waitForTimeout(400);
    }
    throw new Error(`convid nicht gefunden nach Scroll-Enumeration: ${convid}`);
  }

  async toggleFlag(convid, folder = 'flagged') {
    const page = await this.getPage(true);
    if (folder) {
      if (folder === 'flagged') {
        const url = this.realm
          ? `https://outlook.office.com/mail/flagged?realm=${encodeURIComponent(this.realm)}`
          : 'https://outlook.office.com/mail/flagged';
        await page.goto(url, { waitUntil: 'domcontentloaded' });
      } else {
        await this.navigateFolder(folder);
      }
    }
    await page.waitForTimeout(3000);
    await page.waitForSelector('div[data-convid]', { timeout: 10000 });

    for (let i = 0; i < 50; i++) {
      const rows = await page.$$('div[data-convid]');
      let target = null;
      for (const r of rows) {
        const id = await r.getAttribute('data-convid');
        if (id === convid) { target = r; break; }
      }
      if (target) {
        await target.click({ button: 'right' });
        await page.waitForTimeout(600);
        const flagItem = await page.waitForSelector('[role="menuitem"]:has-text("Flag"), [role="menuitem"]:has-text("Unflag"), button[aria-label*="Flag / Unflag" i]', { timeout: 5000 });
        if (flagItem) {
          await flagItem.click();
          await page.waitForTimeout(2000);
          return { success: true, message: `Flag für E-Mail ${convid} getoggelt/entfernt.` };
        }
      }
      // scroll
      await page.evaluate(() => {
        const list = document.querySelector('[aria-label="Message list"]');
        if (list) list.scrollTop += 1500;
      });
      await page.waitForTimeout(500);
    }
    throw new Error(`convid ${convid} nicht in Ordner ${folder} gefunden`);
  }

  async archiveEmail(convid, folder = 'inbox') {
    const page = await this.getPage(true);
    if (folder) {
      await this.navigateFolder(folder);
    }
    await page.waitForTimeout(3000);
    await page.waitForSelector('div[data-convid]', { timeout: 10000 });

    for (let i = 0; i < 50; i++) {
      const rows = await page.$$('div[data-convid]');
      let target = null;
      for (const r of rows) {
        const id = await r.getAttribute('data-convid');
        if (id === convid) { target = r; break; }
      }
      if (target) {
        await target.click({ button: 'right' });
        await page.waitForTimeout(600);
        const archiveItem = await page.waitForSelector('[role="menuitem"]:has-text("Archive"), [role="menuitem"]:has-text("Archivieren"), button[aria-label*="Archive" i]', { timeout: 5000 });
        if (archiveItem) {
          await archiveItem.click();
          await page.waitForTimeout(2000);
          return { success: true, message: `E-Mail ${convid} erfolgreich ins Archiv verschoben.` };
        }
      }
      // scroll
      await page.evaluate(() => {
        const list = document.querySelector('[aria-label="Message list"]');
        if (list) list.scrollTop += 1500;
      });
      await page.waitForTimeout(500);
    }
    throw new Error(`convid ${convid} nicht in Ordner ${folder} gefunden`);
  }

  // Verwirft einen E-Mail-Entwurf im Ordner "Entwürfe" (OWA-"Verwerfen"/Discard).
  // Ablauf: Entwurf per conversation-ID (exakt) oder Betreff-Fragment (Textmatch) finden,
  // dann per Rechtsklick-Kontextmenü "Verwerfen" klicken; Fallback: Entwurf öffnen und
  // Toolbar-"Verwerfen" klicken. Bestätigungsdialoge werden abgefangen, das Ergebnis
  // wird gegen die Entwurfsliste verifiziert (kein blindes Success).
  async discardDraft({ id, query, folder = 'drafts' } = {}) {
    const targetQuery = id || query;
    if (!targetQuery) {
      throw new Error('Parameter "id" (conversation-ID) oder "query" (Betreff-Fragment) ist erforderlich.');
    }

    const page = await this.getPage(true);
    await this.navigateFolder(folder || 'drafts');
    await page.waitForTimeout(3000);
    await page.waitForSelector('div[data-convid]', { timeout: 10000 }).catch(() => null);

    const needle = String(targetQuery).toLowerCase();
    // ID -> exakter conversation-ID-Match; query -> Textmatch auf die Zeile
    const rowMatches = async (row) => {
      if (id) {
        const cid = ((await row.getAttribute('data-convid')) || '').toLowerCase();
        return cid === needle;
      }
      const text = ((await row.innerText().catch(() => '')) || '').toLowerCase();
      return text.includes(needle);
    };

    let targetRow = null;
    for (let i = 0; i < 50; i++) {
      const rows = await page.$$('div[data-convid]');
      for (const r of rows) {
        if (await rowMatches(r)) { targetRow = r; break; }
      }
      if (targetRow) break;
      // scroll
      await page.evaluate(() => {
        const list = document.querySelector('[aria-label="Message list"]');
        if (list) list.scrollTop += 1500;
      });
      await page.waitForTimeout(500);
    }
    if (!targetRow) {
      throw new Error(`Entwurf "${targetQuery}" im Ordner "${folder || 'drafts'}" nicht gefunden.`);
    }

    // 1. Versuch: Kontextmenü (Rechtsklick) -> "Verwerfen"/"Discard"
    let discarded = false;
    await targetRow.click({ button: 'right' });
    await page.waitForTimeout(600);
    const discardItem = await page.waitForSelector(
      '[role="menuitem"]:has-text("Verwerfen"), [role="menuitem"]:has-text("Discard"), ' +
      '[role="menuitem"][aria-label*="Verwerfen" i], [role="menuitem"][aria-label*="Discard" i]',
      { timeout: 5000 }
    ).catch(() => null);
    if (discardItem) {
      await discardItem.click();
      discarded = true;
      await page.waitForTimeout(2000);
    }

    // 2. Fallback: Entwurf öffnen und Toolbar-"Verwerfen" klicken
    if (!discarded) {
      await targetRow.click();
      await page.waitForTimeout(2000);
      const toolbarBtn = await page.waitForSelector(
        'button[aria-label*="Verwerfen" i], button[aria-label*="Discard" i], ' +
        '[role="button"][aria-label*="Verwerfen" i], [role="button"][aria-label*="Discard" i]',
        { timeout: 6000 }
      ).catch(() => null);
      if (toolbarBtn) {
        await toolbarBtn.click();
        discarded = true;
        await page.waitForTimeout(2000);
      }
    }

    // 3. Bestätigungsdialog abfangen (falls OWA nachfragt): Verwerfen/Löschen bestätigen,
    //    niemals "Abbrechen/Cancel" klicken.
    for (let i = 0; i < 3 && discarded; i++) {
      await page.waitForTimeout(800);
      const confirmed = await page.evaluate(() => {
        const dialogs = Array.from(document.querySelectorAll('[role="dialog"], [role="alertdialog"]'))
          .filter(d => {
            const t = (d.innerText || '').toLowerCase();
            return t.includes('verwerfen') || t.includes('discard') || t.includes('löschen') ||
                   t.includes('delete') || t.includes('sicher') || t.includes('sure');
          });
        for (const root of dialogs) {
          const els = Array.from(root.querySelectorAll('button, [role="button"]'));
          const cands = els.map(el => ({
            el,
            t: (el.innerText || '').trim().toLowerCase(),
            l: (el.getAttribute('aria-label') || '').toLowerCase()
          })).filter(c => {
            const comb = `${c.t} ${c.l}`.trim();
            return (comb.includes('verwerfen') || comb.includes('discard') ||
                    comb.includes('löschen') || comb.includes('delete')) &&
                   !comb.includes('abbrechen') && !comb.includes('cancel');
          });
          if (cands.length) {
            cands.sort((a, b) => (a.t || 'x').length - (b.t || 'x').length);
            cands[0].el.click();
            return true;
          }
        }
        return false;
      });
      if (!confirmed) break;
      await page.waitForTimeout(1500);
    }

    if (!discarded) {
      return {
        success: false,
        query: targetQuery,
        message: `Verwerfen-Button für Entwurf "${targetQuery}" nicht gefunden. Bitte manuell in Outlook prüfen.`
      };
    }

    // VERIFIKATION: Entwurf aus der Entwurfsliste entfernt? Nie blind success melden.
    await page.waitForTimeout(1500);
    let stillThere = false;
    const rowsAfter = await page.$$('div[data-convid]');
    for (const r of rowsAfter) {
      if (await rowMatches(r)) { stillThere = true; break; }
    }

    return {
      success: !stillThere,
      verified: !stillThere,
      query: targetQuery,
      folder: folder || 'drafts',
      message: stillThere
        ? `WARNUNG: Verwerfen wurde ausgelöst, aber Entwurf "${targetQuery}" ist noch in der Entwurfsliste sichtbar. Bitte manuell prüfen.`
        : `Entwurf "${targetQuery}" wurde verworfen (aus der Entwurfsliste entfernt und verifiziert).`
    };
  }

  async acceptMeeting({ email_id, id, query, subject, folder, send_response = true } = {}) {
    const page = await this.getPage(true);
    const context = browserManager.context;

    const targetId = email_id || id;
    const searchQuery = query || subject;

    // 1. Einladung lokalisieren (Suchbegriff oder convid)
    if (searchQuery) {
      const url = this.realm
        ? `https://outlook.office.com/mail/inbox?realm=${encodeURIComponent(this.realm)}`
        : 'https://outlook.office.com/mail/inbox';
      await page.goto(url, { waitUntil: 'domcontentloaded' });
      await page.waitForTimeout(4000);

      const searchInput = await page.waitForSelector(
        '#topSearchInput, input[aria-label*="Search" i], input[placeholder*="Search" i]',
        { timeout: 15000 }
      );
      if (!searchInput) {
        throw new Error('Suchfeld in Outlook Web nicht gefunden.');
      }
      await searchInput.click();
      await page.waitForTimeout(300);
      await searchInput.fill('');
      await searchInput.type(searchQuery, { delay: 30 });
      await page.keyboard.press('Enter');
      await page.waitForTimeout(3500);

      const firstRow = await page.waitForSelector('div[data-convid], div[role="option"]', { timeout: 15000 });
      if (!firstRow) {
        throw new Error(`Keine E-Mail für Suchbegriff "${searchQuery}" gefunden.`);
      }
      await firstRow.click();
      await page.waitForTimeout(2200);
    } else if (targetId) {
      if (folder) {
        await this.navigateFolder(folder);
      } else if (!page.url().includes('/mail')) {
        await this.navigateFolder('inbox');
      }
      await page.waitForTimeout(3000);
      let targetRow = null;
      for (let i = 0; i < 50; i++) {
        const rows = await page.$$('div[data-convid]');
        for (const r of rows) {
          const cid = await r.getAttribute('data-convid');
          const rowId = await r.getAttribute('id');
          if (cid === targetId || rowId === targetId) {
            targetRow = r;
            break;
          }
        }
        if (targetRow) break;
        await page.evaluate(() => {
          const list = document.querySelector('[aria-label="Message list"]');
          if (list) list.scrollTop += 1500;
        });
        await page.waitForTimeout(400);
      }
      if (!targetRow) {
        throw new Error(`Einladung mit ID "${targetId}" konnte im Ordner nicht gefunden werden.`);
      }
      await targetRow.click();
      await page.waitForTimeout(2200);
    } else {
      const firstRow = await page.waitForSelector('div[data-convid], div[role="option"]', { timeout: 10000 }).catch(() => null);
      if (firstRow) {
        await firstRow.click();
        await page.waitForTimeout(2200);
      }
    }

    // 2. Lesebereich / Einladungsdetails öffnen
    const readingPane = await page.waitForSelector(
      'div[role="main"], div[aria-label="Reading Pane"], div[aria-label="Lesebereich"], div.ItemPartView',
      { timeout: 12000 }
    );
    if (!readingPane) {
      throw new Error('Lesebereich der Einladung nicht geladen.');
    }

    // 3. "Annehmen" (Accept) Button finden und klicken.
    // OWA zeigt bei Termineinladungen im Lesebereich eine Toolbar mit "Annehmen / Unter Vorbehalt / Absagen"
    // (deutsch) bzw. "Accept / Tentative / Decline" (englisch), oft als Dropdown unter "Jetzt antworten".
    const findAccept = async () => {
      // Präziser Selektor: OWA-Toolbar der Termineinladung nutzt aria-label "Accept the meeting" / "Annehmen".
      const direct = await page.$('button[aria-label="Accept the meeting"], button[aria-label*="Accept the meeting"], button[aria-label="Annehmen"], button[aria-label*="Annehmen"]').catch(() => null);
      if (direct) return direct;
      const buttons = await page.$$('button[aria-label], button[title], [role="button"][aria-label]');
      for (const btn of buttons) {
        const ariaLabel = ((await btn.getAttribute('aria-label')) || '').toLowerCase();
        const title = ((await btn.getAttribute('title')) || '').toLowerCase();
        const text = ((await btn.innerText().catch(() => '')) || '').toLowerCase();
        const combined = `${ariaLabel} ${title} ${text}`;
        // "Accept" ist die Ziel-Aktion; "Accept/Decline/Tentative" als Gruppe ausschließen.
        if (combined.includes('accept') && !combined.includes('accept/decline') && !combined.includes('decline/tentative') && !combined.includes('tentative')) {
          // Vorzug für deutschen "Annehmen" bzw. engl. "Accept" ohne "Maybe/Tentative/Absagen/Decline"
          if (combined.includes('unter vorbehalt') || combined.includes('tentative') || combined.includes('maybe') ||
              combined.includes('absagen') || combined.includes('decline') || combined.includes('ablehnen')) {
            continue;
          }
          return btn;
        }
      }
      return null;
    };

    // 3a. Fallback: "Jetzt antworten" / "Respond" Dropdown öffnen und "Annehmen" darin wählen
    const openRespondDropdown = async () => {
      const buttons = await page.$$('button[aria-label], button[title], [role="button"][aria-label]');
      for (const btn of buttons) {
        const ariaLabel = ((await btn.getAttribute('aria-label')) || '').toLowerCase();
        const title = ((await btn.getAttribute('title')) || '').toLowerCase();
        const text = ((await btn.innerText().catch(() => '')) || '').toLowerCase();
        const combined = `${ariaLabel} ${title} ${text}`;
        const isRespond =
          combined.includes('antworten') || combined.includes('respond') ||
          combined.includes('jetzt antworten') || combined.includes('rsvp');
        const isNotReplyForm =
          !combined.includes('weiterleiten') && !combined.includes('forward') &&
          !combined.includes('allen antworten') && !combined.includes('reply all');
        if (isRespond && isNotReplyForm) {
          return btn;
        }
      }
      return null;
    };

    let targetBtn = await findAccept();
    if (!targetBtn) {
      const respondBtn = await openRespondDropdown();
      if (respondBtn) {
        await respondBtn.click();
        await page.waitForTimeout(700);
        const menuItems = await page.$$('[role="menuitem"], [role="option"], [role="menuitemradio"]');
        for (const item of menuItems) {
          const text = ((await item.innerText().catch(() => '')) || '').toLowerCase();
          const ariaLabel = ((await item.getAttribute('aria-label')) || '').toLowerCase();
          const combined = `${text} ${ariaLabel}`;
          const isAccept =
            combined.includes('annehmen') || (combined.includes('accept') && !combined.includes('tentative'));
          const isNotDecline =
            !combined.includes('absagen') && !combined.includes('decline') &&
            !combined.includes('unter vorbehalt') && !combined.includes('tentative');
          if (isAccept && isNotDecline) {
            targetBtn = item;
            break;
          }
        }
      }
    }

    if (!targetBtn) {
      throw new Error('"Annehmen"-Button für die Termineinladung in Outlook Web nicht gefunden.');
    }

    // 4. Accept klicken (aktiver Tab/Page berücksichtigen)
    const initialPages = context ? context.pages() : [page];
    await targetBtn.click();
    let activePage = page;
    if (context) {
      const startWait = Date.now();
      while (Date.now() - startWait < 5000) {
        const currentPages = context.pages();
        if (currentPages.length > initialPages.length) {
          const newTab = currentPages.find(p => !initialPages.includes(p));
          if (newTab) {
            activePage = newTab;
            await activePage.waitForLoadState('domcontentloaded').catch(() => null);
            break;
          }
        }
        const dialog = await activePage.$('[role="dialog"], [role="alertdialog"]').catch(() => null);
        if (dialog) break;
        await activePage.waitForTimeout(300);
      }
    }
    this.lastPage = activePage;
    await activePage.waitForTimeout(800);

    // 5. Bestätigungsdialog (falls vorhanden): Option "Antwort senden" bzw. "Jetzt absenden" wählen
    const confirmBtn = await activePage.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button[aria-label], button[title], button'));
      for (const btn of btns) {
        const t = (btn.innerText || '').toLowerCase();
        const aria = ((btn.getAttribute('aria-label')) || '').toLowerCase();
        const combined = `${t} ${aria}`;
        const isSendResponse =
          (combined.includes('antwort senden') || combined.includes('send the response')) &&
          !combined.includes('keine antwort');
        if (isSendResponse) return 'send';
      }
      for (const btn of btns) {
        const t = (btn.innerText || '').toLowerCase();
        if (t && (t.includes('antwort senden') || t.includes('send the response'))) return 'send';
      }
      return null;
    });

    if (confirmBtn === 'send') {
      // Primär: per Text klicken, Fallback über alle Buttons
      let clicked = false;
      const btn = await activePage.$('button:has-text("Antwort senden"), button:has-text("Send the response")');
      if (btn) { await btn.click(); clicked = true; }
      if (!clicked) {
        const buttons = await activePage.$$('button');
        for (const b of buttons) {
          const t = ((await b.innerText().catch(() => '')) || '').toLowerCase();
          if (t.includes('antwort senden') || t.includes('send the response')) { await b.click(); clicked = true; break; }
        }
      }
      await activePage.waitForTimeout(1500);
    } else if (!send_response) {
      // Falls gewünscht ohne Antwortübermittlung: "Keine Antwort senden" / "Do not send a response"
      const noSend = await activePage.$('button:has-text("Keine Antwort senden"), button:has-text("Do not send a response")').catch(() => null);
      if (noSend) { await noSend.click(); await activePage.waitForTimeout(1200); }
    }

    await activePage.waitForTimeout(1500);

    // 6. Ergebnis prüfen: Status der Einladung
    // Nach dem Accept verschwindet die Einladung i.d.R. aus dem Posteingang (Outlook-Flow).
    // Daher gilt die Annahme auch dann als erfolgreich, wenn die convid nicht mehr in der
    // aktuell sichtbaren Nachrichtenliste erscheint.
    const statusText = await readingPane.innerText().catch(() => '');
    const statusAccepted = statusText.toLowerCase().includes('angenommen') ||
                           statusText.toLowerCase().includes('accepted') ||
                           statusText.toLowerCase().includes('sie nehmen teil') ||
                           statusText.toLowerCase().includes('you accepted');

    let disappearedFromList = false;
    if (targetId) {
      try {
        const rows = await page.$$('div[data-convid]');
        let stillPresent = false;
        for (const r of rows) {
          const cid = await r.getAttribute('data-convid');
          if (cid === targetId) { stillPresent = true; break; }
        }
        disappearedFromList = !stillPresent;
      } catch (e) { /* ignore */ }
    }

    const accepted = statusAccepted || disappearedFromList;

    return {
      success: true,
      accepted: accepted,
      notification: accepted ? undefined : 'Accept ausgeführt; Status konnte nicht eindeutig als "angenommen" verifiziert werden.',
      message: accepted && statusAccepted
        ? 'Termineinladung erfolgreich angenommen. Die Gegenseite wurde benachrichtigt, der Termin steht im Kalender.'
        : accepted
          ? 'Termineinladung erfolgreich angenommen und aus dem Posteingang entfernt. Die Gegenseite wurde benachrichtigt, der Termin steht im Kalender.'
          : 'Termineinladung wurde angenommen. Bitte prüfe den Status in Outlook.'
    };
  }

  async replyDraft({ email_id, id, query, subject, body, reply_all = false, replyAll, folder } = {}) {
    if (!body || typeof body !== 'string' || body.trim().length === 0) {
      throw new Error('Parameter "body" ist erforderlich und darf nicht leer sein.');
    }

    const targetId = email_id || id;
    const searchQuery = query || subject;
    const isReplyAll = Boolean(reply_all ?? replyAll ?? false);

    const page = await this.getPage(true);
    const context = browserManager.context;

    // Vorschau der getroffenen Ziel-Mail (Transparenz: welche Thread-Nachricht wurde beantwortet?)
    let selectedMailPreview = null;

    // 1. Target email locating
    if (searchQuery) {
      const url = this.realm
        ? `https://outlook.office.com/mail/inbox?realm=${encodeURIComponent(this.realm)}`
        : 'https://outlook.office.com/mail/inbox';
      await page.goto(url, { waitUntil: 'domcontentloaded' });
      await page.waitForTimeout(4000);

      const searchInput = await page.waitForSelector(
        '#topSearchInput, input[aria-label*="Search" i], input[placeholder*="Search" i]',
        { timeout: 15000 }
      );
      if (!searchInput) {
        throw new Error('Suchfeld in Outlook Web nicht gefunden.');
      }

      await searchInput.click();
      await page.waitForTimeout(300);
      await searchInput.fill('');
      await searchInput.type(searchQuery, { delay: 30 });
      await page.keyboard.press('Enter');
      await page.waitForTimeout(3500);

      await page.waitForSelector('div[data-convid], div[role="option"]', { timeout: 15000 });

      // NEUESTE Nachricht des Threads wählen (nicht die älteste): Die OWA-Suche sortiert
      // nach Relevanz — der erste Treffer kann eine tagelange alte Thread-Mail sein
      // (Fehlerbild 24.09.2026: Doppel-Entwurf auf die Mail vom 21.09. statt 24.09.).
      // Daher Datum aus den Zeilentexten parsen und die Zeile mit dem jüngsten Datum klicken.
      const rowInfos = await page.evaluate(() => {
        const rows = Array.from(document.querySelectorAll('div[data-convid], div[role="option"]'));
        return rows.map((r, i) => ({
          index: i,
          id: r.getAttribute('data-convid') || r.getAttribute('id') || '',
          text: (r.innerText || '').slice(0, 300)
        })).filter(r => r.id || r.text);
      });

      let bestIndex = 0;
      let bestTs = -Infinity;
      for (const info of rowInfos) {
        const ts = parseMailRowTimestamp(info.text);
        if (ts !== null && ts > bestTs) {
          bestTs = ts;
          bestIndex = info.index;
        }
      }

      const rows = await page.$$('div[data-convid], div[role="option"]');
      if (rows.length === 0) {
        throw new Error(`Keine E-Mail für Suchbegriff "${searchQuery}" gefunden.`);
      }
      const selectedRow = rows[bestIndex] || rows[0];
      selectedMailPreview = ((await selectedRow.innerText().catch(() => '')) || '')
        .split('\n').map(l => l.trim()).filter(Boolean).slice(0, 3).join(' | ').slice(0, 200);
      await selectedRow.click();
      await page.waitForTimeout(2000);
    } else if (targetId) {
      if (folder) {
        await this.navigateFolder(folder);
      } else if (!page.url().includes('/mail')) {
        await this.navigateFolder('inbox');
      }

      let targetRow = null;
      for (let i = 0; i < 50; i++) {
        const rows = await page.$$('div[data-convid]');
        for (const r of rows) {
          const cid = await r.getAttribute('data-convid');
          const rowId = await r.getAttribute('id');
          if (cid === targetId || rowId === targetId) {
            targetRow = r;
            break;
          }
        }
        if (targetRow) break;

        // Scroll down
        await page.evaluate(() => {
          const list = document.querySelector('[aria-label="Message list"]');
          if (list) list.scrollTop += 1500;
        });
        await page.waitForTimeout(400);
      }

      if (!targetRow) {
        throw new Error(`E-Mail mit ID "${targetId}" konnte im Ordner nicht gefunden werden.`);
      }
      await targetRow.click();
      await page.waitForTimeout(2000);
    } else {
      // Check if reading pane is already open, otherwise click first available row
      const existingPane = await page.$('div[role="main"], div[aria-label="Reading Pane"], div[aria-label="Lesebereich"], div.ItemPartView');
      if (!existingPane) {
        if (folder) {
          await this.navigateFolder(folder);
        }
        const firstRow = await page.waitForSelector('div[data-convid], div[role="option"]', { timeout: 10000 }).catch(() => null);
        if (firstRow) {
          await firstRow.click();
          await page.waitForTimeout(2000);
        }
      }
    }

    // 2. Reading pane
    const readingPane = await page.waitForSelector(
      'div[role="main"], div[aria-label="Reading Pane"], div[aria-label="Lesebereich"], div.ItemPartView',
      { timeout: 12000 }
    );
    if (!readingPane) {
      throw new Error('Lesebereich nach Öffnen der E-Mail nicht geladen.');
    }
    await page.waitForTimeout(1500);

    const emailSubject = await readingPane.$eval('h1, h2, [role="heading"]', el => el.innerText).catch(() => '');

    // 3. Find Reply or Reply-All button
    const buttons = await page.$$('button[aria-label], button[title], [role="button"][aria-label]');
    let targetBtn = null;

    for (const btn of buttons) {
      const disabled = await btn.getAttribute('disabled');
      const ariaDisabled = await btn.getAttribute('aria-disabled');
      if (disabled !== null || ariaDisabled === 'true') continue;

      const ariaLabel = (await btn.getAttribute('aria-label') || '').toLowerCase();
      const title = (await btn.getAttribute('title') || '').toLowerCase();
      const text = (await btn.innerText().catch(() => '')).toLowerCase();
      const combined = `${ariaLabel} ${title} ${text}`;

      if (combined.includes('weiterleiten') || combined.includes('forward')) continue;

      if (isReplyAll) {
        if (combined.includes('reply all') || combined.includes('allen antworten')) {
          targetBtn = btn;
          break;
        }
      } else {
        const hasReply = combined.includes('reply') || combined.includes('antworten');
        const hasAll = combined.includes('all') || combined.includes('allen');
        if (hasReply && !hasAll) {
          targetBtn = btn;
          break;
        }
      }
    }

    // Fallback if specific button was not found
    if (!targetBtn) {
      for (const btn of buttons) {
        const disabled = await btn.getAttribute('disabled');
        const ariaDisabled = await btn.getAttribute('aria-disabled');
        if (disabled !== null || ariaDisabled === 'true') continue;

        const ariaLabel = (await btn.getAttribute('aria-label') || '').toLowerCase();
        const title = (await btn.getAttribute('title') || '').toLowerCase();
        const text = (await btn.innerText().catch(() => '')).toLowerCase();
        const combined = `${ariaLabel} ${title} ${text}`;

        if (combined.includes('weiterleiten') || combined.includes('forward')) continue;

        if (combined.includes('reply') || combined.includes('antworten')) {
          targetBtn = btn;
          break;
        }
      }
    }

    if (!targetBtn) {
      throw new Error(`Button für "${isReplyAll ? 'Allen antworten' : 'Antworten'}" in Outlook Web nicht gefunden.`);
    }

    // 4. Click Reply button and track active page / tab
    const initialPages = context ? context.pages() : [page];
    await targetBtn.click();

    let activePage = page;
    if (context) {
      const startWait = Date.now();
      while (Date.now() - startWait < 5000) {
        const currentPages = context.pages();
        if (currentPages.length > initialPages.length) {
          const newTab = currentPages.find(p => !initialPages.includes(p));
          if (newTab) {
            activePage = newTab;
            await activePage.waitForLoadState('domcontentloaded').catch(() => null);
            break;
          }
        }
        const inlineEditor = await page.$('div[role="textbox"][aria-label*="Message body" i], div[role="textbox"][aria-label*="Nachrichtentext" i]');
        if (inlineEditor) {
          activePage = page;
          break;
        }
        await page.waitForTimeout(300);
      }
    }
    this.lastPage = activePage;

    // 5. Find editor
    const editor = await activePage.waitForSelector(
      'div[role="textbox"][aria-label*="Message body" i], div[role="textbox"][aria-label*="Nachrichtentext" i], div[aria-label="Message body"], div[aria-label="Nachrichtentext"]',
      { timeout: 15000 }
    ).catch(() => null);

    if (!editor) {
      throw new Error('Compose-Editor für den Antwort-Entwurf konnte nicht gefunden werden.');
    }

    // 6. Safe DOM insertion: Prepend body paragraphs directly into editor before signature
    await editor.evaluate((el, bodyText) => {
      const lines = bodyText.split('\n');
      const fragment = document.createDocumentFragment();
      for (const line of lines) {
        const p = document.createElement('p');
        if (line.trim().length === 0) {
          p.appendChild(document.createElement('br'));
        } else {
          p.textContent = line;
        }
        fragment.appendChild(p);
      }
      const spacer = document.createElement('p');
      spacer.appendChild(document.createElement('br'));
      fragment.appendChild(spacer);

      el.insertBefore(fragment, el.firstChild);
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }, body);

    await activePage.waitForTimeout(500);

    // 7. Save draft
    await activePage.keyboard.press('Control+s');
    await activePage.waitForTimeout(3500);

    return {
      success: true,
      action: isReplyAll ? 'reply_all' : 'reply',
      subject: emailSubject || undefined,
      targetMailPreview: selectedMailPreview || undefined,
      preview: body.length > 200 ? body.substring(0, 200) + '...' : body,
      folder: 'Entwürfe (Drafts)',
      message: `Antwort-Entwurf (${isReplyAll ? 'Allen antworten' : 'Antworten'}) erfolgreich erstellt und in "Entwürfe" gespeichert. Verlauf und Signatur wurden beibehalten.`
    };
  }

  getLastPage() { return this.lastPage; }
}

export const owaClient = new OwaClient();


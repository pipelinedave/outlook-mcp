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

// Erkennt Event-Signaturen in einem ARIA-Label/Text. Reine Funktion (unit-testbar).
//  - Zeitfenster: "09:00 to 12:00" / "14:00 bis 15:00" / "9:00 AM - 10:00 AM"
//  - Ganztag: "all day event" / "all-day" / "ganztägig" / "ganztages"
//  - Statuswort (Busy|Tentative|Free|Besetzt|Vorbehalt|Frei) ZUSAMMEN mit einem Datum
export function isCalendarEventLabel(text) {
  if (!text) return false;
  const t = String(text);
  const timeRangeRe = /\d{1,2}:\d{2}\s*(?:AM|PM)?\s*(?:to|bis|[–—-])\s*\d{1,2}:\d{2}/i;
  if (timeRangeRe.test(t)) return true;
  const allDayRe = /all[- ]?day|ganztägig|ganztages/i;
  if (allDayRe.test(t)) return true;
  const statusRe = /\b(busy|tentative|free|besetzt|vorbehalt|frei)\b/i;
  const dateRe = /(\d{1,2}\.\s*[A-Za-zäöüÄÖÜß]{3,})|([A-Za-zäöüÄÖÜß]{3,}\.?\s+\d{1,2})|(\d{1,2}\s+[A-Za-zäöüÄÖÜß]{3,})/;
  return statusRe.test(t) && dateRe.test(t);
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
export function assembleFreeBusyDiagnostics({ fresh = null, early = null, weekButtonProbe = [], abortReason = 'none', budget = null } = {}) {
  const base = fresh || early || { url: '', headerCandidates: [], dayHeaderCandidates: [], weekNavCandidates: [], ariaButtonCount: 0, ariaLabelSample: [] };
  const out = {
    ...base,
    weekButtonProbe,
    abortReason
  };
  if (early && early.url) out.urlAfterNavigation = early.url;
  if (budget) out.budget = budget;
  return out;
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
    // Kein Zeitfenster im Label -> Ganztag-Event (start/end bewusst null)
    const title = label
      .replace(/[,;]?\s*(Ganztägig|Ganztägiges Ereignis|All day|All-day event).*/i, '')
      .replace(/[,;:\s]+$/, '')
      .trim() || '(Ganztägig)';
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

    // Prüft, ob Event-Elemente (Zeit- oder Ganztag-Labels) im Grid gerendert sind.
    const hasEventElements = async () => page.evaluate(() => {
      const timeRe = /\d{1,2}:\d{2}/;
      const allDayRe = /ganztägig|all day|all-day/i;
      return Array.from(document.querySelectorAll('[role="button"][aria-label]')).some((b) => {
        const label = b.getAttribute('aria-label') || '';
        return timeRe.test(label) || allDayRe.test(label);
      });
    }).catch(() => false);

    // Zustandswarten statt starrem Timeout: auf gerenderten Wochen-/Tages-Header
    // ODER Event-Elemente warten (Live-Fail 24.09.2026: 6s reichten nicht, das Grid
    // war noch nicht da -> 0 Slots, "Work week" als einziger Text). Poll-Intervall
    // 500ms mit Budget-Check je Iteration (Runde-4-Forderung A).
    const waitForCalendarRendered = async (budgetMs) => {
      const waitDeadline = Date.now() + budgetMs;
      let rendered = { text: null, startDate: null, endDate: null };
      let events = false;
      while (Date.now() < waitDeadline && remaining() > 0) {
        rendered = await readRenderedCalendarWeek(page);
        if (rendered.startDate) {
          events = await hasEventElements();
          break;
        }
        events = await hasEventElements();
        if (events) break;
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
    const clickWeekButton = async (direction, prevStart) => {
      let foundAny = false;
      let clickedAny = false;
      for (const cand of weekButtonCandidates(direction)) {
        if (remaining() < CLICK_COST_MS) {
          weekButtonProbe.push({ selector: cand.selector, kind: cand.kind, found: null, clicked: false, note: 'Budget aufgebraucht' });
          break;
        }
        const loc = page.locator(cand.selector).first();
        let found = false;
        try {
          await loc.waitFor({ state: 'visible', timeout: Math.min(1000, Math.max(100, remaining())) });
          found = true;
        } catch (e) { found = false; }
        if (!found) {
          weekButtonProbe.push({ selector: cand.selector, kind: cand.kind, found: false, clicked: false });
          continue;
        }
        foundAny = true;
        try {
          await loc.click({ timeout: Math.min(2000, Math.max(100, remaining())) });
        } catch (e) {
          weekButtonProbe.push({ selector: cand.selector, kind: cand.kind, found: true, clicked: false, clickError: String((e && e.message) || e).slice(0, 120) });
          continue;
        }
        clickedAny = true;
        const r = await waitRenderAfterClick(prevStart, phase(FREEBUSY_CLICK_RENDER_WAIT_MS));
        weekButtonProbe.push({ selector: cand.selector, kind: cand.kind, found: true, clicked: true, verified: r.verified });
        if (shouldTryNextCandidate(r.verified)) continue; // wirkungslos -> nächster Kandidat
        return { clicked: true, week: r.week, verified: r.verified, foundAny, clickedAny };
      }
      return { clicked: false, week: null, verified: 'not-found', foundAny, clickedAny };
    };

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
    const isRequestedWeek = (r) => Boolean(
      r && r.startDate && r.endDate &&
      r.startDate <= rangeStartDate && rangeStartDate <= r.endDate
    );

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

    for (let step = 0; step < 30; step++) {
      if (rendered.startDate && isRequestedWeek(rendered)) break;
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
          forcedNextRetryDone = true;
          const forced = await clickWeekButton(1, rendered.startDate);
          if (forced.clicked) {
            rendered = forced.week;
            hasEvents = hasEvents || (await hasEventElements());
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

      const r = await clickWeekButton(Math.sign(deltaWeeks), rendered.startDate);
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
        budget: freeBusyBudgetReport()
      });
      return errorResult(
        `Kalenderansicht nicht geladen — Kalender-Tab nicht erreicht oder Woche nicht gerendert. ` +
        `Weder ein datierter Wochen-Header noch Event-Elemente gefunden (nach Navigation, Warten und timeboxter Eskalation). ` +
        `Bitte erneut versuchen bzw. Session prüfen (outlook_login).`,
        { renderedWeek: rendered, diagnostics }
      );
    }

    // dataQuality: 'verified' = Wochen-Header gelesen und passt zum Range;
    // 'unverified' = Header nicht lesbar (Slots werden trotzdem geliefert, mit Warnung).
    const dataQuality = (rendered.startDate && rendered.endDate) ? 'verified' : 'unverified';

    // Alle Event-Elemente über ihre ARIA-Labels sammeln (roh, Dedupe + Klassifikation in Node).
    // Lockerer Match (Live-Fail 24.09.2026): Zeitmuster (HH:MM) ODER Ganztag-Kennzeichnung
    // genügt — DE/EN-Statuswörter ("Busy"/"Besetzt"/"to"/"bis") sind NICHT mehr Bedingung,
    // da die Labels je nach Sprache/Status variieren. Arbeitsort-Banner sind keine Belegung.
    const rawLabels = await page.evaluate(() => {
      const buttons = Array.from(document.querySelectorAll('[role="button"][aria-label]'));
      const timeRe = /\d{1,2}:\d{2}/;
      const allDayRe = /ganztägig|ganztages|all day|all-day/i;
      const nonEventRe = /arbeitsort|working location/i;
      const labels = [];
      for (const btn of buttons) {
        const label = (btn.getAttribute('aria-label') || '').trim();
        if (!label || label.length > 300) continue;
        if (nonEventRe.test(label)) continue;
        if (!timeRe.test(label) && !allDayRe.test(label)) continue;
        labels.push(label);
      }
      return labels;
    });

    // Deduplizieren (gleiche Events können mehrfach im DOM hängen), strukturiert parsen
    // und anhand der Wortliste klassifizieren (Busy/Tentative/Free).
    const seen = new Set();
    const parsedSlots = [];
    for (const label of rawLabels) {
      if (seen.has(label)) continue;
      seen.add(label);
      const slot = parseEventLabel(label, defaultYear);
      if (!slot) continue;
      slot.status = classifyEventStatus(label);
      parsedSlots.push(slot);
    }

    // Beauty-Fix (Live Runde 2): Work-Location-Banner und titellose Pseudo-Slots
    // ("(ohne Titel)", Büro/Homeoffice/URL-Schnipsel) sind keine echte Belegung
    // und werden aus den Slots ausgeschlossen.
    const realSlots = parsedSlots.filter((slot) =>
      slot.status !== 'non_event' && !isWorkLocationPseudoSlot(slot.title)
    );

    // Dedupe: identisches Datum+Zeit nur 1x; überlappende Duplikate gleicher Zeit
    // zusammenführen (bevorzugt wird der Slot mit echtem Titel behalten).
    const dedupedSlots = dedupeBusySlots(realSlots);

    // Nur Slots zurückgeben, die mit dem angefragten Range [start, end) überlappen.
    // Ganztag-Events belegen 00:00–24:00 ihres Tages. Explizit freie Slots ('free')
    // sind keine Belegung und werden ausgeschlossen.
    const rangeStartIso = String(startDt).includes('T') ? startDt : `${startDt}T00:00:00`;
    const rangeEndIso = String(endDt).includes('T') ? endDt : `${endDt}T23:59:59`;
    const rangeStartMs = new Date(rangeStartIso).getTime();
    const rangeEndMs = new Date(rangeEndIso).getTime();

    const busySlots = dedupedSlots
      .filter((slot) => slot.status !== 'free')
      .filter((slot) => {
        const dayStartMs = new Date(`${slot.date}T00:00:00`).getTime();
        const s = slot.allDay ? dayStartMs : new Date(`${slot.date}T${slot.start}:00`).getTime();
        const e = slot.allDay ? dayStartMs + 86400000 : new Date(`${slot.date}T${slot.end}:00`).getTime();
        return s < rangeEndMs && e > rangeStartMs;
      });

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
    if (dataQuality === 'verified' && abortReason !== 'none') {
      note += ` HINWEIS: Navigation vorzeitig beendet (${abortReason}); die angezeigte Woche wurde verifiziert.`;
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
      dataQuality,
      busySlots,
      count: busySlots.length,
      note
    };
    // Reiche Diagnostik bei unverifizierten Daten mitliefern (Runde-3/4-Forderung):
    // URL (inkl. früher Snapshot), Header-/Tageskopf-Kandidaten, ARIA-Sample,
    // Klick-Protokoll, Abbruchgrund und Budget-Report.
    if (dataQuality === 'unverified') {
      result.diagnostics = assembleFreeBusyDiagnostics({
        fresh: await freshDiag(),
        early: earlyDiagnostics,
        weekButtonProbe,
        abortReason,
        budget: freeBusyBudgetReport()
      });
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


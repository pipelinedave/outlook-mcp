import TurndownService from 'turndown';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { browserManager } from './browser.js';
import { config } from './config.js';

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

// Dateinamen auf sichere, plattformneutrale Form kürzen (keine Pfad-Trenner,
// keine Steuerzeichen, Längenlimit mit erhaltener Endung).
function sanitizeFilename(name) {
  let s = String(name || '').replace(/[\\/:*?"<>|\r\n\t]+/g, '_').replace(/\s+/g, ' ').trim();
  if (!s) s = 'anhang';
  const ext = path.extname(s);
  if (s.length > 120) {
    s = (ext ? s.slice(0, 120 - ext.length) : s).trim() + ext;
  }
  return s;
}

// ---------------------------------------------------------------------------
// Reply-Draft Body-Verifikation (Problem 1, Playwright-Serialisierung)
// ---------------------------------------------------------------------------
// OWA/ContentEditable-Editoren übernehmen Inhalte teils nur über Tastatur-/
// Clipboard-Input statt über DOM-Mutationen. Nach dem Einfügen muss deshalb
// VERIFIZIERT werden, dass der Text wirklich im Editor gelandet ist, bevor
// gespeichert wird (Fehlerbild: Entwurf war leer).
//
// Reine Funktionen, unit-testbar: Die DOM-Abfrage liefert einen Rohzustand
// (textLength, sample, contentEditable, targetTag), diese Prädikate entscheiden
// über "Text ist nachweislich drin" und über "Fallback nötig".

// Prüft, ob die erste nicht-leere Zeile des erwarteten Bodys im Editor-Text
// nachweisbar ist. Vergleich gegen den GESAMTEN innerText (nicht nur den
// Anfang), weil OWA den Body an anderer Editor-Position platzieren kann.
// Whitespace-Normalisierung (25.10. / 15.10.2026-Befund): OWA rendert Umbrüche
// im Editor anders, als der Solltext sie liefert — der rohe Substring-Vergleich
// blockierte sonst berechtigte Saves. Gross/Klein-Abweichungen zählen bewusst
// NICHT als exakter Treffer (dafür: classifyBodyMatch -> 'fuzzy').
export function verifyBodyInserted(text, expectedBody) {
  const t = String(text || '').trim();
  const firstLine = String(expectedBody || '').split('\n').find((l) => l.trim().length > 0) || '';
  if (!firstLine || !t) return false;
  if (t.includes(firstLine.trim())) return true;
  return normalizeBodyForCompare(t).includes(normalizeBodyForCompare(firstLine));
}

// Der Editor ist genau dann ein echter Eingabe-Editor, wenn er NICHT auf
// contenteditable=false steht (dann wäre es ein reiner Anzeige-Knoten, der
// die DOM-Insertion nicht übernehmen würde).
export function isWritableEditor(contentEditable) {
  return String(contentEditable) !== 'false';
}

// Echte-Editor-Prüfung (Live-Befund 09.10.2026): Im OWA-Compose-Fenner gibt
// es MEHRERE Knoten mit aria-label "Message body" — einen echten Editor
// (role="textbox" / contenteditable="true") und einen Wrapper ohne
// contenteditable. Ein Wrapper liefert innerText mit CSS-Quelltext statt
// Nachrichtentext; jede darauf gestützte Verifikation ist wertlos.
// Der Editor wird deshalb als ECHTE EIGENSCHAFT bestimmt, nicht über sein Label.
export function isRealComposeEditor({ contentEditable = null, role = '' } = {}) {
  return contentEditable === 'true' || String(role) === 'textbox';
}

// Unversehrtheits-Prüfung für den Tastatur-Fallback: Der VORHER vorhandene
// Inhalt (E-Mail-Verlauf, Signatur) muss nach dem Tippen noch nachweisbar sein.
// Fehlt er, hat das Tippen den Bestand überschrieben — das Entwurfs-Ergebnis
// wäre dann inhaltlich verfälscht und darf NICHT gespeichert werden.
// Verglichen wird ein Präfix (200 Zeichen), weil OWA beim Tippen Leerzeilen und
// Signatur-Whitespace umsortieren kann; der ANFANG des Bestands bleibt aber
// stabil. Ein leerer Bestand (frisches Compose) ist trivially unversehrt.
export function isHistoryPreserved(beforeText, afterText, probeLength = 200) {
  const before = String(beforeText || '').trim();
  if (!before) return true; // nichts vorhanden -> nichts zu verlieren
  const probe = before.slice(0, Math.max(1, probeLength));
  return String(afterText || '').includes(probe);
}

// Gesamturteil über die DOM-Rückmeldung der Insertions-Verifikation.
// true  -> Text nachweislich im Editor, Speichern ist sicher.
// false -> kein Beweis (Text fehlt ODER Editor nicht schreibbar) -> Fallback.
export function resolveBodyInserted({ text = '', expectedBody = '', contentEditable = null } = {}) {
  if (!isWritableEditor(contentEditable)) return false;
  return verifyBodyInserted(text, expectedBody);
}

// Whitespace-normalisierte Vergleichsbasis: OWA kann Umbrueche/Leerzeichen im
// Editor anders rendern, als der Solltext sie liefert. Gross/Kleinschreibung
// bleibt bewusst ERHALTEN (Abweichungen davon sind Autokorrektur-Artefakte
// und werden von classifyBodyMatch als 'fuzzy' gemeldet, nicht kaschiert).
export function normalizeBodyForCompare(s) {
  return String(s || '').replace(/\s+/g, ' ').trim();
}

// Dreistufiger Body-Nachweis (Befund 15.10.2026):
//   'exact' -> erste Soll-Zeile 1:1 im Editor-Text (raw oder whitespace-normalisiert)
//   'fuzzy' -> erste Soll-Zeile nur case-insensitive gefunden (typisches
//              Autokorrektur-Artefakt: "dem" -> "Dem" nach "bzw. ")
//   'none'  -> nicht nachweisbar
export function classifyBodyMatch(text, expectedBody) {
  const t = String(text || '');
  const firstLine = String(expectedBody || '').split('\n').find((l) => l.trim().length > 0) || '';
  if (!firstLine || !t.trim()) return 'none';
  const normT = normalizeBodyForCompare(t);
  const normF = normalizeBodyForCompare(firstLine);
  if (normT.includes(normF)) return 'exact';
  if (normT.toLowerCase().includes(normF.toLowerCase())) return 'fuzzy';
  return 'none';
}

// Kompakte Soll/Ist-Diagnose fuer Fehler-Returns — verhindert den stillen
// Misserfolg der harten Zeichen-Verifikation (sie blockierte den Save, ohne
// klar zu sagen, WELCHER Text erwartet und WAS stattdessen gefunden wurde).
export function buildBodySollIstDiff(expectedBody, actualText) {
  const expLines = String(expectedBody || '').split('\n').map((l) => l.trim()).filter(Boolean);
  const actLines = String(actualText || '').split('\n').map((l) => l.trim()).filter(Boolean);
  return {
    match: classifyBodyMatch(actualText, expectedBody),
    expectedLength: String(expectedBody || '').length,
    actualLength: String(actualText || '').length,
    expectedFirstLine: (expLines[0] || '').slice(0, 160),
    actualFirstLine: (actLines[0] || '').slice(0, 160)
  };
}

// Toleranter Beweis, dass ein erwarteter Wert (Absender/Empfaenger/Betreff)
// in den gesammelten Beweis-Texten der Quellmail steckt (Befund 15.10.2026,
// Fehl-Reply auf Prignitz/Isenberg statt UHK). Toleranzstufen:
//   1) normalisierter Substring
//   2) E-Mail: Local-Part oder Domain einzeln (Anzeige zeigt oft nur Namen)
//   3) Namen: ALLE Token (>=3 Zeichen) muessen vorkommen ("Spieker" in "Spieker, Leonie")
export function evidenceMatchesExpected(evidence, expected) {
  const exp = normalizeBodyForCompare(expected).toLowerCase();
  if (!exp) return false;
  const ev = normalizeBodyForCompare(evidence).toLowerCase();
  if (!ev) return false;
  if (ev.includes(exp)) return true;
  if (exp.includes('@')) {
    const at = exp.indexOf('@');
    const local = exp.slice(0, at).replace(/[^\w.+-]/g, '');
    const domain = exp.slice(at + 1);
    if (local.length >= 3 && ev.includes(local)) return true;
    if (domain.length >= 4 && ev.includes(domain)) return true;
    return false;
  }
  const tokens = exp.replace(/,/g, ' ').split(/\s+/).filter((t) => t.length >= 3);
  if (tokens.length && tokens.every((t) => ev.includes(t))) return true;
  return false;
}

// Flexibles Datum-Input fuer die Reply-Assertion: 'YYYY-MM-DD' oder 'DD.MM.YYYY'
// -> { y, m, d } | null. Bewusst nur Datumsanteile — die Uhrzeit der Quellmail
// ist fuer den Thread-Beleg irrelevant.
export function parseFlexibleDateInput(s) {
  const str = String(s || '').trim();
  let m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return { y: +m[1], m: +m[2], d: +m[3] };
  m = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) return { y: +m[3], m: +m[2], d: +m[1] };
  return null;
}

// REPLY-ZIEL-ASSERTION (Sicherheitsgate, Befund 15.10.2026): Prueft die
// gewaehlte Quellmail gegen die erwarteten Werte, BEVOR der Reply-Button
// geklickt wird. Kein Treffer = Abbruch — keine implizite "naechste beste
// Mail"-Logik. Rueckgabe:
//   { ok: true,  unverified: [...] }              — alle Pruefungen bestanden (oder unverifizierbar, dort gelistet)
//   { ok: false, reason, details }                — MISMATCH: Abbruch
export function resolveReplyTargetAssertion({ evidence = '', senderEvidence = '', recipientEvidence = '', expectedFrom, expectedTo, expectedSubject, expectedDate, rowTimestamp = null } = {}) {
  const unverified = [];
  const details = {};

  if (expectedSubject) {
    if (evidenceMatchesExpected(evidence, expectedSubject)) {
      details.subject = 'bestaetigt';
    } else {
      return {
        ok: false,
        reason: 'expected_subject',
        details: { ...details, subject: `NICHT gefunden: "${expectedSubject}"` }
      };
    }
  }

  if (expectedFrom) {
    if (evidenceMatchesExpected(senderEvidence || evidence, expectedFrom)) {
      details.from = 'bestaetigt';
    } else {
      return {
        ok: false,
        reason: 'expected_from',
        details: { ...details, from: `NICHT gefunden: "${expectedFrom}"` }
      };
    }
  }

  if (expectedTo) {
    // Empfaenger-Chips sind in OWA teils eingeklappt ("An: 3 weitere") — dann
    // ist der Check unverifizierbar und wird gemeldet, statt zu blockieren.
    if (!recipientEvidence || !String(recipientEvidence).trim()) {
      unverified.push('expected_to (Empfänger im Lesebereich nicht sichtbar)');
    } else if (evidenceMatchesExpected(recipientEvidence, expectedTo)) {
      details.to = 'bestaetigt';
    } else {
      return {
        ok: false,
        reason: 'expected_to',
        details: { ...details, to: `NICHT gefunden: "${expectedTo}"` }
      };
    }
  }

  if (expectedDate) {
    const want = parseFlexibleDateInput(expectedDate);
    if (!want) {
      unverified.push(`expected_date (Format unlesbar: "${expectedDate}" — bitte YYYY-MM-DD oder DD.MM.YYYY)`);
    } else if (rowTimestamp === null || rowTimestamp === undefined) {
      unverified.push('expected_date (Datum der Quellmail nicht aus der Listenzeile parsebar)');
    } else {
      const got = new Date(rowTimestamp);
      const sameDay = got.getFullYear() === want.y && (got.getMonth() + 1) === want.m && got.getDate() === want.d;
      if (sameDay) {
        details.date = 'bestaetigt';
      } else {
        return {
          ok: false,
          reason: 'expected_date',
          details: {
            ...details,
            date: `Quellmail ist vom ${String(got.getDate()).padStart(2, '0')}.${String(got.getMonth() + 1).padStart(2, '0')}.${got.getFullYear()}, erwartet ${expectedDate}`
          }
        };
      }
    }
  }

  return { ok: true, unverified, details };
}

// Beweis, dass eine Empfänger-Adresse in einer gelesenen Empfängerliste steckt
// (Befund 15.10.2026: CC via &cc=DeepLink wurde von OWA teils nicht persistiert,
// der Tool-Return sagte trotzdem cc ok). Tolerant gegenüber Chip-Labels der Form
// "Name <mail@x.de>" und Mehrfach-Eingaben (Semikolon/Komma-getrennt).
export function recipientListContains(list, expected) {
  if (!expected || !String(expected).trim()) return true;
  const hay = (Array.isArray(list) ? list : [list])
    .map((x) => String(x || '').toLowerCase())
    .join(' ; ');
  if (!hay.trim()) return false;
  const parts = String(expected).toLowerCase().split(/[;,]+/)
    .map((s) => s.trim())
    .filter(Boolean);
  if (parts.length === 0) return hay.includes(String(expected).toLowerCase().trim());
  // Jede erwartete Adresse muss nachweisbar sein; Alternative: Match auf den
  // reinen Adressteil (Chip-Label "Name <mail@x.de>" enthält die Adresse).
  return parts.every((p) => {
    if (hay.includes(p)) return true;
    const addr = p.match(/<([^>]+)>/);
    return addr ? hay.includes(addr[1].trim()) : false;
  });
}

// Findet eine E-Mail-Adresse in freiem Suchtext (fuer den Sent-Enumeration-
// Fallback: "Mail an uhk@example.de vom 12.10." -> "uhk@example.de").
export function extractEmailLike(s) {
  const m = String(s || '').match(/[\w.+-]+@[\w-]+(?:\.[\w-]+)+/);
  return m ? m[0] : null;
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

  // Ein-Tag-zuerst / EIN Monat: "21–27 September, 2026" (OWA-Form des
  // angezeigten Headers, Live-Beleg 26.09.2026: "21–27 September, 2026,
  // Jump to a specific date or date range." bzw. "… (Week 39)"). Weder das
  // DE- noch das EN- noch das EN-Tag-zuerst-Muster griffen hier: es fehlt der
  // Monatsname NACH dem zweiten Tag. Vor dem Fix fiel dieser Knoten deshalb
  // auf den next-Button zurueck (der als "28 September – 4 October" durchaus
  // parsebar war) — der echte Header wurde also gefunden und dann weggeworfen.
  const dayFirstOneMonth = t.match(/(\d{1,2})\s*[–—-]\s*(\d{1,2})\s+([A-Za-zäöüÄÖÜß]{3,})\.?(?:,?\s*(\d{4}))?/);
  if (dayFirstOneMonth) {
    const month = MONTH_NAMES[dayFirstOneMonth[3].toLowerCase()];
    if (month !== undefined) {
      const year = dayFirstOneMonth[4] ? parseInt(dayFirstOneMonth[4], 10) : currentYear;
      const s = new Date(year, month, parseInt(dayFirstOneMonth[1], 10));
      let e = new Date(year, month, parseInt(dayFirstOneMonth[2], 10));
      if (e < s) e = new Date(year + 1, month, parseInt(dayFirstOneMonth[2], 10));
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

// Dasselbe fuer die Wochen-Header-/Navigations-Muster (FIX 1). Die
// Pattern-Saetze selbst sind serialisierbare Strings und werden ueber die
// evaluate-Grenze gereicht; kompiliert wird hier und dort aus DIESER Quelle.
let compiledWeekRange = null;
let compiledWeekNav = null;
let compiledWeekPatternsSource = null;
function weekRangePatterns() {
  if (!compiledWeekRange || compiledWeekPatternsSource !== WEEK_RANGE_PATTERNS) {
    compiledWeekRange = {};
    for (const key of Object.keys(WEEK_RANGE_PATTERNS)) {
      compiledWeekRange[key] = new RegExp(WEEK_RANGE_PATTERNS[key], 'i');
    }
    compiledWeekPatternsSource = WEEK_RANGE_PATTERNS;
    compiledWeekNav = null;
  }
  return compiledWeekRange;
}
function weekNavPatterns() {
  if (!compiledWeekNav || compiledWeekPatternsSource !== WEEK_RANGE_PATTERNS) {
    weekRangePatterns();
    compiledWeekNav = WEEK_NAV_LABEL_PATTERNS.map((p) => new RegExp(p, 'i'));
  }
  return compiledWeekNav;
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

// ---------------------------------------------------------------------------
// FIX 1 (Live-Befund Lauf 5, 26.09.2026): der angezeigte Wochen-Header wurde
// aus dem ARIA-Label des "Go to next week"-Buttons gelesen.
// ---------------------------------------------------------------------------
// Live-Belege (beide Aufrufe desselben Laufs, Commit 3907f0d):
//   Aufruf 1: angefragt 21.09.–02.10., View stand korrekt auf KW 39
//             (21.–27.09.) -> renderedWeek wurde 2026-09-28–2026-10-04
//             gemeldet (= next-Button-Label) -> Abbruch 'max-clicks' nach 6
//             Klicks (prev/next/prev/next), obwohl die View RICHTIG war.
//   Aufruf 2: next-Button-Label = "28 September – 4 October, 2026" = Zielwoche
//             -> weekVerified:true, dataQuality:'verified', 0 Klicks,
//             View blieb auf KW 39 -> Falschpass; alle 10 Events wurden als
//             'date-out-of-range' verworfen (21.–25.09.).
// Der echte Header lag die ganze Zeit im DOM:
//   "21–27 September, 2026, Jump to a specific date or date range."
//   "21–27 September, 2026 (Week 39)"   (dort als 'header-misclassified'
//                                        verworfen)
//
// KERN: Header und Navigations-Button haben dieselbe FORM (Datumsbereich +
// Monatsnamen). Unterschieden sind sie nur an der NAVIGATIONS-ABSICHT, und
// die steht im Label-ANFANG. Deshalb ist das entscheidende Kriterium NICHT
// "sieht wie ein Datumsbereich aus", sondern "ist es ein Datumsbereich UND
// kein Navigations-Label".
//
// Die Positiv-Bedingung bleibt der Bereich: ein gueltiger Header MUSS einen
// ableitbaren Datumsbereich liefern (parseWeekHeaderText). Die Bereich-Pruefung
// wird VOR jeder Text-Heuristik angewandt — die Reihenfolge in
// resolveRenderedWeekFromCandidates ist genau das.
//
// Als Strings (serialisierbar ueber die page.evaluate-Grenze, vgl.
// EVENT_LABEL_PATTERNS): die Browser-Seite braucht dieselbe Definition, ohne
// die evaluate-Grenze zu verletzen.
export const WEEK_NAV_LABEL_PATTERNS = [
  // EN: "Go to previous week 21–27 September, 2026 (Week 39)" / "Go to next week 28 September – 4 October, 2026"
  String.raw`\bgo\s+to\s+(?:the\s+)?(?:previous|prior|preceding|next)\s+week\b`,
  String.raw`\bgo\s+to\s+(?:the\s+)?week\s+number\b`,
  String.raw`\b(?:previous|prior|preceding|next)\s+week\b`,
  // EN: "Week number 39" als eigenstaendiger Indikator
  String.raw`\bweek\s+number\s*\d{1,2}\b`,
  // DE: "Zur vorherigen Woche" / "Nächste Woche" / "Woche 41" (KW-Schalter)
  String.raw`\b(?:vorherige[nrs]?|vorige[nrs]?|nächste[nrs]?|nachfolgende[nrs]?|naechste[nrs]?)\s+woche\b`,
  String.raw`\bwochennummer\s*\d{1,2}\b`,
  String.raw`\bwoche\s+(?:vorherige[nrs]?|zur(?:ü|ue)ck|vor|nach)\b`
];

// Ist das Label eine Wochennavigation? Dann beschreibt es die Woche, die der
// Button ANSTEUERN wuerde — sie ist per Definition NICHT die angezeigte.
// Reine Funktion, unit-testbar, Spiegel der Browser-Logik in
// collectCalendarCandidates.
export function isWeekNavLabel(text) {
  if (!text) return false;
  const t = String(text).replace(/\s+/g, ' ').trim();
  if (!t) return false;
  for (const re of weekNavPatterns()) if (re.test(t)) return true;
  return false;
}

// Der angezeigte Header traegt die Absicht "springe zu Datum/Datumspanne"
// ("Jump to a specific date or date range") — bzw. die reine Angabe der
// angezeigten Woche mit Wochennummer ("21–27 September, 2026 (Week 39)").
// Beide Formen sind Header und wurden in Lauf 5 als 'header-misclassified'
// verworfen bzw. nie als Header-Kandidat gefuehrt.
export const WEEK_DISPLAYED_HEADER_PATTERN = String.raw`jump\s+to\s+a\s+specific\s+date|date\s+range`;

// Ein Datumsbereich als serialisierbare Muster (Browser-Seite braucht dieselbe
// Definition ueber die evaluate-Grenze, vgl. EVENT_LABEL_PATTERNS):
//   range:  "21–27 September", "28 September – 4 October", "28.9. – 4.10.2026"
//   dotted: "28.9." — der Ersatzbeleg fuer einen Bereich OHNE Monatsnamen
export const WEEK_RANGE_PATTERNS = {
  range: String.raw`\d{1,2}\.?\s*[\u2013\u2014-]\s*\d{1,2}`,
  dotted: String.raw`\d{1,2}\.\d{1,2}`
};

// Sieht dieser Text aus wie ein Wochen-Header? Positiv: er traegt einen
// Datumsbereich MIT Datumsanker (Monatsname oder numerischer Tag) UND ist
// keine Navigation. Der Bereich kommt VOR jeder Text-Heuristik.
export function isWeekHeaderShaped(text) {
  if (!text || isWeekNavLabel(text)) return false;
  const t = String(text).replace(/\s+/g, ' ').trim();
  const re = eventPatterns();
  const rr = weekRangePatterns();
  if (!rr.range.test(t)) return false;
  return re.dateish.test(t) || rr.dotted.test(t);
}

// WOHER kam der gelesene Wochenwert? Diese Herkunft entscheidet, ob er als
// Wochen-Nachweis gilt (isHeaderReadVerified) — ein aus Prev/Next-Buttons
// ABGELEITETER Wert ist eine Ableitung, keine Anzeige.
export const WEEK_READ_SOURCES = {
  // Datumsbereich aus einem Knoten der Header-Region gelesen.
  displayed: 'displayed-header',
  // Aus einem Tages-Spaltenkopf abgeleitet (Montag der Woche).
  column: 'day-column',
  // Aus den Prev/Next-Buttons abgeleitet (Mitte zwischen beiden) — GESCHAETZT.
  navDerived: 'nav-derivation',
  // Einzel-Datum mit Monatsnamen abgeleitet — GESCHAETZT.
  singleDate: 'single-date',
  // Numerischer Tageskopf abgeleitet — GESCHAETZT.
  numericDay: 'numeric-day'
};

// Nur diese beiden Herkuenfte belegen den ANGEZEIGTEN Wochen-Header. Die
// uebrigen sind Ableitungen und duerfen keine 'weekVerified' erkaufen.
export function isDisplayedWeekSource(source) {
  return source === WEEK_READ_SOURCES.displayed || source === WEEK_READ_SOURCES.column;
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

// FIX 4 (Live-Lauf 2): busySlotsReliable und die eigene Widerspruchs-Warnung
// standen gleichzeitig: busySlotsReliable:true und "18 Event-Element(e) im DOM
// gefunden, aber 0 Busy-Slots extrahiert". Ein Ergebnis, das sich selbst
// widerspricht, ist nicht auswertbar — deshalb ist der Widerspruch ein HARTES
// false und keine Ermessensfrage.
// Reine Funktion (unit-testbar), benutzt vom Aufrufer; die Warnung selbst
// bleibt in slotExtractionWarning (Single Source of Truth fuer den Text).
export function resolveBusySlotsReliable({ dataQuality = null, eventElementsFound = 0, busySlotCount = 0 } = {}) {
  if (!isBelievableBusySlots(dataQuality)) return false;
  return slotExtractionWarning(Number(eventElementsFound) || 0, Number(busySlotCount) || 0) === null;
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
  return page.evaluate((navPatterns) => {
    // FIX 1: dieselbe Navigations-Erkennung wie auf der Node-Seite
    // (WEEK_NAV_LABEL_PATTERNS). Sie wird als ARGUMENT ueber die
    // evaluate-Grenze gereicht — eine nackte Modul-Referenz waere dort ein
    // ReferenceError (siehe collectError unten).
    const navRes = [];
    for (const p of (navPatterns || [])) navRes.push(new RegExp(p, 'i'));
    const isNavLabel = (text) => {
      const t = String(text || '').replace(/\s+/g, ' ').trim();
      if (!t) return false;
      for (const re of navRes) if (re.test(t)) return true;
      return false;
    };
    const cap = (arr, n) => arr.slice(0, n);
    // Ein Navigations-Label ist NIE ein Header-Kandidat: es nennt die Woche,
    // die der Button ANSTEUERN wuerde, nicht die ANGEZEIGTE.
    const pushHeader = (arr, t) => { if (!isNavLabel(t)) push(arr, t); };
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
      pushHeader(headerTexts, text);
    }
    // 2) Toolbar-/Button-Elemente mit Datums-Label: OWA rendert den Wochenbereich als
    //    klickbaren Datums-Button. Der Ansichtsumschalter ("Work week"/"Arbeitswoche")
    //    enthält KEIN Datum und fällt hier bewusst heraus.
    for (const el of document.querySelectorAll(
      '[role="banner"] button, [role="banner"] [role="button"], header button, [role="toolbar"] button, [role="toolbar"] [role="button"]'
    )) {
      if (!isVisible(el)) continue;
      const label = el.getAttribute('aria-label') || el.innerText || el.getAttribute('title') || '';
      if (dateish.test(label) && !timeRe.test(label) && !isEventLabel(label)) pushHeader(headerTexts, label);
    }
    // 3) Alle aria-labels mit Datums-/Monatsmuster, aber OHNE Uhrzeit (z.B. der
    //    Datumsauswahl-Button). Zeitmuster-/Event-Labels sind Event-Elemente, keine Header.
    //    FIX 1: die Prev/Next-Buttons fallen hier durch — genau hier stand in
    //    Lauf 5 "Go to next week 28 September – 4 October, 2026" und wurde als
    //    gerenderte Woche gemeldet.
    for (const el of document.querySelectorAll('[aria-label]')) {
      if (!isVisible(el)) continue;
      const label = el.getAttribute('aria-label') || '';
      if (dateish.test(label) && !timeRe.test(label) && label.length <= 120 && !isEventLabel(label)) pushHeader(headerTexts, label);
    }
    // 3b) Prev/Next-Wochen-Buttons. FIX 1: die alte Bedingung verlangte ZUSAETZLICH
    //     eine Wochennummer ("(Week 39)"). Der Live-Button vom 26.09. hatte aber
    //     KEINE — "Go to next week 28 September – 4 October, 2026" fiel deshalb
    //     durch und blieb ein Header-Kandidat. Jetzt entscheidet allein die
    //     Navigations-Absicht (identisch zu isWeekNavLabel auf der Node-Seite).
    for (const el of document.querySelectorAll('[aria-label], [title]')) {
      if (!isVisible(el)) continue;
      const label = el.getAttribute('aria-label') || el.getAttribute('title') || '';
      if (isNavLabel(label)) push(navTexts, label);
    }
    // 4) Tages-Spaltenköpfe der Woche — mehrere Selektor-Varianten (OWA nutzt je nach
    //    Build role="columnheader", th-Elemente oder Grid-Köpfe)
    for (const el of document.querySelectorAll(
      '[role="columnheader"], [role="grid"] [role="columnheader"], th, [role="grid"] th'
    )) {
      if (!isVisible(el)) continue;
      const text = el.getAttribute('aria-label') || el.innerText;
      if (isEventLabel(text)) continue;
      pushHeader(dayHeaderTexts, text);
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
        if ((weekdayRe.test(s) && lineDateRe.test(s)) || rangeRe.test(s)) pushHeader(dayHeaderTexts, s);
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
  }, WEEK_NAV_LABEL_PATTERNS).catch((e) => ({
    // BUG 5 (Diagnose, Runde 4): Der frueher nackte `catch` lieferte exakt die
    // Form "nichts gefunden". Damit waren zwei voellig verschiedene Ursachen NICHT
    // unterscheidbar:
    //   H1: der Header-Bereich war noch nicht gerendert,
    //   H2: dieses evaluate selbst ist fehlgeschlagen (ReferenceError/Timeout) —
    //       die Kandidatenliste ist dann namlich ebenfalls leer.
    // Live-Runde 4 zeigte headerCandidates/dayHeaderCandidates/weekNavCandidates
    // alle LEER bei gleichzeitig vorhandenen Event-Chips (anderer evaluate,
    // erfolgreich). Der Fehler wird deshalb benannt statt als "nichts gefunden"
    // maskiert zu werden.
    headerTexts: [], dayHeaderTexts: [], navTexts: [], ariaButtonCount: 0, ariaLabelSample: [],
    collectError: String((e && e.message) || e).slice(0, 160)
  }));
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
      ({ pats, plan, clockStrict, weekStart, fastMode, maxElements, maxCensusNodes, maxStructureNodes, loadPats, loadingSelector, navPats, rangePats }) => {
        const re = {};
        for (const k of Object.keys(pats)) re[k] = new RegExp(pats[k], 'i');
        // FIX 1: Wochen-Header- vs. Navigations-Erkennung. Beide Pattern-Sätze
        // sind serialisierbare Strings und werden als ARGUMENT gereicht — eine
        // nackte Modul-Referenz waere ueber die evaluate-Grenze ein
        // ReferenceError (vgl. collectError unten). Die Definition selbst ist
        // dieselbe wie auf der Node-Seite (WEEK_NAV_LABEL_PATTERNS /
        // WEEK_RANGE_PATTERNS).
        const nre = (navPats || []).map((p) => new RegExp(p, 'i'));
        const rre = {};
        for (const k of Object.keys(rangePats || {})) rre[k] = new RegExp(rangePats[k], 'i');
        const isNavLabel = (t) => {
          for (let i = 0; i < nre.length; i++) if (nre[i].test(t)) return true;
          return false;
        };
        // Sieht der Text aus wie ein Wochen-Header? Positiv: Datumsbereich MIT
        // Datumsanker UND keine Navigations-Absicht. Spiegel von
        // isWeekHeaderShaped auf der Node-Seite.
        const isWeekHeaderShaped = (t) => {
          if (!t || isNavLabel(t)) return false;
          if (!rre.range || !rre.range.test(t)) return false;
          return re.dateish.test(t) || (rre.dotted ? rre.dotted.test(t) : false);
        };
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
          // FIX 1 (Live-Lauf 5): der angezeigte Wochen-Header wurde an dieser
          // Stelle als 'header-misclassified' VERWORFEN — "21–27 September,
          // 2026 (Week 39)". Er ist der Wochen-Header, kein Muell. Erkannt
          // wird er ueber den BEREICH (starkes Kriterium) und die Abwesenheit
          // einer Navigations-Absicht — nicht ueber eine Text-Heuristik.
          if (isWeekHeaderShaped(t)) return 'displayed-week-header';
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
        const weekHeaderLike = [];
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
              // FIX 1: der angezeigte Wochen-Header ist kein 'header-misclassified'.
              // Er wird als eigener, benannter Beleg gefuehrt — damit die
              // Diagnostik zeigt, dass der Header GEFUNDEN wurde (und nicht,
              // wie vorher, als Muell verworfen).
              else if (r === 'displayed-week-header') weekHeaderLike.push(t);
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
          // FIX 1: der angezeigte Wochen-Header, vom Kollektor als eigener
          // Grund benannt (nicht als 'header-misclassified' verworfen).
          weekHeaderLabels: cap(dedupe(weekHeaderLike), 10),
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
        loadingSelector: LOAD_INDICATOR_SELECTOR,
        // FIX 1: Header- vs. Navigations-Erkennung (siehe WEEK_NAV_LABEL_PATTERNS).
        navPats: WEEK_NAV_LABEL_PATTERNS,
        rangePats: WEEK_RANGE_PATTERNS
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
      sourceTried: [], hiddenCount: 0, headerLikeLabels: [], weekHeaderLabels: [], nonEventLabels: [],
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
  const deriveWeek = (dateIso, text, source) => {
    const d = new Date(dateIso + 'T00:00:00');
    const offsetToMonday = d.getDay() === 0 ? -6 : 1 - d.getDay();
    const monday = new Date(d.getTime() + offsetToMonday * 86400000);
    const sunday = new Date(monday.getTime() + 6 * 86400000);
    return { text: String(text).slice(0, 200), startDate: toIsoDate(monday), endDate: toIsoDate(sunday), weekNumber: null, source, readError: 'no-displayed-header' };
  };
  const currentYear = now.getFullYear();

  // Event-Labels strikt ausschließen (Runde-4-Root-Cause): Ein All-Day-Event
  // ("Urlaub … September 23") darf NIE als Wochen-Header durchkommen.
  const headers = filterCalendarHeaderCandidates(headerTexts);
  const dayHeaders = filterCalendarHeaderCandidates(dayHeaderTexts);

  // FIX 1 (Live-Lauf 5): Navigations-Labels sind NIE der angezeigte Header.
  // Ihr Datumsbereich ist das Klickziel, nicht die Ansicht — in Lauf 5 hat
  // der erste parsebare Nav-Kandidat die Zielwoche zweimal kapert: einmal
  // mitten in der RICHTIGEN Woche (Falschabbruch nach 6 Klicks, 'max-clicks')
  // und einmal als Falschpass (weekVerified:true, View blieb auf KW 39).
  const displayedHeaders = headers.filter((t) => !isWeekNavLabel(t));
  const dayHeadersOnly = dayHeaders.filter((t) => !isWeekNavLabel(t));

  // 1) Wochen-Range-Header (Header-Bereich zuerst, dann Tageskopf-/Body-Kandidaten)
  //    POSITIV-KRITERIUM: der Text muss einen ableitbaren DATUMSBEREICH
  //    liefern (parseWeekHeaderText). Diese Bereich-Pruefung steht VOR jeder
  //    Text-/Nummern-Heuristik — sie ist das staerkste Kriterium.
  for (const text of displayedHeaders) {
    const parsed = parseWeekHeaderText(text, now);
    if (parsed) return { text: text.slice(0, 200), startDate: parsed.startDate, endDate: parsed.endDate, weekNumber: null, source: WEEK_READ_SOURCES.displayed };
  }
  for (const text of dayHeadersOnly) {
    const parsed = parseWeekHeaderText(text, now);
    if (parsed) return { text: text.slice(0, 200), startDate: parsed.startDate, endDate: parsed.endDate, weekNumber: null, source: WEEK_READ_SOURCES.column };
  }
  // 2) Prev/Next-Button-Ableitung (belastbarer Fallback, Runde-4-Forderung 2).
  //    Das ist eine ABLEITUNG (Mitte zwischen prev und next), keine Anzeige.
  //    FIX 1.4: Sie ersetzt den angezeigten Header NICHT — sie wird als
  //    readError gemeldet und zaehlt nicht als Wochen-Nachweis
  //    (isHeaderReadVerified), damit aus einer Schätzung keine Zertifizierung
  //    wird. Nur der next-Button-Text als SOLCHE (sein Klickziel) wird nie
  //    zur gerenderten Woche.
  const nav = resolveWeekFromNavButtons(navTexts, now);
  if (nav && nav.startDate) {
    return {
      text: nav.text, startDate: nav.startDate, endDate: nav.endDate, weekNumber: nav.weekNumber,
      source: WEEK_READ_SOURCES.navDerived,
      readError: 'no-displayed-header',
      derivedFrom: 'nav-buttons',
      navOnlyCandidates: headers.concat(navTexts).filter((t) => isWeekNavLabel(t)).slice(0, 5)
    };
  }
  // 3) Einzel-Datum mit Monatsnamen
  for (const text of displayedHeaders) {
    const slot = parseEventLabel(text, currentYear);
    if (slot && slot.date) return deriveWeek(slot.date, text, WEEK_READ_SOURCES.singleDate);
  }
  for (const text of dayHeadersOnly) {
    const slot = parseEventLabel(text, currentYear);
    if (slot && slot.date) return deriveWeek(slot.date, text, WEEK_READ_SOURCES.singleDate);
  }
  // 4) Numerische Tagesangaben (Spaltenköpfe + kurze Header-Texte)
  for (const text of dayHeadersOnly) {
    const dateIso = parseNumericDayHeaderText(text, now);
    if (dateIso) return deriveWeek(dateIso, text, WEEK_READ_SOURCES.numericDay);
  }
  for (const text of displayedHeaders) {
    if (String(text).length > 60) continue; // lange Texte nur mit stärkerem Muster akzeptieren
    const dateIso = parseNumericDayHeaderText(text, now);
    if (dateIso) return deriveWeek(dateIso, text, WEEK_READ_SOURCES.numericDay);
  }

  // Nichts Datierbares gefunden -> bewusst NULL zurückgeben (kein "Work week"-Fehlgreifer).
  return {
    text: null, startDate: null, endDate: null, weekNumber: null,
    source: null,
    // FIX 1: der Grund wird BENANNT. "Kein angezeigter Header lesbar" und
    // "Header-Auswertung fehlgeschlagen" (collectError) sind sonst nicht
    // unterscheidbar — und der Aufrufer MUSS daraus headerReadError machen
    // statt ersatzweise den Button-Text zu nehmen.
    readError: 'no-displayed-header',
    navOnlyCandidates: headers.concat(navTexts).filter((t) => isWeekNavLabel(t)).slice(0, 5)
  };
}

// Liest die gerenderte Kalenderwoche aus der OWA-Wochenansicht aus.
// Rückgabe: { text, startDate, endDate, source, readError } — startDate/endDate
// als ISO-Strings 'YYYY-MM-DD', alles null wenn der Header nicht gefunden oder
// nicht parsebar war.
// `readError` unterscheidet die drei Nicht-Fälle:
//   - collectError:            das evaluate selbst ist fehlschlug (Code-Defekt)
//   - 'no-displayed-header':   gerendert, aber kein ANGEZEIGTER Header lesbar
//   - null:                    ein Header wurde gelesen
// (BUG 5, Live-Runde 4: vorher waren "Header noch nicht gerendert" und
// "Auswertung fehlgeschlagen" nicht unterscheidbar.)
// FIX 1 (Live-Lauf 5): `source` sagt, WOHER der Wert stammt. Nur
// 'displayed-header'/'day-column' sind ein gelesener Wochen-Header.
export async function readRenderedCalendarWeek(page) {
  const collected = await collectCalendarCandidates(page);
  const week = resolveRenderedWeekFromCandidates(collected);
  return collected && collected.collectError ? { ...week, readError: collected.collectError } : week;
}

// FIX 1: Ist der Wochenwert ein GELESENER angezeigter Header — oder eine
// Ableitung? Nur die erste Art ist ein Wochen-Nachweis. Aus Prev/Next-Buttons
// abgeleitete, aus Einzel-Datum geschätzte oder numerisch geschätzte Werte
// sind Schätzungen und dürfen kein 'weekVerified' erkaufen.
export function isHeaderReadVerified(rendered) {
  return Boolean(
    rendered && rendered.startDate && rendered.endDate && isDisplayedWeekSource(rendered.source)
  );
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
    ariaLabelSample: collected.ariaLabelSample,
    // BUG 5: 'collectError' present -> der Header wurde nicht "nicht gefunden",
    // sondern die Auswertung selbst ist gescheitert. Ohne dieses Feld sind beide
    // Faelle in der Diagnostik identisch (alle Kandidatenlisten leer).
    collectError: collected.collectError || null
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
    // BUG 5: 'Header-Auswertung FEHLGESCHLAGEN' ist eine andere Ursache als
    // 'keine Kandidaten' (= Header noch nicht gerendert). Ohne diese Zeile sind
    // beide in der Diagnostik nicht unterscheidbar.
    `Header-Auswertung: ${diag.collectError ? `FEHLGESCHLAGEN (${String(diag.collectError).slice(0, 120)})` : 'ok'}`,
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
// je max. 4s Re-Render-Wartezeit pro Klick.
//
// FIX 3 (Live-Lauf 5, 26.09.2026): 6 -> 2. Belege aus Lauf 5 Aufruf 1:
//   6 Klicks in der Folge prev/next/prev/next, dann abortReason 'max-clicks',
//   budget.usedMs 8706 — bei einer View, die die GANZE ZEIT korrekt auf KW 39
//   stand. Die Klicks haben die Woche nicht der Zielwoche naehergebracht, sie
//   haben sie nur hin und her gezogen.
//
// Fuer die Kalenderwochen-Navigation genuegen ZWO Klicks (eine Woche vor und
// eine Woche zurueck). Ist mehr noetig, gibt es den ehrlichen Abbruchgrund
// 'too-far' — die startdate-URL wird als EIN Versuch versucht (und in Lauf 5
// hat sie die View nicht bewegt), aber nicht als Beleg gewertet.
export const FREEBUSY_MAX_BUTTON_CLICKS = 2;
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

// BUG 5 (Live-Befund 26.09.2026, Lauf 4): Budget fuer das WARTEN auf einen
// datierbaren Wochen-Header (Zustand 1). Live war die Verweigerung des
// Wochen-Klicks korrekt, kam aber nach 13ms — es wurde auf eine Ansicht
// verzichtet, deren Header noch gar nicht gerendert war. Der Wert ist eine
// PHASE des Gesamtbudgets, kein Extra: phase() deckelt ihn gegen die Restzeit.
export const FREEBUSY_VIEW_WAIT_BUDGET_MS = 8000;

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
// Beobachtungsprotokoll je Poll (Bug 3; in Bug 4 auf datumsqualifiziert umgestellt):
//   datumsqualifizierter Beleg  -> 'found' (Header der Zielwoche ODER ein Label
//                                            aus dem angefragten Zeitraum)
//   loading === true            -> Spinner: NOCH NICHT FERTIG, weiter pollen
//   Events, alle ausserhalb des angefragten Zeitraums
//                               -> weder Treffer noch Leer-Beweis: weiter pollen
//                                  bis zum Budget, Abbruchgrund
//                                  'events-outside-requested-range'
//   0 Events, kein Spinner      -> Ruhe-Fenster sammeln; erst nach settleMs
//                                  ununterbrochen ruhig -> 'not-found' mit
//                                  abortReason 'rendered-empty-proven'
//   Budget erschoepft            -> 'budget-exhausted' mit abortReason
//                                     'wait-budget-exhausted' (Budget
//                                     tatsaechlich ausge schoepft)
// ---------------------------------------------------------------------------
// Bug 4 — der Warte-Loop ist WOCHENBLIND (Live-Befund 26.09.2026, Lauf 4)
// ---------------------------------------------------------------------------
// Live-Kennzahlen des Kontrolllaufs (Angefragt: KW 40 ab 2026-09-28):
//     eventWaitPolls: 1   waitedForEventsMs: 13   eventWaitBudgetMs: 20000
//     eventWaitOutcome: 'found'   abortReason: null
// 13 von 20000ms = 0,07 % des Budgets. Grund: `eventElementsFound > 0` war mit
// 18 Event-Knoten der FALSCHEN Woche (KW 39, 21.–26.09.) erfuellt. Die
// Abbruchbedingung der Wartephase war reine Knotenzahl — sie kann das Datum
// nicht unterscheiden. `eventElementsFound` ist die Laenge von `labels`
// (Kollektor), und der Range-Filter (`date-out-of-range') laeuft ERST danach in
// der Abschluss-Sammlung. Die Wartephase erfuhr die falsche Woche also nie.
//
// DIE REGEL: `found` ist erst belegt, wenn DATUMSQUALIFIZIERTER Beleg vorliegt:
//   (1) ein datierbarer Wochen-Header (startDate+endDate) der Zielwoche, ODER
//   (2) mindestens ein Label aus dem angefragten Zeitraum.
// Beide Belege werden WAEHREND des Wartens bei jedem Poll neu erhoben — der
// Kalender darf zwischen den Polls umschalten, ein einmaliger Vorab-Blick am
// Anfang waere genau der Fehler, den diese Regel abstellt.
// Ohne Zielwoche konfiguriert ist KEIN Beleg moeglich; dann wird 'found'
// ausdruecklich NICHT gemeldet ('wait-target-unknown'), damit ein
// weekenblind konfigurierter Aufruf nicht still Erfolg meldet.
export const EVENT_WAIT_ABORT_REASONS = {
  renderedEmpty: 'rendered-empty-proven',
  budgetExhausted: 'wait-budget-exhausted',
  notPolled: 'wait-not-polled',
  budgetUnavailable: 'wait-budget-unavailable',  // BUG 4: Events sichtbar, aber keins aus dem angefragten Zeitraum. Das ist
  // KEIN Leer-Beweis (der Leer-Beweis bleibt 'rendered-empty-proven') und auch
  // kein Treffer — die Ansicht zeigt dann etwas anderes als die Zielwoche.
  eventsOutsideRange: 'events-outside-requested-range',
  // BUG 4: Die Zielwoche wurde datumsqualifiziert belegt (Header gerendert),
  // aber es wurde kein einziger Event-Knoten gelesen. Das ist eine fehlende
  // Lesbarkeit des Grids, KEIN freier Kalender.
  noNodesInTargetWeek: 'no-event-nodes-in-target-week',
  // BUG 4: Ohne konfigurierten Zielzeitraum ist keine Qualifikation moeglich.
  targetUnknown: 'wait-target-unknown',
  // FIX 2 (Live-Lauf 5): ein DATIERTER Header der FALSCHEN Woche. Das ist ein
  // GEGENBELEG, kein Beleg — und es ist auch kein Leer-Beweis: die Ansicht
  // zeigt nachweislich eine andere Woche als die angefragte. In Lauf 5 war es
  // das, was nach FIX 1 (Header = angezeigte Woche) uebrig bleibt: 10 echte
  // Event-Knoten, alle ausserhalb des Zeitraums, und ein Header, der KW 39
  // statt KW 40 sagt. Vorher fiel das unter 'events-outside-requested-range'
  // bzw. - bei 0 Knoten - sogar unter den Leer-Beweis 'rendered-empty-proven'.
  headerOutsideRange: 'header-outside-requested-range'
};

// Die fuenf Beobachtungs-Kategorien der Wartephase. Genau EINE Klassifikation,
// damit "noch nicht fertig", "nachweislich leer", "falsche Woche" und
// "datumsqualifiziert fertig" nicht wieder an mehreren Stellen auseinanderlaufen
// (das war der Live-Befund: 'found' nach einem Poll, weil NUR die Zahl zaehlte).
export const EVENT_WAIT_STATES = {
  targetWeek: 'target-week-rendered',
  inRange: 'in-range-labels',
  stillLoading: 'still-loading',
  outsideRange: 'events-outside-requested-range',
  emptyCandidate: 'empty-candidate',
  targetUnknown: 'target-unknown',
  // FIX 2: datierter Header, der die Zielwoche NICHT enthaelt. Vor der
  // Event-Zahl, weil der Header die staerkere Information traegt.
  headerOutsideRange: 'header-outside-requested-range'
};

// Zaehlt Labels, die den angefragten Zeitraum treffen. Bewusst ueber DIESELBE
// Prueflogik wie die Abschluss-Extraktion (parseEventLabel + partitionSlotsByRange):
// wenn die Wartephase eine andere Definition benutzt als die Extraktion, kann
// sie "fertig" melden, fuer die es am Ende keine Slots gibt.
export function countLabelsInRange(labels = [], { rangeStartMs = null, rangeEndMs = null, defaultYear = null } = {}) {
  const parsed = [];
  let undated = 0;
  for (const label of labels || []) {
    const slot = parseEventLabel(label, defaultYear);
    if (!slot) { undated++; continue; }
    parsed.push(slot);
  }
  const measurable = Number.isFinite(rangeStartMs) && Number.isFinite(rangeEndMs);
  if (!measurable) return { inRange: 0, outOfRange: 0, undated, unmeasured: true };
  const { kept, rejected } = partitionSlotsByRange(parsed, rangeStartMs, rangeEndMs);
  let outOfRange = 0;
  for (const r of rejected) if (r.reason === 'date-out-of-range') outOfRange++;
  return { inRange: kept.length, outOfRange, undated: parsed.length - kept.length - outOfRange, unmeasured: false };
}

// Klassifiziert EINE Beobachtung der Wartephase. Reine Funktion, EIN Aufrufort.
// `weekHeader` ist das Ergebnis des Wochen-Reads, `inRangeLabels` die Zahl der
// Labels aus dem angefragten Zeitraum.
//
// FIX 2 (Live-Lauf 5): 'found' ist BEREICHS-qualifiziert, nicht nur
// DATUMS-qualifiziert. Live: eventWaitDateQualified:true bei
// eventWaitInRangeLabels:0 -> found nach 41ms von 20000ms Budget. Ein
// datierter Header der FALSCHEN Woche ist ein GEGENBELEG ("die Ansicht zeigt
// nachweislich eine andere Woche"), kein Beleg. Er bekommt deshalb einen
// eigenen, benannten Zustand (header-outside-requested-range) und loest
// weder 'found' noch den Leer-Beweis aus.
export function eventWaitObservationState({
  eventElementsFound = 0,
  loading = false,
  weekHeader = null,
  inRangeLabels = 0,
  targetStart = null,
  targetConfigured = false
} = {}) {
  const n = Number(eventElementsFound) || 0;
  const inRange = Number(inRangeLabels) || 0;
  // Ohne Zielzeitraum ist KEIN datumsqualifizierter Beleg moeglich — auch dann
  // nicht, wenn Labels vorhanden sind: "im angefragten Zeitraum" ist ohne
  // Zielzeitraum nicht definierbar. Diese Pruefung steht VOR den Beleg-
  // Kriterien, sonst wuerde ein weekenblind konfigurierter Aufruf doch noch
  // 'found' melden.
  if (!targetConfigured) {
    return { category: EVENT_WAIT_STATES.targetUnknown, dateQualified: false, evidence: null, eventElementsFound: n, inRangeLabels: inRange, loading: Boolean(loading) };
  }
  // Beleg 1: der ANGEZEIGTE Header (nicht ein Nav-Button, vgl. FIX 1) traegt
  // die Zielwoche. renderedWeekContains prueft den GANZEN Range, nicht nur den
  // Start — ein Header, der die Zielwoche nur teilweise ueberlappt, gilt nicht
  // als Beleg fuer die angefragte Woche.
  if (renderedWeekContains(weekHeader, targetStart)) {
    return { category: EVENT_WAIT_STATES.targetWeek, dateQualified: true, evidence: 'week-header', eventElementsFound: n, inRangeLabels: inRange, loading: Boolean(loading) };
  }
  // Beleg 2: mindestens ein Label aus dem angefragten Zeitraum.
  if (inRange > 0) {
    return { category: EVENT_WAIT_STATES.inRange, dateQualified: true, evidence: 'in-range-label', eventElementsFound: n, inRangeLabels: inRange, loading: Boolean(loading) };
  }
  if (loading) {
    return { category: EVENT_WAIT_STATES.stillLoading, dateQualified: false, evidence: null, eventElementsFound: n, inRangeLabels: inRange, loading: true };
  }
  // FIX 2: Header ist DATIERT, aber die Zielwoche ist nicht drin. VOR der
  // Event-Zahl, weil der Header die staerkere Information traegt: er sagt
  // nicht nur "nichts passt", sondern "es ist die falsche Woche".
  if (weekHeader && weekHeader.startDate && weekHeader.endDate) {
    return { category: EVENT_WAIT_STATES.headerOutsideRange, dateQualified: false, evidence: null, eventElementsFound: n, inRangeLabels: inRange, loading: false, renderedWeek: `${weekHeader.startDate}–${weekHeader.endDate}` };
  }
  if (n > 0) {
    // Live-Fall Runde 4: 18 Knoten, alle ausserhalb des angefragten Zeitraums.
    return { category: EVENT_WAIT_STATES.outsideRange, dateQualified: false, evidence: null, eventElementsFound: n, inRangeLabels: inRange, loading: false };
  }
  return { category: EVENT_WAIT_STATES.emptyCandidate, dateQualified: false, evidence: null, eventElementsFound: n, inRangeLabels: inRange, loading: false };
}

export async function waitForCalendarEvents({
  observe,
  budgetMs,
  pollMs = FREEBUSY_EVENT_WAIT_POLL_MS,
  settleMs = FREEBUSY_EVENT_WAIT_SETTLE_MS,
  targetStart = null,
  now = () => Date.now(),
  sleep
} = {}) {
  const startedAt = now();
  const waited = () => now() - startedAt;
  const targetConfigured = Boolean(targetStart);
  const result = (extra) => ({
    found: false,
    outcome: 'not-found',
    waitedMs: waited(),
    polls: 0,
    loadingPolls: 0,
    outsideRangePolls: 0,
    // FIX 2: Polls mit datiertem Header der FALSCHEN Woche.
    headerOutsideRangePolls: 0,
    loading: false,
    dateQualified: false,
    abortReason: null,
    ...extra
  });
  if (typeof observe !== 'function' || !budgetMs || budgetMs <= 0) {
    // Kein Budget -> ehrlich benennen statt "0 Polls bei 20s Budget" zu melden.
    return result({ outcome: 'budget-exhausted', abortReason: EVENT_WAIT_ABORT_REASONS.budgetUnavailable });
  }
  if (!targetConfigured) {
    // BUG 4: Ohne Zielwoche kann diese Phase nichts datumsqualifizieren. Sie
    // pollt deshalb nicht 20s auf eine Beleglosigkeit, sondern benennt es.
    return result({ outcome: 'not-found', abortReason: EVENT_WAIT_ABORT_REASONS.targetUnknown });
  }
  const doSleep = typeof sleep === 'function' ? sleep : (ms) => new Promise((r) => setTimeout(r, ms));
  let polls = 0;
  let loadingPolls = 0;
  let outsideRangePolls = 0;
  // FIX 2: "Header der falschen Woche" ist ein eigener Zaehler — sichtbar
  // getrennt von "Events ausserhalb des Zeitraums", weil es andere Ursache
  // UND andere Konsequenz hat (die Woche ist bekannt und falsch).
  let headerOutsideRangePolls = 0;
  let quietSince = null;   // Beginn des ununterbrochenen Ruhe-Fensters
  let lastLoading = false;
  let lastState = EVENT_WAIT_STATES.emptyCandidate;
  while (waited() < budgetMs) {
    polls++;
    // eslint-disable-next-line no-await-in-loop
    const o = (await observe()) || {};
    const st = eventWaitObservationState({ ...o, targetStart, targetConfigured });
    lastState = st.category;
    lastLoading = st.loading === true;
    if (st.dateQualified) {
      return {
        found: true, outcome: 'found', waitedMs: waited(), polls, loadingPolls, outsideRangePolls,
        headerOutsideRangePolls, loading: lastLoading, dateQualified: true, evidence: st.evidence,
        weekHeader: o.weekHeader || null, inRangeLabels: st.inRangeLabels, abortReason: null
      };
    }
    if (lastLoading) {
      // BEWEIS FUER "NOCH NICHT FERTIG": der Spinner haelt den Poll am Leben.
      loadingPolls++;
      quietSince = null;
    } else if (st.category === EVENT_WAIT_STATES.headerOutsideRange) {
      // FIX 2: datierter Header der falschen Woche. startet KEIN Ruhe-Fenster
      // (es ist ein GEGENBELEG, kein Leer-Beweis) und invalidiert ein
      // bestehendes. Es wird bis zum Budget weiter gepollt — die Zielwoche
      // kann noch umschalten (Klick/Navigation/Rendering).
      headerOutsideRangePolls++;
      quietSince = null;
    } else if (st.category === EVENT_WAIT_STATES.outsideRange) {
      // BUG 4: Events sind da, aber keine aus dem angefragten Zeitraum. Das
      // startet KEIN Ruhe-Fenster — ein Leer-Beweis waere hier eine Falschaussage.
      // Genauso wird ein bestehendes Ruhe-Fenster invalidiert, weil sich das DOM
      // nachweislich geaendert hat. Es wird bis zum Budget weiter gepollt.
      outsideRangePolls++;
      quietSince = null;
    } else if (st.category === EVENT_WAIT_STATES.emptyCandidate) {
      if (quietSince === null) quietSince = now();
      const quietMs = waited() - (quietSince - startedAt);
      if (quietMs >= (settleMs || 0)) {
        return {
          found: false, outcome: 'not-found', waitedMs: waited(), polls, loadingPolls, outsideRangePolls,
          headerOutsideRangePolls, loading: false, dateQualified: false, weekHeader: o.weekHeader || null,
          abortReason: EVENT_WAIT_ABORT_REASONS.renderedEmpty
        };
      }
    } else {
      quietSince = null;
    }
    if (waited() + pollMs >= budgetMs) break;
    // eslint-disable-next-line no-await-in-loop
    await doSleep(pollMs);
  }
  return {
    found: false, outcome: 'budget-exhausted', waitedMs: waited(), polls, loadingPolls, outsideRangePolls,
    headerOutsideRangePolls, loading: lastLoading, dateQualified: false, lastState,
    // BUG 4/FIX 2: der Abbruchgrund nennt den zuletzt beobachteten Zustand,
    // damit "Budget erschoepft, weil die falsche Woche stand" (Header
    // datiert) von "Budget erschoepft, weil Events ausserhalb standen" und von
    // "Budget erschoepft, weil gar nichts kam" unterscheidbar bleibt.
    abortReason: lastState === EVENT_WAIT_STATES.headerOutsideRange
      ? EVENT_WAIT_ABORT_REASONS.headerOutsideRange
      : lastState === EVENT_WAIT_STATES.outsideRange
        ? EVENT_WAIT_ABORT_REASONS.eventsOutsideRange
        : EVENT_WAIT_ABORT_REASONS.budgetExhausted
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
//
// BUG 4: Die Beobachtung ist nicht mehr "wie viele Knoten", sondern
//   { eventElementsFound, loading, weekHeader, inRangeLabels }.
// `readWeek()` liest den gerenderten Wochen-Header; `range` definiert den
// angefragten Zeitraum (in ms, wie in getFreeBusy). OHNE `range` bleibt
// `inRangeLabels` null (unmeasured) — die Wartephase meldet dann bewusst
// 'wait-target-unknown' und NIEMALS 'found'. Ein weekenblind gebauter Beobachter
// ist damit nicht mehr moeglich, ohne dass es benannt wird.
//
// Der Header-Read ist zeitlich gedrosselt (headerReadIntervalMs): readWeek geht
// ueber collectCalendarCandidates und liest document.body.innerText. Bei 500ms
// Pollintervall und 20s Budget waeren das 40 teure Reads; gedrosselt wird der
// Wochenzustand trotzdem bei JEDEM Poll geprueft, an dem er die Entscheidung
// ueberhaupt aendern kann (wenn inRangeLabels > 0, ist ohnehin entschieden).
export function makeEventObserver(collect, {
  readWeek = null,
  range = null,
  defaultYear = null,
  headerReadIntervalMs = FREEBUSY_EVENT_WAIT_POLL_MS * 2,
  now = () => Date.now()
} = {}) {
  if (typeof collect !== 'function') throw new TypeError('makeEventObserver: collect muss eine Funktion sein');
  let lastSelector = null;
  let lastHeaderReadAt = -Infinity;
  let lastHeader = null;
  const rangeStartMs = range && Number.isFinite(range.rangeStartMs) ? range.rangeStartMs : null;
  const rangeEndMs = range && Number.isFinite(range.rangeEndMs) ? range.rangeEndMs : null;
  const readWeekThrottled = async () => {
    if (typeof readWeek !== 'function') return null;
    const t = now();
    if (t - lastHeaderReadAt < (headerReadIntervalMs || 0)) return lastHeader;
    lastHeaderReadAt = t;
    try { lastHeader = await readWeek(); } catch (e) { lastHeader = null; }
    return lastHeader;
  };
  const observeCalendarEvents = async () => {
    // eslint-disable-next-line no-await-in-loop
    const r = await collect(lastSelector) || {};
    const n = Number(r.eventElementsFound) || 0;
    if (n > 0 && r.selectorUsed) lastSelector = r.selectorUsed;
    // In-Range-Zaehlung mit DERSELBEN Prueflogik wie die Abschluss-Extraktion.
    const counts = countLabelsInRange(r.labels || [], { rangeStartMs, rangeEndMs, defaultYear });
    // Der Wochen-Read wird nur gebraucht, wenn die Labels allein nicht
    // qualifizieren koennen — und dann gedrosselt.
    const weekHeader = counts.inRange > 0 ? null : await readWeekThrottled();
    return {
      eventElementsFound: n,
      loading: r.loading === true,
      weekHeader,
      inRangeLabels: counts.unmeasured ? null : counts.inRange,
      labelsOutOfRange: counts.unmeasured ? null : counts.outOfRange,
      labelsUndated: counts.undated
    };
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
    // ABER: BUG 4 — wenn die Wartephase gar nicht starten durfte, weil kein
    // Zielzeitraum konfiguriert war, ist 'not-polled' eine Falschaussage (es
    // wurde nicht "nicht gepollt", es war nicht pollbar). Der echte Grund der
    // Wartephase steht dann in w.abortReason und wird zuerst genannt.
    if (wait && wait.abortReason === EVENT_WAIT_ABORT_REASONS.targetUnknown) {
      return { outcome: 'not-found', abortReason: EVENT_WAIT_ABORT_REASONS.targetUnknown };
    }
    return { outcome: 'not-polled', abortReason: EVENT_WAIT_ABORT_REASONS.notPolled };
  }
  const w = wait || {};
  // BUG 4: Die Wartephase endet datumsqualifiziert (z.B. der Wochen-Header der
  // Zielwoche war gerendert), die Abschluss-Sammlung las aber 0 Knoten. Das ist
  // ein GEMESSENER Zustand, kein Logikfehler und kein Leer-Beweis: der Header
  // war da, das Grid war nicht lesbar. 'budget-exhausted' waere hier eine
  // Falschaussage (das Budget wurde gar nicht ausgeschoepft).
  if (w.dateQualified === true) {
    return { outcome: 'found', abortReason: EVENT_WAIT_ABORT_REASONS.noNodesInTargetWeek };
  }
  if (w.abortReason === EVENT_WAIT_ABORT_REASONS.renderedEmpty) {
    return { outcome: 'not-found', abortReason: EVENT_WAIT_ABORT_REASONS.renderedEmpty };
  }
  if (w.abortReason === EVENT_WAIT_ABORT_REASONS.budgetUnavailable) {
    return { outcome: 'budget-exhausted', abortReason: EVENT_WAIT_ABORT_REASONS.budgetUnavailable };
  }
  if (w.abortReason === EVENT_WAIT_ABORT_REASONS.targetUnknown) {
    return { outcome: 'not-found', abortReason: EVENT_WAIT_ABORT_REASONS.targetUnknown };
  }
  if (w.abortReason === EVENT_WAIT_ABORT_REASONS.eventsOutsideRange) {
    return { outcome: 'budget-exhausted', abortReason: EVENT_WAIT_ABORT_REASONS.eventsOutsideRange };
  }
  // FIX 2: der datierte Header der falschen Woche ist ein eigener Grund und
  // wird nicht auf den generischen Budget-Abbruch abgebildet — sonst waere er
  // im Ergebnis nicht mehr von "es kam gar nichts" unterscheidbar.
  if (w.abortReason === EVENT_WAIT_ABORT_REASONS.headerOutsideRange) {
    return { outcome: 'budget-exhausted', abortReason: EVENT_WAIT_ABORT_REASONS.headerOutsideRange };
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

// ---------------------------------------------------------------------------
// BUG 5 — die drei Zustaende der Wochenansicht (Live-Befund 26.09.2026, Lauf 4)
// ---------------------------------------------------------------------------
// Live-Befund Runde 4 (belegt, nicht vermutet):
//   urlAfterNavigation: .../calendar/view/week?...&startdate=2026-09-28
//   renderedWeek: { text: null, startDate: null, endDate: null }
//   headerCandidates: []  dayHeaderCandidates: []  weekNavCandidates: []
//   census.dayColumns: []   weekVerified: false
//   weekButtonProbe: { skipped: 'view-not-ready', clicked: false,
//                       note: 'guard-vor-klick' }   abortReason: 'view-not-ready'
//   -> 0,07 % des Event-Wartebudgets verbraucht, 1 Poll.
// Der Inhalt war KW 39, also die Woche VOR der Zielwoche. Der Guard hat den
// Klick korrekt verweigert — und genau damit die einzige Navigation verhindert,
// die zur Zielwoche gefuehrt haette. Das Werkzeug verweigerte, wo es handeln
// musste.
//
// URSACHE, an der Quelle belegt (nicht geraten):
//   1) waitForCalendarRendered() hatte im Zweig `else if (events) break;`:
//      "kein Header, aber Body-Events" galt als FERTIG. Genau das ist der
//      Live-Zustand (18 Event-Knoten, 0 Header) — die Ansicht wurde also nach
//      dem ersten Poll (~500ms) verlassen, long bevor der Header gerendert
//      sein konnte. Beleg: waitedForEventsMs 13 / polls 1 bei Budget 20000.
//   2) Der Guard selbst war richtig: ohne Header ist die Wochenlage UNBEKANNT.
//      Er wurde nur viel zu frue gefragt.
// Also: die Wochen-Navigation ist nicht als Defekt belegt (Runde 3 hat per
// Klick auf "Next week" eine datierbare Zielwoche erreicht, und in Runde 3
// existierten die Nav-Indikatoren "… (Week 39)"/"… (Week 41)" — in Runde 4
// waren sie nicht da, passend zu "Header-Bereich noch nicht gerendert").
// Der Fix ist deshalb eine ZUSTANDSKORREKTUR, kein Reparaturversuch an der
// Navigation: Zustand 1 wird abgeWARTET (mit Budget und Spinner-Respekt),
// Zustand 3 darf klicken — mit anschliessender Verifikation des Headers.
export const WEEK_VIEW_STATES = {
  // Ansicht bereit, Zielwoche gerendert -> fertig, KEIN Klick.
  targetWeek: 'target-week-reached',
  // Ansicht bereit, aber falsche Woche -> Wochen-Navigation Klick ist erlaubt.
  wrongWeek: 'wrong-week',
  // Ansicht nicht bereit (kein datierbarer Header) -> WARTEN, nicht klicken.
  notReady: 'view-not-ready'
};

// DIE Zustandsklassifikation. Einzige Quelle fuer "welcher Zustand gilt" —
// benutzt von der aeusseren Schritt-Schleife UND vom Klickpfad (ueber
// weekClickDecision). Bewusst OHNE die alte hasEvents-Ausnahme: die Annahme
// "sichtbare Events belegen eine gerenderte Ansicht" hat den Live-Befund
// ueberlebt, obwohl sie im Klickpfad nie wirksam war (dort wurde hasEvents gar
// nicht durchgereicht) — sie steht damit fuer einen Zustand, den es nicht gibt.
// Ohne datierbaren Header ist die Wochen- und damit die BUTTON-Position
// unbekannt: die Wochen-Buttons liegen im Header-Bereich, der genau dann noch
// nicht da ist. Klicken kann dort nichts treffen und im schlechtesten Fall die
// Zielwoche ueberspringen.
export function weekViewState({ rendered = null, targetStart = null } = {}) {
  const headerRendered = Boolean(rendered && rendered.startDate && rendered.endDate);
  if (!headerRendered) {
    return { state: WEEK_VIEW_STATES.notReady, ready: false, week: rendered || null, clickable: false };
  }
  if (renderedWeekContains(rendered, targetStart)) {
    return { state: WEEK_VIEW_STATES.targetWeek, ready: true, week: rendered || null, clickable: false };
  }
  return { state: WEEK_VIEW_STATES.wrongWeek, ready: true, week: rendered || null, clickable: true };
}

// EIN Entscheidungs-Funktion, benutzt von der aeusseren Schritt-Schleife UND
// vom Klickpfad. Beide Stellen konnten vorher auseinanderlaufen — genau das war
// der Live-Befund. Reine Funktion, damit der Draht testbar ist. Sie ist eine
// Abbildung von weekViewState: EIN Zustand, EIN Mapping, keine zweite Wahrheit.
export function weekClickDecision({ rendered, targetStart } = {}) {
  const view = weekViewState({ rendered, targetStart });
  if (view.state === WEEK_VIEW_STATES.targetWeek) {
    return { skip: true, reason: WEEK_CLICK_GUARD_REASONS.targetReached, state: view.state, rendered: view.week, clickable: false };
  }
  if (view.state === WEEK_VIEW_STATES.notReady) {
    return { skip: true, reason: WEEK_CLICK_GUARD_REASONS.viewNotReady, state: view.state, rendered: view.week, clickable: false };
  }
  return { skip: false, reason: WEEK_CLICK_GUARD_REASONS.click, state: view.state, rendered: view.week, clickable: true };
}

// Protokolleintrag fuer eine Guard-Entscheidung. BEIDE Ausgaenge werden
// protokolliert — ein Guard, der nichts aufschreibt, ist nicht von einem
// Guard, der nicht existiert, unterscheidbar. Genau daran ist der Live-Befund
// gescheitert: `note:'guard-vor-klick'` hätte sofort verraten, dass der Guard
// nicht griff.
export function weekProbeGuardEntry(decision, note = 'guard-vor-klick', waitedMs = null) {
  const renderedStart = decision && decision.rendered ? decision.rendered.startDate : null;
  if (decision && decision.reason === WEEK_CLICK_GUARD_REASONS.viewNotReady) {
    return {
      skipped: 'view-not-ready',
      renderedWeek: renderedStart || null,
      clicked: false,
      guard: WEEK_CLICK_GUARD_REASONS.viewNotReady,
      note,
      // BUG 5: unterscheidet "sofort verweigert" von "erst nach dem
      // Wartebudget verweigert". Im Live-Fall war die Verweigerung korrekt,
      // sie kam aber nach 13ms — also viel zu frue.
      ...(Number.isFinite(waitedMs) ? { waitedMs: Math.round(waitedMs) } : {})
    };
  }
  return { ...weekProbeSkipEntry(renderedStart, note), guard: WEEK_CLICK_GUARD_REASONS.targetReached };
}

// ---------------------------------------------------------------------------
// Der ECHTE Klickpfad — exportiert und per Dependency-Injection testbar
// ---------------------------------------------------------------------------
// Zustand 1: WARTEN statt klicken — die Wartefunktion fuer den fehlenden
// Wochen-Header. Exportiert und per Dependency-Injection testbar, weil sie der
// Pfad ist, der in Runde 4 viel zu frue verlassen wurde.
// ---------------------------------------------------------------------------
// Kernpunkt: der Wochenzustand wird bei JEDEM Poll neu gelesen. Der Kalender
// darf zwischen den Polls umschalten — "warte bis der Header da ist" heisst
// "pruefe waehrend des Wartens", nicht "einmal am Anfang schauen". Ein Spinner
// (loading) verhindert das Ruhe-Fenster genauso wie in der Event-Wartephase:
// "noch nicht fertig" ist ein eigener Zustand, kein Leer-Beweis.
//
// Rueckgabe: { state, rendered, polls, waitedMs, exhausted }
//   state === 'view-not-ready' && exhausted === true -> ehrlicher Abbruch
export async function waitForWeekViewReady({
  readWeek = null,
  sleep = null,
  budgetMs = 0,
  pollMs = FREEBUSY_EVENT_WAIT_POLL_MS,
  settleMs = FREEBUSY_EVENT_WAIT_SETTLE_MS,
  readEvents = null,
  targetStart = null,
  now = () => Date.now()
} = {}) {
  const startedAt = now();
  const waited = () => now() - startedAt;
  const doSleep = typeof sleep === 'function' ? sleep : (ms) => new Promise((r) => setTimeout(r, ms));
  let polls = 0;
  let loadingPolls = 0;
  let quietSince = null;
  let last = { text: null, startDate: null, endDate: null };
  if (typeof readWeek !== 'function' || !(budgetMs > 0)) {
    return { state: WEEK_VIEW_STATES.notReady, rendered: last, polls: 0, loadingPolls: 0, waitedMs: waited(), exhausted: true };
  }
  while (waited() < budgetMs) {
    polls++;
    // eslint-disable-next-line no-await-in-loop
    last = (await readWeek()) || last;
    const state = weekViewState({ rendered: last, targetStart });
    if (state.state !== WEEK_VIEW_STATES.notReady) {
      return { state: state.state, rendered: last, polls, loadingPolls, waitedMs: waited(), exhausted: false };
    }
    // eslint-disable-next-line no-await-in-loop
    const stillLoading = typeof readEvents === 'function' ? ((await readEvents()) || {}).loading === true : false;
    if (stillLoading) {
      // BEWEIS FUER "NOCH NICHT FERTIG" — der Spinner haelt den Poll am Leben
      // und macht ein bestehendes Ruhe-Fenster ungueltig.
      loadingPolls++;
      quietSince = null;
    } else {
      if (quietSince === null) quietSince = now();
      const quietMs = waited() - (quietSince - startedAt);
      if (quietMs >= (settleMs || 0)) {
        return { state: WEEK_VIEW_STATES.notReady, rendered: last, polls, loadingPolls, waitedMs: waited(), exhausted: true, settled: true };
      }
    }
    if (waited() + pollMs >= budgetMs) break;
    // eslint-disable-next-line no-await-in-loop
    await doSleep(pollMs);
  }
  return { state: WEEK_VIEW_STATES.notReady, rendered: last, polls, loadingPolls, waitedMs: waited(), exhausted: true };
}

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
  afterClick = null,
  clickWaitMs = FREEBUSY_CLICK_RENDER_WAIT_MS,
  clickCostMs = FREEBUSY_CLICK_RENDER_WAIT_MS,
  remainingMs = () => Infinity,
  waitForReady = null,
  onGuardBlocked = null
} = {}) {
  const list = Array.isArray(candidates) && candidates.length ? candidates : weekButtonCandidates(direction);
  const notReached = (reason, rendered) => ({
    clicked: false, week: rendered || null, verified: 'guard-blocked', foundAny: false, clickedAny: false,
    targetReached: reason === WEEK_CLICK_GUARD_REASONS.targetReached,
    effect: CLICK_EFFECTS.unverified,
    guard: reason
  });

  // Der Header-Read ist memoisiert: nur der erste Kandidat kostet einen Read.
  // Nach jedem Klick wird er verworfen, weil erst der Klick die Woche aendern kann.
  let targetWeekChecked = false;
  let targetWeekNow = null;
  let waitedMs = 0;
  const guardNow = async () => {
    if (!targetWeekChecked) {
      targetWeekChecked = true;
      let wk = null;
      try { wk = await readWeek(); } catch (e) { wk = null; }
      targetWeekNow = weekClickDecision({ rendered: wk, targetStart });
    }
    return targetWeekNow;
  };

  // BUG 5, Zustand 1: WARTEN statt klicken. Wird nur gebraucht, wenn der Guard
  // "Ansicht nicht bereit" sagt; ein Guard, der die Zielwoche bestaetigt, wartet
  // nie. Die gelieferte Wartephase wird in waitedMs summiert und im
  // Probe-Eintrag ausgewiesen — damit ist im Live-Fall sichtbar, ob der Klick
  // nach 13ms oder nach dem Wartebudget verweigert wurde.
  const waitOnce = async () => {
    if (typeof waitForReady !== 'function') return null;
    const w = await waitForReady();
    if (w && Number.isFinite(Number(w.waitedMs))) waitedMs += Number(w.waitedMs);
    // Der memoisierte Guard-Zustand ist nach dem Warten veraltet: der Header
    // kann jetzt datierbar sein.
    targetWeekChecked = false;
    targetWeekNow = null;
    return w || null;
  };

  let foundAny = false;
  let clickedAny = false;

  for (const cand of list) {
    // GUARD VOR DER PROBE — fuer JEDE Entscheidung wird protokolliert.
    let decision = await guardNow();
    if (decision.reason === WEEK_CLICK_GUARD_REASONS.viewNotReady) {
      // Zustand 1: erst warten, dann neu entscheiden. Der Kalender kann
      // waehrend des Wartens umschalten — deshalb wird der Wochenzustand nach
      // dem Warten frisch gelesen, nicht aus dem alten Urteil uebernommen.
      await waitOnce();
      decision = await guardNow();
    }
    if (decision.skip) {
      probe.push(weekProbeGuardEntry(decision, 'guard-vor-klick', waitedMs));
      if (typeof onGuardBlocked === 'function') {
        onGuardBlocked({ reason: decision.reason, waitedMs: Math.round(waitedMs), polls: probe.length });
      }
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
    // FIX 3: die Annahme des Klicks wird am ANGEZEIGTEN Header bewertet
    // (classifyClickEffect), nicht daran festgemacht, dass sich ueberhaupt
    // etwas bewegt hat ('changed'). Der Aufrufer entscheidet danach ueber
    // advanced/overshoot/unverified — er KANN nicht mehr blind weiterklicken.
    const effect = classifyClickEffect({ prevStart, newWeek: r && r.week, targetStart, direction });
    probe.push({ selector: cand.selector, kind: cand.kind, found: true, clicked: true, verified: r.verified, effect });
    if (shouldTryNextCandidate(r.verified)) continue; // wirkungslos -> naechster Kandidat
    return {
      clicked: true, week: r.week, verified: r.verified, effect, foundAny, clickedAny,
      targetReached: effect === CLICK_EFFECTS.reached, guard: WEEK_CLICK_GUARD_REASONS.click
    };
  }
  return {
    clicked: false, week: null, verified: 'not-found', effect: CLICK_EFFECTS.unchanged,
    foundAny, clickedAny, targetReached: false, guard: null
  };
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
// BUG 4: `abortReason` ist das fuehrende Unterscheidungsmerkmal — die
// Abbruchgruende sind seit dem datumsqualifizierten Umbau praeziser als die
// Ausgaenge, und zwei davon (falsche Woche sichtbar / Zielwoche ohne Knoten)
// fallen sonst in die generischen Texte fuer "nichts gefunden".
export function eventWaitWarning({ outcome, waitedMs = 0, budgetMs = null, abortReason = null, renderedWeek = null } = {}) {
  if (abortReason === EVENT_WAIT_ABORT_REASONS.headerOutsideRange) {
    return `Nach ${Math.round(waitedMs)}ms Wartephase (Budget ${budgetMs === null ? 'n/a' : budgetMs}ms) stand ein ` +
      `DATIERTER Wochen-Header im DOM, der die angefragte Woche NICHT enthaelt` +
      `${renderedWeek ? ` (Header: ${renderedWeek})` : ''}. ` +
      `Das ist ein Gegenbeleg, kein Beleg: die Ansicht zeigt nachweislich eine andere Woche. ` +
      `Weder Treffer noch Leer-Beweis — bis zum Budget weitergepollt. ` +
      `Siehe diagnostics.renderedWeek/weekHeaderLabels.`;
  }
  if (abortReason === EVENT_WAIT_ABORT_REASONS.eventsOutsideRange) {
    return `Nach ${Math.round(waitedMs)}ms Wartephase (Budget ${budgetMs === null ? 'n/a' : budgetMs}ms) waren ` +
      `Event-Knoten im DOM sichtbar, aber KEINER aus dem angefragten Zeitraum. ` +
      `Die Ansicht zeigt damit etwas anderes als die angefragte Woche — das ist weder ein Treffer ` +
      `noch ein Leer-Beweis. Siehe diagnostics.renderedWeek/weekNavCandidates.`;
  }
  if (abortReason === EVENT_WAIT_ABORT_REASONS.noNodesInTargetWeek) {
    return `Der Wochen-Header der angefragten Woche war datumsqualifiziert belegt, aber im DOM wurde ` +
      `KEIN einziger Event-Knoten gelesen. 0 Busy-Slots bedeuten damit KEINEN leeren Kalender, ` +
      `sondern ein nicht auslesbares Grid der richtigen Woche (siehe diagnostics.census/sourceTried).`;
  }
  if (abortReason === EVENT_WAIT_ABORT_REASONS.targetUnknown) {
    return `Die Event-Wartephase lief ohne konfigurierten Zielzeitraum und konnte deshalb keinen ` +
      `datumsqualifizierten Beleg erbringen. Sie wurde NICHT gepollt — das ist ein ` +
      `Konfigurationsfehler des Aufrufers, kein DOM-Zustand.`;
  }
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

// ---------------------------------------------------------------------------
// FIX 3 (Live-Lauf 5): die Navigation war ungerichtet, unverifiziert und
// unbegrenzt.
// ---------------------------------------------------------------------------
// Lauf 5 Aufruf 1, Live-Belege:
//   - 6 Klicks in der Folge prev/next/prev/next, dann abortReason 'max-clicks',
//     budget.usedMs 8706 — obwohl die View die GANZE ZEIT korrekt auf KW 39 stand
//     und der angefragte Range 21.09.–02.10. KW 39 EINSCHLIESST.
//   - Die Verifikation stuetzte sich auf 'changed' ("hat sich irgendetwas
//     bewegt"), nicht darauf, ob sich die Woche der Zielwoche NAEHERT.
//
// 'changed' ist die schwächste mögliche Aussage. Notwendig ist die
// WIRKUNG in Bezug auf das Ziel: hat sich die angezeigte Woche in
// Klickrichtung um eine Woche bewegt, steht sie jetzt wochenrichtig, oder hat
// sie die Zielwoche uebersprungen?
export const CLICK_EFFECTS = {
  // Vor oder nach dem Klick war kein datierter Header lesbar -> nicht bewertbar.
  unverified: 'unverified',
  // Header lesbar, Startdatum unveraendert -> der Klick hat nichts bewegt.
  unchanged: 'unchanged',
  // Die Zielwoche ist jetzt angezeigt.
  reached: 'reached',
  // Der Klick hat die Woche in Klickrichtung um eine Woche bewegt.
  advanced: 'advanced',
  // Der Klick hat die Woche in die GEGENRICHTUNG bewegt (oder mehr als eine
  // Woche gesprungen): die Zielwoche wurde uebersprungen.
  overshoot: 'overshoot'
};

// Wirkung EINES Klicks, bewertet am ANGEZEIGTEN Header (nicht an 'changed').
// Reine Funktion, EIN Aufrufort (runWeekButtonClick).
export function classifyClickEffect({ prevStart = null, newWeek = null, targetStart = null, direction = 0 } = {}) {
  const newStart = newWeek && newWeek.startDate ? newWeek.startDate : null;
  if (!newStart || !prevStart) return CLICK_EFFECTS.unverified;
  if (newStart === prevStart) return CLICK_EFFECTS.unchanged;
  // Die Zielwoche ist das staerkste Signal und wird VOR der Richtung geprueft.
  if (renderedWeekContains(newWeek, targetStart)) return CLICK_EFFECTS.reached;
  const stepWeeks = Math.round(
    (new Date(newStart + 'T00:00:00') - new Date(prevStart + 'T00:00:00')) / (7 * 86400000)
  );
  const dir = Math.sign(direction);
  if (dir !== 0 && Math.sign(stepWeeks) === dir) return CLICK_EFFECTS.advanced;
  return CLICK_EFFECTS.overshoot;
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
export function assembleFreeBusyDiagnostics({ fresh = null, early = null, weekButtonProbe = [], abortReason = 'none', budget = null, eventElementsFound = null, selectorUsed = null, rawEventLabels = null, rejectionReasons = null, waitedForEventsMs = null, eventWaitOutcome = null, eventWaitAbortReason = null, eventWaitLoadingPolls = null, loadingIndicators = null, headerReadError = null, navAttempts = null } = {}) {
  const base = fresh || early || { url: '', headerCandidates: [], dayHeaderCandidates: [], weekNavCandidates: [], ariaButtonCount: 0, ariaLabelSample: [] };
  const out = {
    ...base,
    weekButtonProbe,
    abortReason
  };
  if (early && early.url) out.urlAfterNavigation = early.url;
  if (budget) out.budget = budget;
  // FIX 1: "kein angezeigter Header lesbar" ist ein eigener, benannter Grund
  // und wird auch in den FRUEHEN Abbruchpfaden sichtbar — sonst waere dort die
  // Diagnose "nichts gefunden" (und die Versuchung, den Button-Text zu nehmen).
  if (headerReadError) out.headerReadError = headerReadError;
  // FIX 3: die Navigations-Versuche mit ihrem Beleg (weekBefore/weekAfter/
  // movedView). Die URL-Korrektheit wird nirgends als Beleg der View gewertet.
  if (Array.isArray(navAttempts) && navAttempts.length) out.navAttempts = navAttempts;
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

// ---------------------------------------------------------------------------
// Teilnehmer-Verfügbarkeit (GetSchedule via OWA-GraphQL-Gateway)
//
// Live-Befund 27.10.2026 (Probe 1–6): Die OWA-App lädt Teilnehmer-Verfügbarkeit
// NICHT über service.svc/GetUserAvailabilityInternal (401 ohne Canary), sondern
// über das GraphQL-Gateway:
//   POST {origin}/outlookgatewayb2/graphql   (origin: outlook.cloud.microsoft)
//   Authorization: Bearer <JWT>  (aud outlook.office.com, ~24h gültig, scopes
//   Calendars.ReadWrite*) — OWA hängt den Token an JEDEN dieser Calls.
//   Body: [{ operationName:'GetSchedule', variables:{ input:{ userId,
//   availabilityViewInterval, startTime/endTime {dateTime,timeZone}, schedules[]
//   }}, query }]
// Der Token (und userId + Gateway-URL) werden aus dem Eigen-Traffic der
// Kalender-App beim Boot abgegriffen (page.on('request')) und gecacht. Der
// POST läuft über den Playwright-APIRequestContext (context.request) — kein
// CORS, Cookies automatisch, Timeouts hart deckelbar.
//
// availabilityView-Format: EIN Zeichen je Slot (intervalMinutes, Default 15):
//   0=free 1=tentative 2=busy 3=oof 4=workingElsewhere (sonst unknown).
// ---------------------------------------------------------------------------

// Status-Map je availabilityView-Zeichen.
export const AVAILABILITY_STATUS_MAP = {
  0: 'free', 1: 'tentative', 2: 'busy', 3: 'oof', 4: 'workingElsewhere'
};

// Hartes Gesamtzeitbudget für den KOMPLETTEN getAttendeeAvailability-Call.
// Muss sicher UNTER dem MCP-Client-Timeout liegen (vgl. FREEBUSY_TOTAL_BUDGET_MS).
export const ATTENDEE_TOTAL_BUDGET_MS = 40000;
// Max. Wartezeit auf den Boot-GraphQL-Call (Token-Abgriff) je Versuch.
export const ATTENDEE_TOKEN_WAIT_MAX_MS = 25000;
// HTTP-Timeout des Gateway-POSTs (einzelner Request).
export const ATTENDEE_HTTP_TIMEOUT_MS = 20000;
// Sicherheits-Marge, ab der ein gecachter Token als ablaufend gilt (2 min).
export const ATTENDEE_TOKEN_MARGIN_MS = 120000;
// Obergrenze der Teilnehmer je Anfrage (Gateway-Kosten + Antwortgröße).
export const ATTENDEE_MAX_ATTENDEES = 20;

// GraphQL-Query (kompakt, verifiziert gegen Live-Antwort 27.10.2026 — Probe 6:
// 200 OK mit schedules/availabilityView/scheduleItems/workingHours).
export const GET_SCHEDULE_QUERY = `query GetSchedule($input: GetScheduleInput) {
  getSchedule(request: $input) {
    schedules {
      availabilityView
      error { message responseCode diagnosticData }
      scheduleId
      scheduleItems {
        isPrivate location status subject isException isReminderSet isMeeting isRecurring id
        startTime { dateTime timeZone { name } }
        endTime { dateTime timeZone { name } }
      }
      workingHours { daysOfWeek startTime endTime timeZone { name bias } }
    }
  }
}`;

// Normalisiert die attendees-Eingabe: Strings -> E-Mail-Kandidaten (der erste
// E-Mail-Like-Teil gewinnt, extractEmailLike), lowercase, Dedupe, Cap. Reine
// Funktion: wirft NICHT, liefert { valid, invalid }.
export function normalizeAttendeeList(attendees, max = ATTENDEE_MAX_ATTENDEES) {
  const valid = [];
  const invalid = [];
  if (!Array.isArray(attendees)) return { valid, invalid };
  const seen = new Set();
  for (const raw of attendees) {
    if (valid.length >= max) break;
    if (typeof raw !== 'string') { if (raw != null) invalid.push(String(raw)); continue; }
    const email = extractEmailLike(raw);
    if (!email) { invalid.push(raw); continue; }
    const key = email.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    valid.push(key);
  }
  return { valid, invalid };
}

// Dekodiert die exp-Claim eines JWT (Payload, base64url) in ms. Wirft NICHT.
export function parseJwtExp(token) {
  try {
    const parts = String(token || '').split('.');
    if (parts.length < 2) return null;
    const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    const payload = JSON.parse(Buffer.from(b64, 'base64').toString('utf8'));
    return Number.isFinite(payload.exp) ? payload.exp * 1000 : null;
  } catch (e) {
    return null;
  }
}

// Dekodiert ausgewählte Identitäts-Claims eines JWT. Quelle für die
// GetSchedule-userId, wenn der OWA-Boot-Call sie nicht hergibt (Live-Befund
// 27.10.2026: upn/unique_name = Login-Alias, z.B. ext-dhal@d-velop.de).
export function parseJwtIdentity(token) {
  try {
    const parts = String(token || '').split('.');
    if (parts.length < 2) return {};
    const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    const payload = JSON.parse(Buffer.from(b64, 'base64').toString('utf8'));
    const pick = (v) => (typeof v === 'string' && v.includes('@')) ? v : null;
    return {
      upn: pick(payload.upn),
      uniqueName: pick(payload.unique_name),
      name: typeof payload.name === 'string' ? payload.name : null
    };
  } catch (e) {
    return {};
  }
}

// Ist der Token (noch) verwendbar? Marge: kurz vor Ablauf gilt als abgelaufen.
export function isTokenUsable(token, { marginMs = ATTENDEE_TOKEN_MARGIN_MS, nowMs = Date.now() } = {}) {
  const exp = parseJwtExp(token);
  return typeof token === 'string' && token.length > 0 && Number.isFinite(exp) && (nowMs + marginMs) < exp;
}

// Offset einer Zeitzone zu UTC für einen Zeitpunkt (ms; DST-sicher via Intl).
export function timeZoneOffsetMs(utcMs, timeZone) {
  const dtf = new Intl.DateTimeFormat('en-US', {
    timeZone, hour12: false,
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit'
  });
  const parts = {};
  for (const p of dtf.formatToParts(new Date(utcMs))) parts[p.type] = p.value;
  const asUtc = Date.UTC(
    Number(parts.year), Number(parts.month) - 1, Number(parts.day),
    Number(parts.hour) % 24, Number(parts.minute), Number(parts.second)
  );
  return asUtc - Math.floor(utcMs / 1000) * 1000;
}

// Naive Zeitangabe in einer Ziel-Zeitzone -> UTC-ISO. Werte MIT Offset/Z
// bleiben unverändert. Zwei Iterationen fangen DST-Grenzfälle.
export function zonedTimeToUtcIso(naiveIso, timeZone = 'Europe/Berlin') {
  const s = String(naiveIso || '').trim();
  if (!s) return null;
  if (/[zZ]$|[+-]\d{2}:?\d{2}$/.test(s)) {
    const direct = new Date(s).getTime();
    return Number.isNaN(direct) ? null : new Date(direct).toISOString();
  }
  const guessMs = Date.parse(s + 'Z');
  if (Number.isNaN(guessMs)) return null;
  let utcMs = guessMs;
  for (let i = 0; i < 2; i++) {
    const corrected = guessMs - timeZoneOffsetMs(utcMs, timeZone);
    if (corrected === utcMs) break;
    utcMs = corrected;
  }
  return new Date(utcMs).toISOString();
}

// Baut den GraphQL-Request-Body (Array-Wrapper, wie OWA selbst sendet).
// dateTime-Strings werden als UTC übergeben (ohne 'Z', wie der OWA-Eigencall).
export function buildGetScheduleRequest({ userId, attendeeEmails, startUtcIso, endUtcIso, intervalMinutes = 15 }) {
  const stripZ = (iso) => String(iso || '').replace(/\.(\d{3})Z$/, '.$1').replace(/Z$/, '');
  return [{
    operationName: 'GetSchedule',
    variables: {
      input: {
        userId: userId || attendeeEmails[0],
        availabilityViewInterval: intervalMinutes,
        startTime: { dateTime: stripZ(startUtcIso), timeZone: { name: 'UTC' } },
        endTime: { dateTime: stripZ(endUtcIso), timeZone: { name: 'UTC' } },
        schedules: [...attendeeEmails]
      }
    },
    query: GET_SCHEDULE_QUERY
  }];
}

// Dekodiert availabilityView in zusammengeführte Segmente, auf das Fenster
// geclippt. Reine Funktion, wirft NICHT (ungültig -> leere Segmente).
export function decodeAvailabilityView(view, { startUtcIso, endUtcIso, intervalMinutes = 15 } = {}) {
  const v = String(view || '');
  const startMs = Date.parse(startUtcIso || '');
  const endMs = Date.parse(endUtcIso || '');
  const raw = { segments: [], busySegments: [], slotCount: 0 };
  if (!v || !Number.isFinite(startMs) || !Number.isFinite(endMs) || !(intervalMinutes > 0)) return raw;
  const out = [];
  for (let i = 0; i < v.length; i++) {
    const slotStart = startMs + i * intervalMinutes * 60000;
    const slotEnd = slotStart + intervalMinutes * 60000;
    if (slotStart >= endMs) break;
    if (slotEnd <= startMs) continue;
    const code = v[i];
    const from = Math.max(slotStart, startMs);
    const to = Math.min(slotEnd, endMs);
    out.push({
      status: AVAILABILITY_STATUS_MAP[code] || 'unknown',
      startUtcIso: new Date(from).toISOString(),
      endUtcIso: new Date(to).toISOString(),
      minutes: Math.round((to - from) / 60000)
    });
  }
  // Zusammenführen gleicher Nachbar-Status (Slots desselben Status).
  const segments = [];
  for (const s of out) {
    const last = segments[segments.length - 1];
    if (last && last.status === s.status) {
      last.endUtcIso = s.endUtcIso;
      last.minutes += s.minutes;
    } else {
      segments.push({ ...s });
    }
  }
  return {
    segments,
    busySegments: segments.filter(s => s.status !== 'free'),
    slotCount: v.length
  };
}

// Parst die Gateway-Antwort: Array-Wrapper ODER nacktes Objekt, GraphQL-
// errors-Array, Nicht-JSON. Liefert { ok, schedules, errorInfo }.
export function parseGetScheduleResponse(bodyText) {
  let json = null;
  try {
    json = JSON.parse(String(bodyText || ''));
  } catch (e) {
    return { ok: false, schedules: [], errorInfo: { kind: 'non-json', snippet: String(bodyText || '').slice(0, 400) } };
  }
  const envelope = Array.isArray(json) ? json[0] : json;
  if (!envelope || typeof envelope !== 'object') {
    return { ok: false, schedules: [], errorInfo: { kind: 'empty-envelope', snippet: String(bodyText || '').slice(0, 200) } };
  }
  if (Array.isArray(envelope.errors) && envelope.errors.length) {
    return { ok: false, schedules: [], errorInfo: { kind: 'graphql-errors', errors: envelope.errors.slice(0, 5) } };
  }
  const schedules = envelope?.data?.getSchedule?.schedules;
  if (!Array.isArray(schedules)) {
    return { ok: false, schedules: [], errorInfo: { kind: 'missing-schedules', snippet: String(bodyText || '').slice(0, 300) } };
  }
  return { ok: true, schedules, errorInfo: null };
}

// Mergt überlappende/sortierte Intervall-Listen (ms) zu einer disjunkten Liste.
function mergeBusyIntervals(intervals) {
  const sorted = [...intervals].filter(i => i.endMs > i.startMs).sort((a, b) => a.startMs - b.startMs);
  const out = [];
  for (const iv of sorted) {
    const last = out[out.length - 1];
    if (last && iv.startMs <= last.endMs) {
      last.endMs = Math.max(last.endMs, iv.endMs);
    } else {
      out.push({ startMs: iv.startMs, endMs: iv.endMs });
    }
  }
  return out;
}

// Invertiert Busy-Intervalle in Freie innerhalb des Fensters.
function freeIntervalsFor(busyMerged, windowStartMs, windowEndMs) {
  const free = [];
  let cursor = windowStartMs;
  for (const iv of busyMerged) {
    if (iv.startMs > cursor) free.push({ startMs: cursor, endMs: Math.min(iv.startMs, windowEndMs) });
    cursor = Math.max(cursor, iv.endMs);
    if (cursor >= windowEndMs) break;
  }
  if (cursor < windowEndMs) free.push({ startMs: cursor, endMs: windowEndMs });
  return free.filter(f => f.endMs > f.startMs);
}

// Schnittmenge zweier sortierter, disjunkter Freier-Listen.
function intersectTwoFreeLists(a, b) {
  const out = [];
  let i = 0, j = 0;
  while (i < a.length && j < b.length) {
    const startMs = Math.max(a[i].startMs, b[j].startMs);
    const endMs = Math.min(a[i].endMs, b[j].endMs);
    if (endMs > startMs) out.push({ startMs, endMs });
    if (a[i].endMs < b[j].endMs) i++; else j++;
  }
  return out;
}

// Gemeinsame freie Slots aller Teilnehmer (min. Dauer gefiltert).
export function intersectFreeSegments(busySegmentsPerAttendee, { startUtcIso, endUtcIso, minDurationMinutes = 15 } = {}) {
  const startMs = Date.parse(startUtcIso || '');
  const endMs = Date.parse(endUtcIso || '');
  if (!Number.isFinite(startMs) || !Number.isFinite(endMs) || endMs <= startMs) return [];
  let common = freeIntervalsFor([], startMs, endMs); // komplettes Fenster
  for (const busyList of busySegmentsPerAttendee) {
    const merged = mergeBusyIntervals(busyList.map(s => ({
      startMs: Date.parse(s.startUtcIso),
      endMs: Date.parse(s.endUtcIso)
    })).filter(i => Number.isFinite(i.startMs) && Number.isFinite(i.endMs)));
    const free = freeIntervalsFor(merged, startMs, endMs);
    common = intersectTwoFreeLists(common, free);
    if (!common.length) return [];
  }
  const minMs = minDurationMinutes * 60000;
  return common
    .filter(f => (f.endMs - f.startMs) >= minMs)
    .map(f => ({
      start: new Date(f.startMs).toISOString(),
      end: new Date(f.endMs).toISOString(),
      durationMinutes: Math.round((f.endMs - f.startMs) / 60000)
    }));
}

// UTC-ISO -> lesbare Lokalzeit (de-DE) für die Agent-Ausgabe.
export function formatUtcIsoInZone(iso, timeZone = 'Europe/Berlin') {
  const ms = Date.parse(iso || '');
  if (!Number.isFinite(ms)) return null;
  return new Intl.DateTimeFormat('de-DE', {
    timeZone, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', hour12: false
  }).format(new Date(ms));
}

export class OwaClient {
  constructor(realm = config.defaultRealm) {
    this.realm = realm;
    // Token-Cache für das GraphQL-Gateway (Teilnehmer-Verfügbarkeit):
    // Token + Ablauf, Gateway-URL und signed-in userId werden aus dem
    // Eigen-Traffic der Kalender-App abgegriffen (captureScheduleToken).
    this._scheduleToken = null;
    this._scheduleTokenExpMs = 0;
    this._scheduleGraphqlUrl = null;
    this._scheduleUserId = null;
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

  // TEXT-INJECTION OHNE AUTOKORREKTUR (Befund 15.10.2026): keyboard.type() tippt
  // zeichenweise — die OWA-Autokorrektur kapitalisiert wortunabhängig Wörter
  // nach "bzw. " ("dem"->"Dem", "zum"->"Zum") und Absatzanfänge
  // ("kurzes"->"Kurzes"). Zwei Wege umgehen die Autokorrektur vollständig:
  //   1) keyboard.insertText() (CDP Input.insertText) — erzeugt ein echtes
  //      Input-Event OHNE Key-Events, also keinen Autokorrektur-Pfad, und ist
  //      serialisierbar (das Editor-Modell verarbeitet den Input-Event)
  //   2) Clipboard-Paste (navigator.clipboard.writeText + Ctrl+V) — im
  //      Live-Betrieb an Teams erprobt; kein Autokorrektur-Pfad
  // Nach jedem Weg wird der Editor-Text gelesen und gegen den Solltext geprüft.
  // MISSERFOLG wird NIEMALS still gemeldet: Rueckgabe enthaelt immer Methode,
  // Trefferklasse und den Editor-Stand (fuer Soll/Ist-Diff durch den Aufrufer).
  async _injectBodyText(activePage, editor, body) {
    const readEditor = async () => (await editor.evaluate((el) => (el.innerText || el.textContent || '').trim()));

    // Weg 1: CDP Input.insertText (kein Autokorrektur-Pfad)
    try {
      await activePage.keyboard.insertText(body);
      await activePage.waitForTimeout(600);
      const editorText = await readEditor();
      const match = classifyBodyMatch(editorText, body);
      if (match !== 'none') {
        return { method: 'cdp-insertText', editorText, match };
      }
    } catch (e) {
      //insertText nicht verfuegbar -> naechster Weg
    }

    // Teil-Insertion rueckgaengig machen, bevor der Clipboard-Weg doppelt einfuegen wuerde
    await activePage.keyboard.press('Control+z').catch(() => null);
    await activePage.waitForTimeout(400);

    // Weg 2: Clipboard-Paste (Caret bleibt an seiner Position — additiv, kein Ctrl+A)
    try {
      await activePage.evaluate((t) => navigator.clipboard.writeText(t), body);
      await activePage.waitForTimeout(200);
      await activePage.keyboard.press('Control+v');
      await activePage.waitForTimeout(800);
      const editorText = await readEditor();
      const match = classifyBodyMatch(editorText, body);
      if (match !== 'none') {
        return { method: 'clipboard-paste', editorText, match };
      }
    } catch (e) {
      //Clipboard-Pfad nicht verfuegbar -> Misserfolg melden
    }

    // Beide Wege gescheitert: Insertion rueckgaengig machen, damit kein
    // verfaelschter Entwurf zurueckbleibt, und offen melden.
    await activePage.keyboard.press('Control+z').catch(() => null);
    await activePage.waitForTimeout(400);
    return { method: null, editorText: await readEditor(), match: 'none' };
  }

  async createDraft({ to, subject, body, cc }) {
    if (!body || typeof body !== 'string' || body.trim().length === 0) {
      throw new Error('Parameter "body" ist erforderlich und darf nicht leer sein.');
    }

    const page = await this.getPage(true);

    const toStr = Array.isArray(to) ? to.join('; ') : (to || '');
    const ccStr = Array.isArray(cc) ? cc.join('; ') : (cc || '');
    const subjectStr = subject || '';
    const bodyStr = body;

    // Important: Do not pass body in URL params, as that forces OWA into plain-text and strips the default signature!
    const composeUrl = buildComposeUrl(this.realm, { to, subject, cc });

    // Der Compose-DeepLink kann einen NEUEN Tab oeffnen (gleiche Faehigkeit wie
    // beim Kalender-DeepLink) — Tabliste vorher merken. (Realm-Multiplex: Context dieses Realms)
    const context = browserManager.getContext(this.realm);
    const pagesBefore = context ? context.pages().length : 0;

    await page.goto(composeUrl, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(5000);

    // Falls ein neuer Tab geoeffnet wurde: auf diesen wechseln
    let activePage = page;
    if (context) {
      const pagesAfter = context.pages();
      if (pagesAfter.length > pagesBefore) {
        activePage = pagesAfter[pagesAfter.length - 1];
        this.lastPage = activePage;
      }
    }

    // ECHTER EDITOR (gleicher Standard wie replyDraft, Live-Befund 09.10.2026):
    // Es gibt MEHRERE Knoten mit aria-label "Message body" — den echten Editor
    // (role=textbox / contenteditable=true) und einen Wrapper ohne
    // contenteditable. Der Wrapper liefert CSS-Quelltext statt Nachrichtentext,
    // deshalb wird er abgewiesen statt stillschweigend verwendet.
    //
    // ROBUSTHEIT (Live-Befund 09.10.2026, adesso-Realm): Direkt nach einem
    // Realm-Wechsel kann OWA den Compose-DeepLink verschlucken (Redirect auf
    // /mail/, App noch im Ladeuebergang). Der Editor ist dann nicht da, obwohl
    // Selektor und Profil korrekt sind. Deshalb: URL pruefen und GENAU EINMAL
    // erneut navigieren, bevor abgebrochen wird.
    const EDITOR_SEL =
      'div[role="textbox"][contenteditable="true"][aria-label*="Message body" i], ' +
      'div[role="textbox"][contenteditable="true"][aria-label*="Nachrichtentext" i], ' +
      '[contenteditable="true"][aria-label*="Message body" i], ' +
      '[contenteditable="true"][aria-label*="Nachrichtentext" i], ' +
      'div[role="textbox"][aria-label*="Message body" i], ' +
      'div[role="textbox"][aria-label*="Nachrichtentext" i]';

    let editor = await activePage.waitForSelector(EDITOR_SEL, { timeout: 15000 }).catch(() => null);

    if (!editor) {
      const urlAfterFirstTry = activePage.url();
      // GENAU EIN Recovery-Versuch: frisch navigieren (kein Loop).
      await activePage.goto(composeUrl, { waitUntil: 'domcontentloaded' });
      await page.waitForTimeout(6000);
      editor = await activePage.waitForSelector(EDITOR_SEL, { timeout: 15000 }).catch(() => null);
      if (!editor) {
        throw new Error(
          `Compose-Editor für den Entwurf konnte nicht gefunden werden. Es wurde nichts gespeichert. ` +
          `Diagnostik: URL nach dem 1. Versuch: ${urlAfterFirstTry} | URL jetzt: ${activePage.url()} | ` +
          `Tabs: ${(context ? context.pages().length : 1)} (vorher ${pagesBefore}).`
        );
      }
    }

    const editorKind = await editor.evaluate((el) => ({
      contentEditable: el.getAttribute && el.getAttribute('contenteditable'),
      role: (el.getAttribute && el.getAttribute('role')) || '',
      tag: el.tagName
    }));
    if (!isRealComposeEditor(editorKind)) {
      throw new Error(
        `Gefundener Compose-Knoten ist kein echter Editor ` +
        `(contenteditable="${editorKind.contentEditable}", role="${editorKind.role}"). ` +
        `Eine Textpruefung waere wertlos - es wird nichts getippt und nichts gespeichert.`
      );
    }

    // Bestand vor dem Tippen merken (Unversehrtheits-Beweis fuer die
    // OWA-Account-Signatur, die OWA beim Oeffnen automatisch einsetzt).
    const preInsertText = await editor.evaluate((el) => (el.innerText || el.textContent || '').trim());

    // TEXTUEBERNAHME PER TASTATUR — derselbe Serialisierungs-Weg wie bei
    // replyDraft. Die fruehere reine DOM-Insertion (insertBefore + dispatchEvent)
    // wurde von OWA NICHT serialisiert (Live-Befund 09.10.2026: Entwurf
    // gespeichert, Koerper leer). Ein DOM-Eingriff ist fuer das interne
    // Editor-Modell unsichtbar.
    //
    // SICHERHEIT: Caret per DOM-Range an den Inhaltsanfang, KEIN Control+A —
    // das wuerde eine vorhandene Signatur markieren und beim Tippen
    // ueberschreiben.
    await editor.click();
    await activePage.waitForTimeout(300);

    const caretPlaced = await editor.evaluate((el) => {
      try {
        const target = el.firstChild || el;
        const range = document.createRange();
        range.setStart(target, 0);
        range.collapse(true);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        el.focus();
        return true;
      } catch (e) {
        return false;
      }
    });
    if (!caretPlaced) {
      throw new Error('Cursor konnte im Compose-Editor nicht gesetzt werden. Es wurde nichts gespeichert und nichts überschrieben.');
    }

    // TEXTUEBERNAHME autokorrektur-frei (Befund 15.10.2026): insertText/Clipboard
    // statt zeichenweisem keyboard.type — OWA kapitalisierte sonst Wörter nach
    // "bzw. " und Absatzanfänge. Bei Totalversagen wird NICHT gespeichert und
    // die Fehlermeldung enthaelt Soll/Ist (kein stiller Misserfolg).
    const injection = await this._injectBodyText(activePage, editor, bodyStr);
    const editorText = injection.editorText;
    if (!injection.method) {
      throw new Error(
        `Entwurfstext konnte NICHT in den Editor übernommen werden (insertText- und ` +
        `Clipboard-Weg fehlgeschlagen). Es wurde nichts gespeichert. Soll/Ist: ` +
        `${JSON.stringify(buildBodySollIstDiff(bodyStr, editorText))}.`
      );
    }
    if (!isHistoryPreserved(preInsertText, editorText)) {
      throw new Error(
        'ABBRUCH: Beim Einfügen des Texts ging die vorhandene Signatur verloren. ' +
        'Der Entwurf wurde NICHT gespeichert, damit kein Inhaltsverlust entsteht.'
      );
    }

    // Save draft
    await activePage.keyboard.press('Control+s');
    await activePage.waitForTimeout(3000);

    // PERSISTIERTER NACHWEIS — derselbe Beweisstandard wie replyDraft.
    // Bei CC-Auftrag zusätzlich die persistierten Empfänger mitlesen.
    const persisted = await this.verifyPersistedDraftBody(bodyStr, { withRecipients: Boolean(ccStr) });

    // P4 (Befund 15.10.2026): CC via &cc=DeepLink wird von OWA teils NICHT
    // persistiert — der frühere Tool-Return sagte cc ok, der gespeicherte Draft
    // hatte keine CC-Chips. Deshalb: Empfänger des PERSISTIERTEN Drafts gegen
    // den Auftrag prüfen, fehlende CC nachziehen (UI-Chips) und erneut
    // verifizieren; sonst klarer Warn-Return.
    let ccState = null;
    if (ccStr && persisted.verified === true) {
      const recips = persisted.recipients || { to: [], cc: [] };
      const inCc = recipientListContains(recips.cc, ccStr);
      const inTo = recipientListContains(recips.to, ccStr);
      if (inCc) {
        ccState = { ccPersisted: true, ccLocation: 'cc', recipients: recips };
      } else if (inTo) {
        ccState = {
          ccPersisted: false, ccLocation: 'to', recipients: recips,
          ccNote: 'CC-Empfänger ist im Draft unter "An" statt "Cc" — Zustellung gesichert, Verteilung weicht vom Auftrag ab.'
        };
      } else {
        const repair = await this._repairDraftCc(ccStr);
        let ok = false;
        let recips2 = recips;
        let where = null;
        if (repair.repaired) {
          const recheck = await this._readPersistedDraftRecipients();
          if (recheck.found) {
            recips2 = recheck.recipients;
            if (recipientListContains(recips2.cc, ccStr)) { ok = true; where = 'cc'; }
            else if (recipientListContains(recips2.to, ccStr)) { ok = true; where = 'to'; }
          }
        }
        ccState = {
          ccPersisted: ok,
          ...(where ? { ccLocation: where } : {}),
          ccRepaired: Boolean(repair.repaired),
          ...(repair.reason ? { ccRepairReason: repair.reason } : {}),
          recipients: recips2
        };
      }
    }

    const ccMeta = ccState ? {
      ccRequested: ccStr,
      ccPersisted: ccState.ccPersisted,
      ...(ccState.ccLocation ? { ccLocation: ccState.ccLocation } : {}),
      ...(ccState.ccNote ? { ccNote: ccState.ccNote } : {}),
      ...(ccState.ccRepaired !== undefined ? { ccRepaired: ccState.ccRepaired } : {}),
      ...(ccState.ccRepairReason ? { ccRepairReason: ccState.ccRepairReason } : {}),
      recipients: ccState.recipients
    } : {};

    const bodyMeta = {
      bodyInsertSource: injection.method,
      bodyMatch: persisted.match || 'unverified',
      ...(persisted.match === 'fuzzy' ? { bodySollIst: persisted.diff } : {})
    };

    if (persisted.verified !== true) {
      return {
        success: false,
        to: toStr,
        cc: ccStr || undefined,
        subject: subjectStr,
        preview: bodyStr.length > 200 ? bodyStr.substring(0, 200) + '...' : bodyStr,
        ...bodyMeta,
        persistedBodyVerified: persisted.verified === null ? 'unverified' : false,
        folder: 'Entwürfe (Drafts)',
        message:
          `WARNUNG: Der Entwurf wurde gespeichert, sein Text ist im gespeicherten Entwurf ` +
          `aber NICHT nachweisbar. Der Entwurf bleibt zur Sichtung in "Entwürfe" liegen und ` +
          `wurde NICHT automatisch gelöscht. Bitte manuell prüfen und ggf. neu anlegen. ` +
          `Grund: ${persisted.detail} | Soll/Ist: ${JSON.stringify(persisted.diff || buildBodySollIstDiff(bodyStr, ''))}`
      };
    }

    return {
      success: true,
      to: toStr,
      cc: ccStr || undefined,
      subject: subjectStr,
      preview: bodyStr.length > 200 ? bodyStr.substring(0, 200) + '...' : bodyStr,
      ...bodyMeta,
      ...ccMeta,
      persistedBodyVerified: true,
      folder: 'Entwürfe (Drafts)',
      message:
        `Entwurf wurde in Outlook unter "Entwürfe" gespeichert und im gespeicherten Zustand verifiziert ` +
        `(Textübernahme: ${injection.method}${persisted.match === 'fuzzy' ? ', mit Abweichungen — siehe bodySollIst' : ''}).` +
        (ccState && ccState.ccPersisted === false
          ? ` WARNUNG: CC wurde NICHT (bzw. nicht als "Cc") persistiert${ccState.ccNote ? ' — ' + ccState.ccNote : ' — bitte im Draft manuell prüfen, bevor gesendet wird.'}`
          : '')
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

    // Der Kalender-Deep-Link kann einen NEUEN Tab öffnen. (Realm-Multiplex: Context dieses Realms)
    const context = browserManager.getContext(this.realm);
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
    // FIX 1/FIX 3: Beide Felder werden hier deklariert und NICHT weiter unten —
    // die fruehen Abbruchpfade (no-calendar-tab, early-no-dated-header) liegen
    // VOR der Klick-Schleife und brauchen sie schon. Das ist exakt die TDZ-Falle
    // aus 0d47105 und 8eb70cc: ein `let` nach der schliessenden Klammer einer
    // fruehen Rueckgabe.
    let headerReadError = null;
    const navAttempts = [];

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
    //
    // BUG 5: Die Probe liefert jetzt auch den Lade-Indikator. Grund: der
    // Live-Zustand war "Body hat Events, Header fehlt" — und die
    // Gnadenfrist fuer den Header wird waehrend eines Spinners nicht
    // abgearbeitet, weil ein Spinner beweist, dass die Ansicht noch laedt.
    let lastEventSelector = null;
    let lastEventProbe = { events: false, loading: false };
    const probeEvents = async () => {
      const r = await collectCalendarEventLabels(page, { preferSelector: lastEventSelector, weekStartIso: mondayIso, fast: true });
      if (r.eventElementsFound > 0 && r.selectorUsed) lastEventSelector = r.selectorUsed;
      lastEventProbe = { events: r.eventElementsFound > 0, loading: r.loading === true };
      return lastEventProbe;
    };
    const hasEventElements = async () => (await probeEvents()).events;

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
    //
    // BUG 5 (Live-Befund 26.09., Lauf 4): der Zweig `else if (events) break;`
    // behandelte "kein Header, aber Body-Events" als FERTIG. Genau das war der
    // Live-Zustand (18 Event-Knoten, 0 Header) — die Schleife verliess die
    // Ansicht nach dem ersten Poll, lange bevor der Header gerendert sein konnte.
    // Beleg: waitedForEventsMs 13 / polls 1 bei Budget 20000, headerCandidates
    // leer, WOCHENDANN aber 18 Event-Knoten im DOM. Der Header fehlt nicht, weil
    // die Ansicht fertig ist, sondern weil sie noch nicht fertig ist: er ist die
    // einzige belastbare Wochenquelle UND der Bereich, in dem die Wochen-Buttons
    // liegen. Ohne ihn ist die Wochenlage unbekannt — das ist Zustand 1
    // ("warten"), nicht "fertig".
    // Der Zweig gibt der Ansicht darum dieselbe Gnadenfrist wie einem Header und
    // verlaengert sie, solange ein Spinner steht.
    const EVENT_GRACE_MS = 2500;
    const waitForCalendarRendered = async (budgetMs) => {
      const waitDeadline = Date.now() + budgetMs;
      let rendered = { text: null, startDate: null, endDate: null };
      let events = false;
      let headerSeenAt = null;
      let eventsSeenAt = null;
      while (Date.now() < waitDeadline && remaining() > 0) {
        rendered = await readRenderedCalendarWeek(page);
        events = await hasEventElements();
        const loading = lastEventProbe.loading === true;
        if (rendered.startDate) {
          if (!headerSeenAt) headerSeenAt = Date.now();
          if (events) break;
          if (Date.now() - headerSeenAt >= EVENT_GRACE_MS) break; // Gnadenfrist abgelaufen
        } else if (events) {
          // Header fehlt, Body rendert bereits Events -> Zustand 1, nicht "fertig".
          if (!eventsSeenAt || loading) eventsSeenAt = Date.now();
          if (Date.now() - eventsSeenAt >= EVENT_GRACE_MS) break; // Gnadenfrist abgelaufen
        }
        if (Date.now() + 500 > waitDeadline || remaining() <= 0) break;
        await page.waitForTimeout(500);
      }
      return { rendered, hasEvents: events };
    };

    // Zustand 1: WARTEN statt klicken (BUG 5). Der Wochenzustand wird bei jedem
    // Poll neu gelesen — "warte bis der Header da ist" heisst "pruefe waehrend
    // des Wartens", nicht "einmal am Anfang schauen". Der Spinner haelt den Poll
    // am Leben und macht das Ruhe-Fenster ungueltig.
    //
    // Dies ist die EINZIGE Wartefunktion fuer den fehlenden Wochen-Header: die
    // aeussere Schritt-Schleife und der Klickpfad benutzen beide dieselbe
    // (ueber Dependency-Injection), damit die beiden nicht auseinanderlaufen
    // koennen — genau so ist der Live-Befund entstanden.
    const waitForViewReady = async () => {
      const waited = await waitForWeekViewReady({
        readWeek: () => readRenderedCalendarWeek(page),
        readEvents: () => probeEvents(),
        budgetMs: phase(FREEBUSY_VIEW_WAIT_BUDGET_MS),
        pollMs: FREEBUSY_EVENT_WAIT_POLL_MS,
        settleMs: FREEBUSY_EVENT_WAIT_SETTLE_MS,
        targetStart: rangeStartDate,
        sleep: (ms) => page.waitForTimeout(ms)
      });
      return { ...waited, events: hasEvents || (await hasEventElements()) };
    };

    // BUG 5 — Navigation in den drei belegten Wegen, in dieser Reihenfolge
    // bewertet (siehe Bericht):
    //  (a) Wochen-Buttons ("Go to next week …") — braucht einen GERENDERTEN
    //      Header-Bereich. In Runde 4 existierten die Button-Indikatoren nicht
    //      (weekNavCandidates: []), waehrend sie in Runde 3 da waren UND der
    //      Klick dort eine datierbare Zielwoche erreicht hat. Die Navigation ist
    //      also nicht als Defekt belegt — sie war schlicht noch nicht da. In
    //      Zustand 1 ist sie deshalb gar nicht erst erlaubt (kein Header ->
    //      keine Buttons -> keine Wochenrichtung bekannt).
    //  (b) "Jump to a specific date or date range" — ein moeglicher dritter Weg,
    //      aber ohne live verifizierten Selektor. Blindes DOM-Raten waere hier
    //      ein Reparaturversuch ohne Beleg; deshalb NICHT implementiert.
    //  (c) harter Reload der startdate-URL. Das ist der gewaehlte Weg fuer
    //      Zustand 1: er ist der EINZIGE der drei, der weder einen gerenderten
    //      Header noch eine geratene Wochenrichtung braucht, und er zielt genau
    //      auf das belegte Symptom: die URL enthielt startdate=2026-09-28, der
    //      Inhalt war KW 39 — der Parameter wurde also vom Client-Store
    //      ueberstimmt. Ein harter Reload setzt ihn durch.
    const hardReloadOptimized = async () => {
      const url = page.url() || '';
      if (url.includes('startdate=')) {
        await page.reload({ waitUntil: 'domcontentloaded', timeout: phase(10000) }).catch(() => null);
        return true;
      }
      return gotoCalendar(optimizedUrl);
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
    //
    // BUG 5: Der Klickpfad darf den Zustand 1 nicht mehr "durchreichen" und
    // seinen Aufrufer mit einem Abbruch belassen, ohne gewartet zu haben. Er
    // wartet selbst — ueber DIE SELBE Wartefunktion, die auch die aeussere
    // Schleife benutzt (waitForViewReady, Dependency-Injection). Damit kann der
    // Aufrufer nicht versehentlich einen anderen Warteweg unterstellen, und der
    // Test kann den Warte-Schritt am echten Pfad fahren.
    // Bleibt die Ansicht danach ungerendert, ist die Verweigerung ehrlich UND
    // das Wartebudget ist verbraucht: der Probe-Eintrag nennt waitedMs.
    //
    // viewWaitMsSpent wird VOR der clickWeekButton-Definition deklariert: die
    // beiden vorherigen Runden hatten je eine TDZ-Stelle in diesem Block
    // (0d47105, 8eb70cc) — ein `let` nach der schliessenden Klammer einer
    // asynchronen Closure ist genau die Klasse Fehler, die hier zweimal
    // passiert ist.
    let viewWaitMsSpent = 0;
    let lastViewWait = null;
    const clickWeekButton = async (direction, prevStart) => {
      const guard = await runWeekButtonClick({
        direction,
        prevStart,
        targetStart: rangeStartDate,
        page,
        probe: weekButtonProbe,
        readWeek: () => readRenderedCalendarWeek(page),
        afterClick: (prev, budgetMs) => waitRenderAfterClick(prev, phase(budgetMs)),
        clickWaitMs: FREEBUSY_CLICK_RENDER_WAIT_MS,
        clickCostMs: CLICK_COST_MS,
        remainingMs: remaining,
        waitForReady: async () => {
          if (viewWaitMsSpent >= FREEBUSY_VIEW_WAIT_BUDGET_MS) return null;
          const before = viewWaitMsSpent;
          const w = await waitForViewReady();
          viewWaitMsSpent = before + Math.round(w.waitedMs || 0);
          return w;
        },
        onGuardBlocked: (info) => { lastViewWait = info || null; }
      });
      return guard;
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
    // Woche der Zielwoche entspricht.
    //
    // FIX 3: die Klick-Rationierung ist auf FREEBUSY_MAX_BUTTON_CLICKS (= 2)
    // verschärft. Eine Kalenderwochen-Navigation braucht fuer den belegten Fall
    // EINE Woche vor/zurueck; in Lauf 5 wurden 6 Klicks in der Folge
    // prev/next/prev/next abgesetzt (8706ms Budget) und dann 'max-clicks'
    // gemeldet — bei einer View, die die ganze Zeit richtig war. Bei Mehrbedarf
    // gibt es jetzt einen EHRlichen Abbruchgrund ('too-far'), kein Blindklicken.
    const MAX_BUTTON_CLICKS = FREEBUSY_MAX_BUTTON_CLICKS;
    // FIX 3: die Richtung wird aus der DIFFERenz angefragte vs. angezeigte
    // Woche bestimmt und fuer den Aufruf VERRIEGELT. In Lauf 5 wechselte sie
    // pro Iteration (prev/next/prev/next) — das ist per Definition kein
    // gerichteter Weg zur Zielwoche. Ein Richtungswechsel wird als
    // 'week-overshoot' gemeldet, nicht durchgefuehrt.
    let lockedDirection = null;
    let gotoRetryDone = false;
    let farJumpDone = false;
    // BUG 5: 1x harter Reload der startdate-URL als Notfallweg fuer Zustand 1.
    let reloadDone = false;
    let buttonClicks = 0;
    let abortReason = 'none';
    // Klick-Sperre pro Richtung (Iteration 7, Defekt 2): verhindert, dass bei
    // unlesbarem Header ('unverified') mehrfach in dieselbe Richtung geklickt und
    // die Zielwoche ÜBERSPRUNGEN wird. Die Gegenrichtung (Rückweg) bleibt frei.
    let lastClickDirection = null;
    let weekConfirmedAfterClick = false;

    for (let step = 0; step < 30; step++) {
      // EIN Zustand, beide Aufrufer (Klickpfad und diese Schleife):
      // weekViewState ist die Klassifikation, weekClickDecision nur ihre
      // Abbildung auf skip/reason. Vorher hatte diese Schleife einen eigenen,
      // schwaecheren Test (isRequestedWeek) und der Klickpfad einen zweiten mit
      // hasEvents-Ausnahme — genau daraus entstanden die Live-Befunde.
      //
      // BUG 5, die drei Zustaende:
      //   1) view-not-ready  -> WARTEN (nicht klicken), Notfall-Reload, dann
      //                         ehrlicher Abbruch mit verbrauchtem Wartebudget
      //   2) target-week     -> fertig, KEIN Klick
      //   3) wrong-week      -> Wochen-Klick, danach wird der Header verifiziert
      const view = weekViewState({ rendered, targetStart: rangeStartDate });

      // --- Zustand 2: Zielwoche gerendert. Fertig, kein Klick. -------------
      if (view.state === WEEK_VIEW_STATES.targetWeek) {
        weekButtonProbe.push(weekProbeGuardEntry(weekClickDecision({ rendered, targetStart: rangeStartDate })));
        break;
      }
      if (remaining() <= 0) { abortReason = 'budget'; break; }

      // --- Zustand 1: Ansicht nicht bereit. WARTEN, NICHT KLICKEN. ----------
      if (view.state === WEEK_VIEW_STATES.notReady) {
        // BUG 5 (Live-Befund Lauf 4): hier wurde mit 'view-not-ready' abgebrochen
        // — nach 13ms, ohne den Header je abzuwarten. Damit verweigerte das
        // Werkzeug genau die Navigation, die zur Zielwoche gefuehrt haette.
        // Der Wochenzustand wird jetzt WAEHREND des Wartens neu gelesen
        // (waitForWeekViewReady, Spinner-respektierend).
        if (viewWaitMsSpent < FREEBUSY_VIEW_WAIT_BUDGET_MS) {
          const w = await waitForViewReady();
          rendered = w.rendered || rendered;
          hasEvents = w.events === true;
          if (!w.exhausted) continue;         // Zustand frisch bewerten
        }
        // Das Wartebudget ist verbraucht und der Header fehlt immer noch. Der
        // einzige verbleibende Weg, der weder Header noch Wochenrichtung
        // braucht: harter Reload der startdate-URL (1x, budgetiert).
        if (!reloadDone && remaining() >= FREEBUSY_MIN_REMAINING_FOR_ESCALATION_MS) {
          reloadDone = true;
          // FIX 3: der Reload bleibt EIN Versuch, und seine WIRKUNG wird
          // belegt statt behauptet. In Lauf 5 war die URL nach dem Reload
          // korrekt (startdate=2026-09-28) und die View trotzdem KW 39 — die
          // URL-Korrektheit darf nie als Beleg der View gelten.
          const before = await readRenderedCalendarWeek(page);
          await hardReloadOptimized();
          const w = await waitForViewReady();
          rendered = w.rendered || rendered;
          hasEvents = w.events === true;
          navAttempts.push({
            kind: 'startdate-reload',
            url: page.url() || null,
            weekBefore: before.startDate || null,
            weekAfter: rendered.startDate || null,
            movedView: Boolean(rendered.startDate && rendered.startDate !== before.startDate)
          });
          continue;
        }
        // Ehrlicher Abbruch: 'view-not-ready' NACH verbrauchtem Wartebudget. Der
        // Probe-Eintrag nennt die gewartete Zeit, damit ein sofortiges und ein
        // ausgereiftes Scheitern unterscheidbar bleiben.
        weekButtonProbe.push(weekProbeGuardEntry(
          weekClickDecision({ rendered, targetStart: rangeStartDate }), 'guard-nach-wartebudget', viewWaitMsSpent
        ));
        abortReason = 'view-not-ready';
        break;
      }

      // --- Zustand 3: Woche gerendert, aber falsch. Delta berechnen und per
      // Buttons navigieren. -------------------------------------------------
      const deltaWeeks = Math.round(
        (new Date(mondayIso + 'T00:00:00') - new Date(rendered.startDate + 'T00:00:00')) / (7 * 86400000)
      );
      if (deltaWeeks === 0) break; // Zielwoche gerendert -> unten bewerten

      if (buttonClicks >= MAX_BUTTON_CLICKS) { abortReason = 'max-clicks'; break; }

      if (Math.abs(deltaWeeks) > MAX_BUTTON_CLICKS) {
        // Zu weit entfernt für Button-Klicks: 1x startdate-URL als Fernsprung
        // (EIN Versuch, kein Beleg), danach wird wieder verifiziert — nur bei
        // ausreichendem Budget. Wirkt er nicht, bricht der Aufruf mit 'too-far'
        // ab (Lauf 5: URL korrekt, View unveraendert).
        if (!farJumpDone && remaining() >= 15000) {
          farJumpDone = true;
          const before = await readRenderedCalendarWeek(page);
          await gotoCalendar(optimizedUrl);
          const retry = await waitForCalendarRendered(phase(6000));
          rendered = retry.rendered;
          hasEvents = hasEvents || retry.hasEvents;
          navAttempts.push({
            kind: 'startdate-far-jump',
            url: page.url() || null,
            weekBefore: before.startDate || null,
            weekAfter: rendered.startDate || null,
            movedView: Boolean(rendered.startDate && rendered.startDate !== before.startDate)
          });
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
      // FIX 3: die Richtung kommt aus der DIFFERenz angefragte vs. angezeigte
      // Woche (oben berechnet) und wird verriegelt. Ein Wechsel der Richtung
      // waere kein gerichteter Weg zur Zielwoche — er entsteht nur, wenn der
      // Klick die Zielwoche uebersprungen hat, und wird dann ehrlich benannt.
      if (lockedDirection === null) {
        lockedDirection = direction;
      } else if (Math.sign(direction) !== Math.sign(lockedDirection)) {
        abortReason = 'week-overshoot';
        break;
      }
      if (!canClickSameDirection({ lastClickDirection, direction, weekConfirmedAfterClick })) {
        abortReason = 'same-direction-unconfirmed';
        break;
      }

      const r = await clickWeekButton(direction, rendered.startDate);
      if (r.guard === WEEK_CLICK_GUARD_REASONS.viewNotReady) {
        // BUG 5: Der frische Header-Read im Guard fand die Ansicht nicht
        // gerendert — und der Klickpfad hat davor bereits sein Wartebudget
        // verwendet. Das ist ZUSTAND 1, kein Abbruch: zurueck in die Schleife,
        // die dann wartet (und notfalls die startdate-URL hart neu laedt).
        // Bewusst VOR dem !r.clicked-Zweig, sonst wuerde es als
        // 'button-ineffective'/'no-week-button' fehlattet.
        rendered = r.week || rendered;
        continue;
      }
      // FIX 3: Die Zielwoche gilt nur als erreicht, wenn der ANGEZEIGTE Header
      // sie jetzt traegt (classifyClickEffect -> 'reached'), nicht weil sich
      // 'irgendetwas bewegt' hat.
      if (r.targetReached) {
        rendered = r.week;
        hasEvents = hasEvents || (await hasEventElements());
        break;
      }
      // FIX 3: der Klick hat die Woche in die Gegenrichtung bewegt bzw. mehr als
      // eine Woche gesprungen. Weiterzuklicken waere Blindklicken — der Abbruch
      // benennt die tatsaechliche Ursache.
      if (r.clicked && r.effect === CLICK_EFFECTS.overshoot) {
        rendered = r.week || rendered;
        abortReason = 'week-overshoot';
        break;
      }
      // FIX 3: der Klick ist am angezeigten Header NICHT bewertbar (vor oder
      // nach dem Klick unlesbar). In Lauf 5 hat genau das die Schleife am Leben
      // gehalten ('changed' ohne Zielbezug). Ohne Bewertung kein weiterer Klick.
      if (r.clicked && r.effect === CLICK_EFFECTS.unverified) {
        rendered = r.week || rendered;
        abortReason = 'click-unverified';
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
        budget: freeBusyBudgetReport(),
        headerReadError,
        navAttempts
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
        headerReadError,
        waitedForEventsMs: 0,
        eventWaitOutcome: 'not-rendered'
      });
      // BUG 5: der Text nennt jetzt beides, was geprueft wurde. "Keine
      // Event-Elemente" war nach dem datumsqualifizierten Umbau nicht mehr
      // dasselbe wie "kein Event-Knoten aus dem angefragten Zeitraum" — die
      // zweite Aussage ist die ehrliche.
      return errorResult(
        `Kalenderansicht nicht geladen — Kalender-Tab nicht erreicht oder Woche nicht gerendert. ` +
        `Weder ein datierter Wochen-Header noch ein Event-Knoten aus dem angefragten Zeitraum ` +
        `${rangeStartDate}–${rangeEndDate} gefunden (nach Navigation, Warten und timeboxter Eskalation). ` +
        `Bitte erneut versuchen bzw. Session prüfen (outlook_login).`,
        { renderedWeek: rendered, diagnostics }
      );
    }

    // Wochen-Nachweis (unabhaengig von den Events). FIX 1: der Nachweis
    // verlangt einen GELESENEN angezeigten Header, nicht irgendeine Ableitung.
    // Ein aus Prev/Next-Buttons abgeleiteter Wert (source:'nav-derivation') ist
    // eine Schaetzung und traegt keine 'weekVerified' — in Lauf 5 stand
    // weekVerified:true bei View auf KW 39, weil der next-Button-Label die
    // Zielwoche nannte.
    const weekVerified = isHeaderReadVerified(rendered);
    if (!weekVerified && rendered.readError) {
      // FIX 1: kein lesbarer angezeigter Header -> der Grund wird als
      // headerReadError berichtet und der Nav-Button-Text NICHT ersatzweise
      // genommen (der Button-Text steht in navOnlyCandidates/headerCandidates).
      headerReadError = rendered.readError;
    }

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
    //
    // BUG 4 (Live-Befund 26.09., Lauf 4): 'found' wird nur noch bei
    // DATUMSQUALIFIZIERTEM Beleg gemeldet. Live: 18 Event-Knoten der falschen
    // Woche (KW 39) beendeten die Wartephase nach 13ms/1 Poll mit 'found' bei
    // 20000ms Budget. Der Beobachter zaehlt deshalb die Labels aus dem
    // angefragten Zeitraum (mit DERSELBE Prueflogik wie die Abschluss-
    // Extraktion) und liest zusaetzlich den gerenderten Wochen-Header. Die
    // Range-Werte werden hier berechnet und NICHT erst weiter unten — die
    // Wartephase braucht sie vorher (TDZ-Falle, vgl. 0d47105).
    const rangeStartIso = String(startDt).includes('T') ? startDt : `${startDt}T00:00:00`;
    const rangeEndIso = String(endDt).includes('T') ? endDt : `${endDt}T23:59:59`;
    const rangeStartMs = new Date(rangeStartIso).getTime();
    const rangeEndMs = new Date(rangeEndIso).getTime();

    const eventWaitBudget = eventWaitBudgetMs(remaining());
    const observeCalendarEvents = makeEventObserver((preferSelector) => collectCalendarEventLabels(page, {
      preferSelector,
      weekStartIso: rendered.startDate || mondayIso,
      fast: true
    }), {
      readWeek: () => readRenderedCalendarWeek(page),
      range: { rangeStartMs, rangeEndMs },
      defaultYear
    });
    const eventWait = await waitForCalendarEvents({
      observe: observeCalendarEvents,
      budgetMs: eventWaitBudget,
      pollMs: FREEBUSY_EVENT_WAIT_POLL_MS,
      settleMs: FREEBUSY_EVENT_WAIT_SETTLE_MS,
      targetStart: rangeStartDate,
      sleep: (ms) => page.waitForTimeout(ms)
    });
    // BUG 4: `hasEvents` bleibt die DOM-Fakt ("irgendein Event-Knoten lesbar") und
    // steuert GUARD 2. Die datumsqualifizierte Trefferzahl steht getrennt davon in
    // eventWait (found/dateQualified) und darf die beiden nicht vermischen.
    hasEvents = hasEvents || eventWait.found;

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
    // BUG 4: `found` bleibt an die TATSAECHLICH gelesene Knotenzahl gekoppelt
    // (Ehrlichkeits-Gate, unveraendert). Die Datumsqualifikation ist eine
    // getrennte, additive Information: `dateQualified` sagt, OB der angefragte
    // Zeitraum belegt war — unabhängig davon, ob am Ende Knoten lesbar waren.
    hasEvents = hasEvents || eventWait.found;

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
    // FIX 1: der ANGEZEIGTE Wochen-Header ist kein 'header-misclassified'. Er
    // wird nicht verworfen, sondern als Header belegt ausgewiesen — vorher hat
    // das Werkzeug seinen eigenen Beleg gefunden und weggeschmissen.
    const weekHeaderLabels = eventCollection.weekHeaderLabels || [];
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
    const { kept: busySlots, rejected: rangeRejected } = partitionSlotsByRange(dedupedSlots, rangeStartMs, rangeEndMs);
    for (const r of rangeRejected) reject(r.slot && r.slot.raw ? r.slot.raw : '', r.reason);

    // Stabil nach Datum und Beginn sortieren
    busySlots.sort((a, b) =>
      `${a.date}T${a.start || '00:00'}`.localeCompare(`${b.date}T${b.start || '00:00'}`)
    );

    // FIX 4: der Widerspruch wird VOR dem Ergebnis berechnet, nicht danach im
    // Diagnostik-Block. In Lauf 2 stand busySlotsReliable:true GLEICHZEITIG mit
    // der eigenen Warnung "18 Event-Element(e) im DOM gefunden, aber 0
    // Busy-Slots extrahiert" — ein Selbstwiderspruch des Ergebnisses.
    // Die Invariante: sobald der Widerspruchstext entsteht, ist
    // busySlotsReliable ein hartes false.
    const extractionWarning = slotExtractionWarning(eventElementsFound, busySlots.length);
    const busySlotsReliable = resolveBusySlotsReliable({
      dataQuality, eventElementsFound, busySlotCount: busySlots.length
    });

    // Hinweise zusammenbauen (Verifikationsstatus + Range-Abdeckung)
    let note = 'Busy-Slots aus ARIA-Labels der Kalender-Event-Elemente der Wochenansicht extrahiert ' +
      '(Datum/Zeit strukturiert geparst, Work-Location-Pseudo-Slots entfernt, Zeit-Duplikate zusammengeführt, ' +
      'per Wortliste klassifiziert, auf den angefragten Zeitraum gefiltert). ' +
      'Navigation: der angezeigte Wochen-Header ist der EINZIGE Nachweis fuer die Zielwoche; ' +
      'der startdate-Parameter der URL ist ein VERSUCH (der Client-Store kann ihn ueberschreiben — ' +
      'belegt im Lauf vom 26.09.: korrekte URL, View unveraendert) und begruendet keine Daten. ' +
      'Buttons werden nur bei belegter Abweichung und gerichteter Differenz geklickt.';
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
        `(${eventWait.polls} Polls, davon ${eventWait.loadingPolls} mit Lade-Indikator` +
        `${eventWait.outsideRangePolls ? `, ${eventWait.outsideRangePolls} mit Events ausserhalb des angefragten Zeitraums` : ''}; ` +
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
      // Maschinenlesbare Ehrlichkeit: nur bei dataQuality==='verified' UND ohne
      // Extraktions-Widerspruch sind busySlots eine belastbare Aussage ueber
      // Belegung (FIX 4).
      busySlotsReliable,
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
        navAttempts,
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
      // BUG 4/5: Die Datumsqualifikation der Wartephase und das Ergebnis des
      // Zustand-1-Wartens sind die beiden Kennzahlen, an denen sich im naechsten
      // Live-Lauf sofort ablesen laesst, ob der Wartepfad gearbeitet hat.
      diag.eventWaitDateQualified = eventWait.dateQualified === true;
      diag.eventWaitInRangeLabels = eventWait.inRangeLabels === undefined ? null : eventWait.inRangeLabels;
      diag.eventWaitOutsideRangePolls = eventWait.outsideRangePolls || 0;
      // FIX 2: der eigene Zaehler fuer "datierter Header der falschen Woche".
      // Ohne dieses Feld laesst sich in Runde 6 nicht unterscheiden, ob der
      // Wartepfad an einem Header-Gegenbeleg oder an echter Leere gescheitert ist.
      diag.eventWaitHeaderOutsideRangePolls = eventWait.headerOutsideRangePolls || 0;
      if (eventWait.lastState) diag.eventWaitLastState = eventWait.lastState;
      diag.viewWaitMsSpent = Math.round(viewWaitMsSpent);
      // BUG 5: 'Header noch nicht gerendert' und 'Header-Auswertung fehlgeschlagen'
      // waren in Runde 4 nicht unterscheidbar (beide: leere Kandidatenlisten).
      if (rendered.readError) diag.headerReadError = rendered.readError;
      if (headerReadError) diag.headerReadError = headerReadError;
      if (eventWait.weekHeader && eventWait.weekHeader.readError) {
        diag.headerReadError = eventWait.weekHeader.readError;
      }
      if (eventCollection.sourceUsed) diag.sourceUsed = eventCollection.sourceUsed;
      if (eventCollection.sourceTried) diag.sourceTried = eventCollection.sourceTried;
      if (eventCollection.census) diag.census = eventCollection.census;
      if (eventCollection.collectError) diag.collectError = eventCollection.collectError;
      if (eventCollection.dateMissingLabels && eventCollection.dateMissingLabels.length) {
        diag.dateMissingLabels = eventCollection.dateMissingLabels;
      }
      if (weekHeaderLabels.length) diag.weekHeaderLabels = weekHeaderLabels;
      const warnings = [
        extractionWarning,
        eventWaitWarning({
          outcome: eventWait.outcome, waitedMs: eventWait.waitedMs, budgetMs: eventWaitBudget,
          abortReason: eventWait.abortReason,
          renderedWeek: rendered.startDate ? `${rendered.startDate}–${rendered.endDate || '?'}` : null
        })
      ].filter(Boolean);
      if (warnings.length) diag.warning = warnings.join(' ');
      result.diagnostics = diag;
    }
    return result;
  }

  // ===========================================================================
  // Teilnehmer-Verfügbarkeit (outlook_get_freebusy mit attendees-Parameter):
  // getAttendeeAvailability.
  //
  // Fragt die Busy-/Free-Belegung MEHRERER Kalender (Kollegen) für ein
  // Zeitfenster über das OWA-GraphQL-Gateway ab (Live-verifiziert 27.10.2026).
  // Der Bearer-Token kommt aus dem Eigen-Traffic der Kalender-App
  // (captureScheduleToken) — kein Passwort, keine Canary, session-gebunden.
  //
  // Ehrlichkeits-Garantien (Stil getFreeBusy):
  //  - Hartes Gesamtzeitbudget ATTENDEE_TOTAL_BUDGET_MS (40s, sicher unter dem
  //    MCP-Client-Timeout); alle Phasen deckeln sich gegen die Restzeit.
  //  - Fehler werden als strukturiertes Ergebnis (status:'error') MIT
  //    Diagnostik zurückgegeben, NICHT als nackter Throw geworfen — nur die
  //    reine Parametervalidierung wirft (wie getFreeBusy).
  //  - Pro Teilnehmer wird ein API-Fehler (z.B. keine Berechtigung/externes
  //    Postfach) einzeln berichtet, statt den ganzen Call zu killen.
  // ===========================================================================

  // Grifft den Bearer-Token (+ userId, Gateway-URL) aus dem Kalender-Boot-
  // Traffic ab. force=true umgeht den Cache (nach 401).
  async captureScheduleToken(page, { remaining, force = false, triggerUrl = null } = {}) {
    if (!force && isTokenUsable(this._scheduleToken, { nowMs: Date.now() }) && this._scheduleGraphqlUrl) {
      return {
        token: this._scheduleToken,
        userId: this._scheduleUserId,
        url: this._scheduleGraphqlUrl,
        source: 'cache'
      };
    }
    const budget = phaseBudgetMs(remaining(), ATTENDEE_TOKEN_WAIT_MAX_MS, 2000);
    if (budget <= 0) {
      return { token: null, error: 'Kein Zeitbudget mehr für den Token-Abgriff (Gesamtbudget aufgebraucht).' };
    }

    const captured = { token: null, userId: null, url: null, requests: [] };
    const onReq = (req) => {
      try {
        const u = req.url() || '';
        if (!u.includes('outlookgatewayb2/graphql')) return;
        if (captured.requests.length < 20) captured.requests.push(req);
        const auth = req.headers()['authorization'];
        if (auth && !captured.token) {
          captured.token = String(auth).replace(/^Bearer\s+/i, '');
          captured.url = u;
        }
        const pd = req.postData() || '';
        const m = /"userId"\s*:\s*"([^"]+)"/.exec(pd);
        if (m && !captured.userId) captured.userId = m[1];
      } catch (e) { /* Request kann zwischendrin verschwinden — ignorieren */ }
    };
    page.on('request', onReq);
    try {
      // Boot-Trigger: Kalender-Ansicht laden. OWA feuert beim Boot mindestens
      // einen GetSchedule-GraphQL-Call (Live-Befund Probe 5/6, 27.10.2026).
      const url = triggerUrl || (this.realm
        ? `https://outlook.office.com/calendar/view/week?realm=${encodeURIComponent(this.realm)}`
        : 'https://outlook.office.com/calendar/view/week');
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: Math.min(15000, budget) }).catch(() => null);
      const waitEnd = Date.now() + budget;
      while (!captured.token && Date.now() < waitEnd) {
        await page.waitForTimeout(400).catch(() => null);
      }
      // userId nachträglich aus den gecaptured Request-Bodies lesen (postData
      // ist im Request-Event teils noch nicht verfügbar, Live-Befund 27.10.).
      if (!captured.userId) {
        for (const req of captured.requests) {
          try {
            const pd = typeof req.postData === 'function' ? await req.postData() : req.postData();
            const m = /"userId"\s*:\s*"([^"]+)"/.exec(String(pd || ''));
            if (m) { captured.userId = m[1]; break; }
          } catch (e) { /* weiter mit dem nächsten Request */ }
        }
      }
    } finally {
      try { page.off('request', onReq); } catch (e) {}
    }

    if (!captured.token) {
      return {
        token: null,
        error: 'Kein GraphQL-Call mit Bearer-Token beim Kalender-Boot abgefangen. ' +
          'Session ggf. abgelaufen (outlook_login) oder Kalender-App nicht geladen.'
      };
    }
    this._scheduleToken = captured.token;
    this._scheduleTokenExpMs = parseJwtExp(captured.token) || 0;
    // userId-Kaskade: OWA-GetSchedule-Body -> JWT-Claims (upn/unique_name)
    // -> Fallback beim Aufruf (attendees[0], Live-verified im selben Tenant).
    if (!captured.userId) {
      const claims = parseJwtIdentity(captured.token);
      captured.userId = claims.upn || claims.uniqueName || null;
    }
    if (captured.url) {
      try {
        const u = new URL(captured.url);
        this._scheduleGraphqlUrl = u.origin + '/outlookgatewayb2/graphql';
      } catch (e) {
        this._scheduleGraphqlUrl = 'https://outlook.cloud.microsoft/outlookgatewayb2/graphql';
      }
    }
    if (captured.userId) this._scheduleUserId = captured.userId;
    return {
      token: this._scheduleToken,
      userId: this._scheduleUserId,
      url: this._scheduleGraphqlUrl,
      source: force ? 'refresh' : 'boot'
    };
  }

  async getAttendeeAvailability({ attendees, start, end, intervalMinutes = 15, timezone = 'Europe/Berlin' } = {}) {
    // --- Parametervalidierung (throw wie getFreeBusy) ---
    const { valid: validAttendees, invalid: invalidAttendees } = normalizeAttendeeList(attendees);
    if (!validAttendees.length) {
      throw new Error('Parameter "attendees" ist erforderlich: mindestens EINE gültige E-Mail-Adresse ' +
        `(z.B. ["silas.beckmann@d-velop.de"]). Ungültig: ${invalidAttendees.length ? invalidAttendees.map(a => `"${a}"`).join(', ') : 'keine Eingabe'}`);
    }
    if (!start || !end) {
      throw new Error('Parameter "start" und "end" sind erforderlich.');
    }
    let interval = Number.parseInt(intervalMinutes, 10);
    if (!Number.isFinite(interval)) interval = 15;
    interval = Math.min(60, Math.max(5, interval));

    // Fenster -> UTC (naive Angaben gelten als Ziel-Zeitzone, Default Europe/Berlin)
    const startLocal = String(start).length === 10 ? `${start}T00:00:00` : String(start);
    const endLocal = String(end).length === 10 ? `${end}T23:59:59.999` : String(end);
    const startUtcIso = zonedTimeToUtcIso(startLocal, timezone);
    const endUtcIso = zonedTimeToUtcIso(endLocal, timezone);
    if (!startUtcIso || !endUtcIso) {
      throw new Error(`Ungültige Zeitangaben: start="${start}", end="${end}" (ISO erwartet, z.B. "2026-10-28T09:30:00").`);
    }
    const startMs = Date.parse(startUtcIso);
    const endMs = Date.parse(endUtcIso);
    if (endMs <= startMs) {
      throw new Error(`"end" (${endUtcIso}) muss nach "start" (${startUtcIso}) liegen.`);
    }

    const budgetStart = Date.now();
    const deadline = budgetStart + ATTENDEE_TOTAL_BUDGET_MS;
    const remaining = () => deadline - Date.now();

    const errorResult = (message, extra = {}) => ({
      status: 'error',
      error: message,
      attendees: validAttendees,
      requestedRange: { start, end, startUtc: startUtcIso, endUtc: endUtcIso, timezone },
      diagnostics: {
        totalBudgetMs: ATTENDEE_TOTAL_BUDGET_MS,
        usedMs: Date.now() - budgetStart,
        remainingMs: Math.max(0, remaining()),
        ...extra
      }
    });

    // Seite + Session
    let page;
    try {
      page = await browserManager.ensureContext(true, this.realm).then(r => r.page);
      this.lastPage = page;
    } catch (e) {
      return errorResult(`Browser-Kontext nicht verfügbar: ${e.message}`);
    }

    // Token-Abgriff (Cache -> Boot-Traffic)
    let cap = await this.captureScheduleToken(page, { remaining });
    if (!cap.token) {
      return errorResult(`Kein gültiger Auth-Token für das Outlook-Gateway: ${cap.error || 'unbekannter Grund'}`, { tokenSource: cap.source || null });
    }

    const requestBody = buildGetScheduleRequest({
      userId: cap.userId, attendeeEmails: validAttendees, startUtcIso, endUtcIso, intervalMinutes: interval
    });

    // POST (ein Retry bei 401/403 mit Token-Refresh)
    const context = page.context();
    const postSchedule = async () => {
      const httpTimeout = Math.min(ATTENDEE_HTTP_TIMEOUT_MS, Math.max(3000, remaining()));
      return context.request.post(this._scheduleGraphqlUrl, {
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${this._scheduleToken}`,
          'x-ms-appname': 'owa-reactcalendar'
        },
        data: requestBody,
        timeout: httpTimeout
      });
    };

    let resp = null;
    try {
      resp = await postSchedule();
      if (resp.status() === 401 || resp.status() === 403) {
        // Token abgelaufen/ungültig -> einmal frisch abgreifen und retryen.
        cap = await this.captureScheduleToken(page, { remaining, force: true });
        if (!cap.token) {
          return errorResult(`Auth-Token-Refresh fehlgeschlagen (HTTP ${resp.status()}): ${cap.error || 'unbekannter Grund'}`, { httpStatus: resp.status(), tokenSource: 'refresh-failed' });
        }
        resp = await postSchedule();
      }
    } catch (e) {
      return errorResult(`Gateway-POST fehlgeschlagen: ${e.message}`, { tokenSource: cap.source || null });
    }

    const httpStatus = resp.status();
    const bodyText = await resp.text().catch(() => '');
    if (httpStatus !== 200) {
      return errorResult(`Gateway-Antwort HTTP ${httpStatus}: ${bodyText.slice(0, 300) || '(leerer Body)'}`, { httpStatus, tokenSource: cap.source || null });
    }

    const parsed = parseGetScheduleResponse(bodyText);
    if (!parsed.ok) {
      return errorResult(`Gateway-Antwort nicht verwertbar: ${parsed.errorInfo.kind}`, {
        httpStatus, tokenSource: cap.source || null, errorInfo: parsed.errorInfo
      });
    }

    // Pro Teilnehmer dekodieren
    const attendeeResults = [];
    const busyPerAttendee = [];
    for (const s of parsed.schedules) {
      const email = s.scheduleId || s.email?.address || '(unbekannt)';
      const apiError = (s.error && (s.error.message || s.error.responseCode))
        ? { message: s.error.message || null, responseCode: s.error.responseCode || null }
        : null;
      const decoded = decodeAvailabilityView(s.availabilityView, { startUtcIso, endUtcIso, intervalMinutes: interval });
      busyPerAttendee.push(apiError ? [] : decoded.busySegments);
      const minutesByStatus = {};
      for (const seg of decoded.segments) minutesByStatus[seg.status] = (minutesByStatus[seg.status] || 0) + seg.minutes;
      attendeeResults.push({
        email,
        error: apiError,
        availabilityView: s.availabilityView || '',
        slotMinutes: interval,
        busySlots: decoded.busySegments.map(seg => ({
          status: seg.status,
          start: seg.startUtcIso,
          end: seg.endUtcIso,
          startLocal: formatUtcIsoInZone(seg.startUtcIso, timezone),
          endLocal: formatUtcIsoInZone(seg.endUtcIso, timezone)
        })),
        minutesByStatus,
        scheduleItems: (Array.isArray(s.scheduleItems) ? s.scheduleItems : []).slice(0, 20).map(it => ({
          status: it.status || null,
          start: it.startTime?.dateTime || null,
          end: it.endTime?.dateTime || null,
          subject: it.isPrivate ? null : (it.subject || null),
          isPrivate: it.isPrivate === true
        })),
        workingHours: s.workingHours || null
      });
    }

    const commonFreeSlots = intersectFreeSegments(busyPerAttendee, {
      startUtcIso, endUtcIso, minDurationMinutes: interval
    });

    const failed = attendeeResults.filter(a => a.error);
    const note = 'Verfügbarkeit über das OWA-GraphQL-Gateway (GetSchedule) abgefragt — ' +
      'Autorisierung über den Session-Token der eingeloggten Kalender-App. ' +
      'Zeiten in UTC (start/end mit Z) + Lokalzeit (' + timezone + '). ' +
      'commonFreeSlots = Schnittmenge aller freien Intervalle (>= ' + interval + ' Min). ' +
      (failed.length
        ? `WARNUNG: ${failed.length} Teilnehmer hatte(n) einen API-Fehler (${failed.map(f => `${f.email}: ${f.error.responseCode || f.error.message}`).join('; ')}) — deren Belegung ist NICHT enthalten.`
        : 'Alle Teilnehmer ohne API-Fehler.');

    return {
      status: 'ok',
      start: startLocal,
      end: endLocal,
      window: { startUtc: startUtcIso, endUtc: endUtcIso, timezone },
      intervalMinutes: interval,
      attendees: attendeeResults,
      commonFreeSlots,
      count: attendeeResults.length,
      note,
      diagnostics: {
        totalBudgetMs: ATTENDEE_TOTAL_BUDGET_MS,
        usedMs: Date.now() - budgetStart,
        httpStatus,
        tokenSource: cap.source || null,
        tokenExpiresAt: this._scheduleTokenExpMs ? new Date(this._scheduleTokenExpMs).toISOString() : null,
        userId: cap.userId || this._scheduleUserId || null,
        attendeeCount: validAttendees.length,
        invalidAttendees
      }
    };
  }

  // ===========================================================================
  // READ-ONLY Kalender-LESE-Tool (gbrain #5916 / otto Termin-Briefings):
  // outlook_list_calendar_events.
  //
  // Liste ALLER Termine (nicht nur Busy-Slots) im angefragten Zeitraum —
  // Titel, Datum, Start/Ende, Ganztag-Flag, Status. Bewusst READ-ONLY: es
  // wird nur navigiert/gelesen, nichts angelegt, geaendert oder geloescht.
  //
  // Implementierung: nutzt die SELBEN verifizierten Navigations-Primitive wie
  // getFreeBusy (weekViewState/waitForWeekViewReady/runWeekButtonClick), aber
  // mit einem schlankeren, aufs Lesen zugeschnittenen Pfad. Im Gegensatz zu
  // getFreeBusy werden ALLE Event-Knoten zurueckgegeben (busy, free,
  // tentative, Ganztag) statt nur die Busy-Slots.
  //
  // Ehrlichkeits-Garantien (im Stil der restlichen Kalender-Tools):
  //  - Die Zielwoche wird ausschliesslich am ANGEZEIGTEN Wochen-Header
  //    verifiziert (renderedWeekContains). Falsche Wochen liefern KEINE Daten.
  //  - dataQuality 'verified' nur, wenn die Woche verifiziert UND Event-Knoten
  //    gelesen wurden; sonst wird das ehrlich als unzuverlaessig ausgewiesen.
  //  - Pseudo-Slots (Arbeitsort, "(ohne Titel)", URLs) werden entfernt.
  //  - timespanText nennt den angefragten Zeitraum als Vorlage fuer Briefings.
  // ===========================================================================
  async listCalendarEvents({ start, end } = {}) {
    if (!start || !end) {
      throw new Error('Parameter "start" und "end" sind erforderlich.');
    }

    const page = await browserManager.ensureContext(true, this.realm).then(r => r.page);
    this.lastPage = page;

    const startDt = normalizeCalendarDate(start) || start;
    const endDt = normalizeCalendarDate(end) || end;
    const rangeStartDate = String(startDt).slice(0, 10);
    const rangeEndDate = String(endDt).slice(0, 10);
    const defaultYear = parseInt(rangeStartDate.slice(0, 4), 10);

    // Zielwoche: Montag der Woche, die das Startdatum enthaelt (OWA beginnt montags).
    const startDateObj = new Date(rangeStartDate + 'T00:00:00');
    const weekday = startDateObj.getDay(); // 0 = Sonntag
    const offsetToMonday = weekday === 0 ? -6 : 1 - weekday;
    const mondayIso = toIsoDate(new Date(startDateObj.getTime() + offsetToMonday * 86400000));

    const optimizedUrl = buildCalendarWeekUrl(this.realm, mondayIso);
    const provenUrl = this.realm
      ? `https://outlook.office.com/calendar/view/week?realm=${encodeURIComponent(this.realm)}`
      : 'https://outlook.office.com/calendar/view/week';

    // Hartes Gesamtbudget (sicher unter dem MCP-Client-Timeout), analog getFreeBusy.
    const budgetStart = Date.now();
    const deadline = budgetStart + 45000;
    const remaining = () => deadline - Date.now();
    const phase = (capMs) => phaseBudgetMs(remaining(), capMs);
    const CLICK_COST_MS = 1500 + FREEBUSY_CLICK_RENDER_WAIT_MS;

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

    let onCalendar = await gotoCalendar(optimizedUrl);
    if (!onCalendar) onCalendar = await gotoCalendar(provenUrl);
    if (!onCalendar) {
      throw new Error(`Kalender-Tab konnte nicht erreicht werden (URL: ${page.url() || 'unbekannt'}). ` +
        'Möglicherweise ist die Session abgelaufen — bitte outlook_login erneut ausführen.');
    }

    let rendered = { text: null, startDate: null, endDate: null };
    // Zustandswarten auf gerenderten Wochen-Header/Event-Grid (500ms Poll).
    const EVENT_GRACE_MS = 2500;
    const waitForCalendarRendered = async (budgetMs) => {
      const waitDeadline = Date.now() + Math.min(budgetMs, Math.max(0, remaining()));
      let headerSeenAt = null;
      let eventsSeenAt = null;
      let events = false;
      while (Date.now() < waitDeadline) {
        const week = await readRenderedCalendarWeek(page);
        const probe = await collectCalendarEventLabels(page, {
          preferSelector: null, weekStartIso: mondayIso, fast: true
        });
        events = probe.eventElementsFound > 0;
        if (week.startDate) {
          if (!headerSeenAt) headerSeenAt = Date.now();
          if (events || Date.now() - headerSeenAt >= EVENT_GRACE_MS) {
            rendered = week;
            return { rendered: week, hasEvents: events };
          }
        } else if (events) {
          if (!eventsSeenAt) eventsSeenAt = Date.now();
          if (Date.now() - eventsSeenAt >= EVENT_GRACE_MS) {
            rendered = week;
            return { rendered: week, hasEvents: events };
          }
        }
        if (Date.now() + 500 > waitDeadline) break;
        await page.waitForTimeout(500);
      }
      rendered = await readRenderedCalendarWeek(page);
      return { rendered, hasEvents: events };
    };

    await waitForCalendarRendered(phase(8000));

    let hasEvents = Boolean(rendered.startDate);
    const weekButtonProbe = [];
    let viewWaitMsSpent = 0;
    const waitForReady = async () => {
      const w = await waitForWeekViewReady({
        readWeek: () => readRenderedCalendarWeek(page),
        readEvents: () => collectCalendarEventLabels(page, { weekStartIso: mondayIso, fast: true })
          .then((r) => ({ loading: r.loading === true })),
        budgetMs: phase(FREEBUSY_VIEW_WAIT_BUDGET_MS),
        pollMs: FREEBUSY_EVENT_WAIT_POLL_MS,
        settleMs: FREEBUSY_EVENT_WAIT_SETTLE_MS,
        targetStart: rangeStartDate,
        sleep: (ms) => page.waitForTimeout(ms)
      });
      viewWaitMsSpent += Math.round(w.waitedMs || 0);
      rendered = w.rendered || rendered;
      hasEvents = hasEvents || w.state === WEEK_VIEW_STATES.targetWeek;
      return w;
    };
    const waitRenderAfterClick = async (prevStart) => {
      const waitDeadline = Date.now() + Math.min(FREEBUSY_CLICK_RENDER_WAIT_MS, Math.max(0, remaining()));
      let week = { text: null, startDate: null, endDate: null };
      while (Date.now() < waitDeadline) {
        week = await readRenderedCalendarWeek(page);
        if (week.startDate) break;
        if (Date.now() + 500 > waitDeadline) break;
        await page.waitForTimeout(500);
      }
      return { week, verified: evaluateClickVerification(prevStart, week) };
    };

    // Gerichtete, verifizierte, begrenzte Wochen-Navigation (max. 2 Klicks).
    let abortReason = 'none';
    let lastClickDirection = null;
    let weekConfirmedAfterClick = false;
    const MAX_BUTTON_CLICKS = 2;
    let buttonClicks = 0;
    for (let step = 0; step < 30; step++) {
      const view = weekViewState({ rendered, targetStart: rangeStartDate });
      if (view.state === WEEK_VIEW_STATES.targetWeek) break;
      if (remaining() <= 0) { abortReason = 'budget'; break; }
      if (view.state === WEEK_VIEW_STATES.notReady) {
        await waitForReady();
        if (weekViewState({ rendered, targetStart: rangeStartDate }).state === WEEK_VIEW_STATES.targetWeek) break;
        abortReason = 'view-not-ready';
        break;
      }
      const deltaWeeks = Math.round(
        (new Date(mondayIso + 'T00:00:00') - new Date(rendered.startDate + 'T00:00:00')) / (7 * 86400000)
      );
      if (deltaWeeks === 0) break;
      if (buttonClicks >= MAX_BUTTON_CLICKS) { abortReason = 'max-clicks'; break; }
      if (Math.abs(deltaWeeks) > MAX_BUTTON_CLICKS) { abortReason = 'too-far'; break; }
      if (remaining() < CLICK_COST_MS) { abortReason = 'budget'; break; }
      const direction = Math.sign(deltaWeeks);
      if (!canClickSameDirection({ lastClickDirection, direction, weekConfirmedAfterClick })) {
        abortReason = 'same-direction-unconfirmed';
        break;
      }
      const r = await runWeekButtonClick({
        direction,
        prevStart: rendered.startDate,
        targetStart: rangeStartDate,
        page,
        probe: weekButtonProbe,
        readWeek: () => readRenderedCalendarWeek(page),
        afterClick: (prev) => waitRenderAfterClick(prev),
        clickWaitMs: FREEBUSY_CLICK_RENDER_WAIT_MS,
        clickCostMs: CLICK_COST_MS,
        remainingMs: remaining,
        waitForReady
      });
      if (r.guard === WEEK_CLICK_GUARD_REASONS.viewNotReady) {
        rendered = r.week || rendered;
        continue;
      }
      if (r.targetReached) { rendered = r.week; break; }
      if (r.clicked && r.effect === CLICK_EFFECTS.overshoot) { abortReason = 'week-overshoot'; break; }
      if (r.clicked && r.effect === CLICK_EFFECTS.unverified) { abortReason = 'click-unverified'; break; }
      if (!r.clicked) { abortReason = 'button-ineffective'; break; }
      buttonClicks++;
      rendered = r.week;
      lastClickDirection = direction;
      weekConfirmedAfterClick = Boolean(r.week && r.week.startDate);
    }

    // GUARD 1: falsche Woche -> keine Daten statt falscher Daten.
    if (rendered.startDate && !renderedWeekContains(rendered, rangeStartDate)) {
      throw new Error(
        `Kalender-Wochenansicht zeigt nicht die angefragte Woche: angefragt ab ${rangeStartDate}, ` +
        `gerendert ${rendered.startDate} – ${rendered.endDate || '?'}. Es werden bewusst keine Daten ` +
        `der falschen Woche zurueckgegeben.`
      );
    }
    const weekVerified = isHeaderReadVerified(rendered);

    // Event-Knoten sammeln und alle parsen (nicht nur Busy — fuer Briefings).
    const eventCollection = await collectCalendarEventLabels(page, { weekStartIso: rendered.startDate || mondayIso });
    const eventElementsFound = eventCollection.eventElementsFound;
    const dataQuality = resolveDataQuality({ weekVerified, eventElementsFound });

    const seen = new Set();
    const parsed = [];
    for (const label of eventCollection.labels) {
      if (seen.has(label)) continue;
      seen.add(label);
      const slot = parseEventLabel(label, defaultYear);
      if (!slot) continue;
      if (isWorkLocationPseudoSlot(slot.title)) continue;
      slot.status = classifyEventStatus(label);
      parsed.push(slot);
    }
    const deduped = dedupeBusySlots(parsed);

    // Alle Events im angefragten Zeitraum behalten (auch 'free'/Ganztag):
    // partitionSlotsByRange verwirft 'free' — fuer die reine Liste bleiben wir
    // daher bei einer eigenen, status-unabhaengigen Range-Filterung.
    const rangeStartMs = new Date(String(startDt).includes('T') ? startDt : `${rangeStartDate}T00:00:00`).getTime();
    const rangeEndMs = new Date(String(endDt).includes('T') ? endDt : `${rangeEndDate}T23:59:59`).getTime();
    const events = deduped.filter((slot) => {
      const dayStartMs = new Date(`${slot.date}T00:00:00`).getTime();
      const s = slot.allDay ? dayStartMs : new Date(`${slot.date}T${slot.start}:00`).getTime();
      const e = slot.allDay ? dayStartMs + 86400000 : new Date(`${slot.date}T${slot.end}:00`).getTime();
      return s < rangeEndMs && e > rangeStartMs;
    });

    // Stabil nach Datum und Beginn sortieren.
    events.sort((a, b) =>
      `${a.date}T${a.start || '00:00'}`.localeCompare(`${b.date}T${b.start || '00:00'}`)
    );

    const result = {
      status: 'ok',
      calendar: 'Firmenkalender (Outlook Web)',
      timespanText: `${rangeStartDate} bis ${rangeEndDate}`,
      renderedWeek: {
        text: rendered.text,
        startDate: rendered.startDate,
        endDate: rendered.endDate
      },
      weekVerified,
      dataQuality,
      events,
      count: events.length,
      note: `READ-ONLY Liste aller Termine (busy/free/tentative/Ganztag) aus der OWA-Wochenansicht ` +
        `für ${rangeStartDate} bis ${rangeEndDate}. Pseudo-Slots (Arbeitsort, ohne Titel, URLs) entfernt, ` +
        `Duplikate zusammengeführt, nach Woche sortiert.`
    };
    if (dataQuality !== 'verified') {
      result.note += ' WARNUNG: Woche/Event-Grid nicht vollständig verifizierbar — Liste ggf. unvollständig.';
    }
    if (abortReason !== 'none') {
      result.note += ` HINWEIS: Navigation vorzeitig beendet (${abortReason}).`;
    }
    if (eventElementsFound > 0 && events.length === 0) {
      result.note += ` HINWEIS: ${eventElementsFound} Event-Knoten gelesen, aber keine im angefragten Zeitraum.`;
    }
    return result;
  }

  // READ-ONLY Detail-Lesen eines Kalender-Events fuer Termin-Briefings:
  // outlook_get_calendar_event. Findet das Event per Titel-Fragment in der
  // Wochenansicht (openCalendarEventByQuery) und liefert den Detail-Text des
  // Reading-Panels zurueck. BEWUSST KEINE Schreiboperation: es wird nur ein
  // Doppelklick zum Oeffnen des (lesenden) Detail-Panels ausgefuehrt.
  //
  // Rueckgabe: { status, query, titleMatched, detailText } oder wirft.
  async getCalendarEvent({ query } = {}) {
    if (!query || typeof query !== 'string' || query.trim().length === 0) {
      throw new Error('Parameter "query" (Suchbegriff/Titel-Fragment) ist erforderlich.');
    }
    const { page, titleMatched, detailText } = await this.openCalendarEventByQuery(query);
    return {
      status: 'ok',
      calendar: 'Firmenkalender (Outlook Web)',
      query,
      titleMatched,
      detailText,
      note: 'READ-ONLY Detail-Ansicht aus dem OWA-Reading-Panel. Es wurde nichts geaendert.'
    };
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
    let page = browserManager.getPage(this.realm);
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

  // P6 (Befund 15.10.2026): 4x "Target page closed" / "navigation destroyed"
  // pro Session — fragile Reads laufen teils mitten in OWA-Client-Navigationen.
  // Bewährter Workaround aus der Live-Session: Browser schließen -> neu starten
  // -> Read wiederholen. Diese Methode kapselt genau das: transient erkennbare
  // Fehler (Target closed, navigation destroyed, Execution context destroyed)
  // führen zu max. `retries` Wiederholungen MIT sauberem Browser-Restart; das
  // Persistent-Profile behält die Anmeldung. Andere Fehler gehen sofort durch.
  async _withReadRetry(label, fn, { retries = 2 } = {}) {
    // Explizite Varianten: Playwright meldet denselben Abbruch als
    // "Target page closed", "Target closed" oder "Target page, context or
    // browser has been closed" — alle drei plus die Evaluate-/Navigation-
    // Formen sind transient (Befund 15.10.2026).
    const TRANSIENT = /target page closed|target closed|target page, context or browser has been closed|page has been closed|browser has been closed|session closed|navigation destroyed|execution context was destroyed|net::err_aborted/i;
    let lastErr = null;
    for (let attempt = 0; attempt <= retries; attempt++) {
      if (attempt > 0) {
        // Sauberer Restart vor dem nächsten Versuch (Session bleibt im Profil)
        await browserManager.close().catch(() => null);
        await new Promise((r) => setTimeout(r, 1500));
      }
      try {
        return await fn(attempt);
      } catch (e) {
        lastErr = e;
        const msg = String((e && e.message) || e);
        if (!TRANSIENT.test(msg) || attempt === retries) break;
      }
    }
    if (lastErr) {
      lastErr.message = `${label} (nach ${retries} Restart-Retries): ${lastErr.message}`;
    }
    throw lastErr;
  }

  // Oeffentlicher Read-Pfad mit Restart-Retry (P6): fragile Such-Reads laufen
  // mitten in OWA-Client-Navigationen.
  async searchEmails(query, maxResults = 10) {
    return this._withReadRetry('outlook_search', () => this._searchEmailsInner(query, maxResults));
  }

  async _searchEmailsInner(query, maxResults = 10) {
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
    // Sensibles Settle-Wait: die Ergebnisliste re-rendert nach dem Selector-
    // Treffer noch (mid-navigation-Reads waren die Hauptursache von
    // "navigation destroyed" in page.evaluate).
    await page.waitForTimeout(800);

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

  // P5 (Befund 15.10.2026): Ursprungsmails fehlten BOTH in der OWA-Suche ALS AUCH
  // im sentitems-Enumerationsergebnis des Aufrufers — der Reply-Draft entstand
  // ohne Thread-Verlauf. Dieser Fallback enumeriert den Sent-Ordner im
  // Datumsfenster und filtert nach Empfänger/Begriffen direkt auf den Zeilentexten.
  async enumerateSent({ to, query, dateFrom, dateTo, maxScrolls = 200 } = {}) {
    const page = await this.getPage(true);
    await this.navigateFolder('sentitems');
    const seen = new Map();
    let stalled = 0;
    const stopBefore = dateFrom ? parseFlexibleDateInput(dateFrom) : null;
    const stopAfter = dateTo ? parseFlexibleDateInput(dateTo) : null;
    const inWindow = (ts) => {
      if (ts === null) return true; // ohne parsebares Datum nicht ausschließen
      const d = new Date(ts);
      const ymd = { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() };
      const cmp = (b) => ymd.y !== b.y ? ymd.y - b.y : (ymd.m !== b.m ? ymd.m - b.m : ymd.d - b.d);
      if (stopBefore && cmp(stopBefore) < 0) return false;
      if (stopAfter && cmp(stopAfter) > 0) return false;
      return true;
    };

    let stopEarly = false;
    for (let i = 0; i < maxScrolls && !stopEarly; i++) {
      const beforeCount = seen.size;
      const rows = await page.$$('div[data-convid]');
      for (const r of rows) {
        const id = await r.getAttribute('data-convid');
        if (!id || seen.has(id)) continue;
        const text = ((await r.innerText().catch(() => '')) || '');
        const ts = parseMailRowTimestamp(text);
        // Liste ist neueste-zuerst sortiert: Rows vor dem Fensterstart überspringen,
        // unter dem Fensterrand kann die Enumeration abgebrochen werden.
        if (stopBefore && ts !== null && !inWindow(ts)) {
          const d = new Date(ts);
          const ymd = { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() };
          const cmp = (b) => ymd.y !== b.y ? ymd.y - b.y : (ymd.m !== b.m ? ymd.m - b.m : ymd.d - b.d);
          if (cmp(stopBefore) < 0) { stopEarly = true; break; }
          continue;
        }
        const rowText = text.slice(0, 600);
        const matchTo = to ? evidenceMatchesExpected(rowText, to) : null;
        const matchQuery = query ? evidenceMatchesExpected(rowText, query) : null;
        const matchAll = (matchTo === null || matchTo === true) && (matchQuery === null || matchQuery === true);
        if (!matchAll) continue;
        const lines = text.split('\n').map((l) => l.trim()).filter(Boolean);
        seen.set(id, {
          index: seen.size,
          id,
          previewText: lines.join(' | ').slice(0, 300),
          timestampIso: ts !== null ? new Date(ts).toISOString() : null
        });
      }
      if (stopEarly) break;

      // Viewport-Step in der virtualisierten Liste (gleiche Technik wie bulkList)
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
      if (seen.size === beforeCount) { stalled++; } else { stalled = 0; }
      if (stalled >= 15) break;
    }

    const results = [...seen.values()];
    return { folder: 'sentitems', count: results.length, results };
  }

  // P5: Suche mit Fallback-Kette. 1) Normale OWA-Suche. 2) Bei Leere: Sent-
  // Enumeration (Datumsfenster + Empfänger-Filter), angestoßen durch folder=,
  // to= oder eine E-Mail-Adresse im Query. 3) Wenn weiterhin nichts: KLARER
  // "NOT FOUND"-Return mit Strategie-Hinweis — der aufrufende Agent kann damit
  // umgehen (z.B. Reply ohne Thread-Verlauf bewusst erstellen).
  async searchWithFallback(query, { limit = 10, folder, to, dateFrom, dateTo } = {}) {
    const trimmedQuery = String(query || '').trim();

    if (trimmedQuery) {
      try {
        const primary = await this.searchEmails(trimmedQuery, limit);
        if (primary.count > 0) return { ...primary, strategy: 'owa-search' };
      } catch (e) {
        // OWA-Suche selbst fehlgeschlagen -> Fallback weiter versuchen
      }
    }

    const emailInQuery = extractEmailLike(trimmedQuery);
    const folderWantsSent = ['sentitems', 'sent', 'gesendet'].includes(String(folder || '').toLowerCase());
    if (folderWantsSent || to || emailInQuery) {
      const sent = await this.enumerateSent({
        to: to || emailInQuery || undefined,
        query: trimmedQuery && !emailInQuery ? trimmedQuery : (trimmedQuery && to ? trimmedQuery : undefined),
        dateFrom,
        dateTo
      });
      if (sent.count > 0) {
        return {
          query: trimmedQuery || to,
          count: sent.count,
          results: sent.results.slice(0, limit),
          strategy: 'sentitems-enumeration',
          note: 'OWA-Suche lieferte nichts; Treffer stammen aus der Datumsfenster-Enumeration des Sent-Ordners.'
        };
      }
      return {
        query: trimmedQuery || to,
        count: 0,
        results: [],
        strategy: 'sentitems-enumeration',
        note: 'NOT FOUND: OWA-Suche und Datumsfenster-Enumeration im Sent-Ordner lieferten keinen Treffer. Die Mail ist ggf. aus einem anderen Postfach versendet oder im Sent-Ordner noch nicht angekommen (Sync-Lücke).'
      };
    }

    return {
      query: trimmedQuery,
      count: 0,
      results: [],
      strategy: 'owa-search',
      note: 'NOT FOUND: OWA-Suche lieferte keine Treffer. Tipp: Bei vermutlich versendeten Mails den Sent-Ordner durchsuchen (folder="sentitems", ggf. to="<Empfängeradresse>" und date_from/date_to).'
    };
  }

  // Oeffentlicher Read-Pfad mit Restart-Retry (P6): "Target page closed"/
  // "navigation destroyed" traten v.a. beim Lesebereichs-Read auf; der
  // Live-Workaround (close -> search -> get_email(index)) ist hier verallgemeinert.
  async getEmailDetails(indexOrId = 0) {
    return this._withReadRetry('outlook_get_email', () => this._getEmailDetailsInner(indexOrId));
  }

  async _getEmailDetailsInner(indexOrId = 0) {
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

  // Oeffentlicher Read-Pfad mit Restart-Retry (P6): Scroll-Enumeration +
  // Lesebereichs-Read sind die laengsten fragilen Sequenzen. Das harte
  // Zeitlimit gilt pro Versuch (nach Restart wird vollstaendig neu gezaehlt).
  async readEmailByConvId(convid, folder, { timeoutMs = 60000 } = {}) {
    return this._withReadRetry('outlook_read_email', () => this._readEmailByConvIdInner(convid, folder, { timeoutMs }));
  }

  async _readEmailByConvIdInner(convid, folder, { timeoutMs = 60000 } = {}) {
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

  // ---------------------------------------------------------------------------
  // Anlagen-Handling (Tool-Lücke geschlossen, 12.10.2026)
  // ---------------------------------------------------------------------------
  // gbrain #7330: Anlagen (v.a. Screenshot-Blobs) waren in OWA nur als Thumbnail
  // sichtbar und per Tool nicht extrahierbar. Zwei UI-getriebene Wege, passend
  // zum Bridge-Ansatz (echter OWA-Client, keine Graph-API):
  //   A) Anlagen-Chips im Lesebereich -> Hover -> Download-Trigger -> Playwright-
  //      Download-Event (browser.js: acceptDownloads: true) -> saveAs()
  //   B) Inline-Bilder im Body (blob:/data:) via in-page fetch -> base64 -> Datei.
  //      Nur als Fallback, wenn kein Chip-Download gelang (vermeidet Duplikate).

  // Oeffnet die Ziel-E-Mail im Lesebereich (Conversation-ID inkl. Scroll-
  // Enumeration oder Listen-Index) und liefert die Playwright-Page zurueck.
  async _openEmailForAttachments({ index = 0, id, folder, timeoutMs = 60000 } = {}) {
    const page = await this.getPage(true);
    if (folder) {
      await this.navigateFolder(folder);
    }
    await page.waitForSelector('div[data-convid], div[role="option"]', { timeout: 15000 });
    const paneSel = 'div[role="main"], div[aria-label="Reading Pane"], div[aria-label="Lesebereich"], div.ItemPartView';

    if (id) {
      const deadline = Date.now() + timeoutMs;
      const rowLocator = page.locator(`div[data-convid="${String(id).replace(/["\\]/g, '\\$&')}"]`).first();
      while (Date.now() <= deadline) {
        if ((await rowLocator.count().catch(() => 0)) > 0) {
          await rowLocator.click();
          await page.waitForTimeout(2200);
          await page.waitForSelector(paneSel, { timeout: 15000 });
          return page;
        }
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
      throw new Error(`Timeout: Conversation-ID "${id}" wurde für den Anlagenzugriff nicht gefunden (Scroll-Enumeration).`);
    }

    const rows = await page.$$('div[data-convid], div[role="option"]');
    if (rows.length === 0) throw new Error('Keine E-Mails in der Liste gefunden.');
    const target = rows[Math.min(Math.max(0, index), rows.length - 1)];
    await target.click();
    await page.waitForTimeout(2000);
    await page.waitForSelector(paneSel, { timeout: 15000 });
    return page;
  }

  // Scannt den Lesebereich nach Anlagen-Chips und Inline-Bildern. Reine Daten-
  // rueckgabe ohne Klicks — gemeinsame Basis fuer list + save.
  async _scanAttachments(page) {
    return await page.evaluate(() => {
      const pane = document.querySelector('div[role="main"], div[aria-label="Reading Pane"], div[aria-label="Lesebereich"], div.ItemPartView');
      if (!pane) return { error: 'Lesebereich nicht gefunden' };
      const extRe = /\.(png|jpe?g|gif|bmp|webp|heic|tif|tiff|pdf|docx?|xlsx?|pptx?|one|zip|7z|rar|msg|eml|txt|csv|ics|vcf|html?)($|\s|\()/i;
      const typeRe = /(\boutlook\s*item\b|\be-?mail-?nachricht\b|\bnachrichtenelement\b|\bmailnachricht\b)/i;

      // A) Anlagen-Chips: bevorzugt innerhalb eines Anlagen-Kontainers
      //    (aria-label "Anlagen"/"Attachments"), sonst Heuristik ueber
      //    klickbare Elemente mit dateinamenaehnlichem Label.
      const containers = Array.from(pane.querySelectorAll('[aria-label*="nlagen" i], [aria-label*="nlage" i], [aria-label*="ttachment" i]'));
      const chipRoots = [];
      const seenEls = new Set();
      const labelOf = (el) => [el.getAttribute && el.getAttribute('aria-label'), el.getAttribute && el.getAttribute('title'), el.innerText].filter(Boolean).join(' ');
      const acceptChip = (el, strict) => {
        if (!el || seenEls.has(el)) return;
        const label = labelOf(el);
        if (!label.trim()) return;
        // Innerhalb eines Anlagen-Kontainers reicht ein nichtleeres Label — das
        // deckt .msg-"Outlook Item"-Karten ohne Dateiendung ab (gbrain #7471).
        // Im Fallback weiter draussen gelten Dateiendung ODER Typ-Stichwort.
        if (strict && !extRe.test(label) && !typeRe.test(label)) return;
        seenEls.add(el);
        chipRoots.push(el);
      };
      for (const c of containers) {
        c.querySelectorAll('[role="listitem"], [role="option"], [role="button"], button, a').forEach((el) => acceptChip(el, false));
      }
      if (chipRoots.length === 0) {
        pane.querySelectorAll('[role="listitem"], [role="button"], button, a, span[title], div[title]').forEach((el) => acceptChip(el, true));
      }

      const chips = [];
      const seenNames = new Set();
      for (const el of chipRoots) {
        const label = [el.getAttribute('aria-label'), el.getAttribute('title'), el.innerText].filter(Boolean).join(' ');
        const line = label.split(/\n|\s{2,}/).map(function (s) { return s.trim(); }).filter(Boolean).find(function (s) { return extRe.test(s) || typeRe.test(s); }) || label.trim();
        const name = line.replace(/[\s,]*(Outlook Item|E-Mail-Nachricht|Herunterladen|Download)\s*$/i, '').trim().slice(0, 180);
        if (!name || seenNames.has(name)) continue;
        seenNames.add(name);
        chips.push({
          name: name,
          tag: el.tagName,
          role: el.getAttribute('role') || null,
          hasDownloadIcon: !!el.querySelector('[data-icon-name="Download"], [aria-label*="erunterladen" i]')
        });
      }

      // B) Inline-Bilder im Body (blob:/data:) mit sinnvoller Groesse
      const inlineImages = [];
      for (const img of pane.querySelectorAll('img')) {
        const src = img.getAttribute('src') || '';
        if (src.indexOf('blob:') !== 0 && src.indexOf('data:image/') !== 0) continue;
        const r = img.getBoundingClientRect();
        if (r.width < 100 || r.height < 60) continue;
        inlineImages.push({
          srcType: src.indexOf('blob:') === 0 ? 'blob' : 'data',
          alt: (img.getAttribute('alt') || '').slice(0, 80) || null,
          width: Math.round(r.width),
          height: Math.round(r.height)
        });
      }
      return { containerCount: containers.length, chips: chips, inlineImages: inlineImages };
    });
  }

  // Oeffentlicher Read-Pfad mit Restart-Retry (P6): Hover-Downloads sind
  // fragil gegenueber Mid-Navigation-Rerenders.
  async listAttachments(opts = {}) {
    return this._withReadRetry('outlook_list_attachments', () => this._listAttachmentsInner(opts));
  }

  async _listAttachmentsInner({ index = 0, id, folder, timeoutMs = 60000 } = {}) {
    const page = await this._openEmailForAttachments({ index, id, folder, timeoutMs });
    const scan = await this._scanAttachments(page);
    if (scan.error) throw new Error(scan.error);
    const total = (scan.chips ? scan.chips.length : 0) + (scan.inlineImages ? scan.inlineImages.length : 0);
    return {
      target: id || ('index:' + index),
      chips: scan.chips,
      inlineImages: scan.inlineImages,
      containerCount: scan.containerCount,
      hint: total === 0
        ? 'Keine Anlagen/Inline-Bilder im Lesebereich gefunden — Mail ggf. zuerst mit outlook_read_email gegenpruefen.'
        : 'Speichern mit outlook_save_attachments (gleiche Zielangabe index/id/folder).'
    };
  }

  async saveAttachments({ index = 0, id, folder, destDir, includeInlineImages = true, timeoutMs = 60000 } = {}) {
    const page = await this._openEmailForAttachments({ index, id, folder, timeoutMs });
    const scan = await this._scanAttachments(page);
    if (scan.error) throw new Error(scan.error);

    // Zielordner: Default ~/projects/outlook-attachments/<Zeitstempel>;
    // Sicherheits-Guard: NUR innerhalb des Home-Verzeichnisses.
    const home = os.homedir();
    let targetDir;
    if (destDir && String(destDir).trim()) {
      targetDir = path.resolve(String(destDir).replace(/^~(?=\/|$)/, home));
    } else {
      const stamp = new Date().toISOString().replace(/[:T]/g, '-').slice(0, 19);
      targetDir = path.join(home, 'projects', 'outlook-attachments', stamp);
    }
    if (!(targetDir === home || targetDir.startsWith(home + path.sep))) {
      throw new Error('destDir muss innerhalb des Home-Verzeichnisses liegen (abgelehnt: ' + targetDir + ')');
    }
    fs.mkdirSync(targetDir, { recursive: true });

    const saved = [];
    const failed = [];
    const usedNames = new Set();
    const uniqueTarget = (name) => {
      const safe = sanitizeFilename(name);
      const ext = path.extname(safe);
      const base = ext ? safe.slice(0, safe.length - ext.length) : safe;
      let candidate = safe;
      let n = 1;
      while (usedNames.has(candidate.toLowerCase())) {
        candidate = base + '-' + n + ext;
        n += 1;
      }
      usedNames.add(candidate.toLowerCase());
      return path.join(targetDir, candidate);
    };

    // --- Weg A: Chips ueber Download-Trigger + Playwright-Download-Event ---
    // Handles einmalig im Browser sammeln (gleiche Logik wie _scanAttachments,
    // aber als ElementHandles, damit ohne Re-Query geklickt werden kann).
    const chipArray = await page.evaluateHandle(() => {
      const pane = document.querySelector('div[role="main"], div[aria-label="Reading Pane"], div[aria-label="Lesebereich"], div.ItemPartView');
      if (!pane) return [];
      const extRe = /\.(png|jpe?g|gif|bmp|webp|heic|tif|tiff|pdf|docx?|xlsx?|pptx?|one|zip|7z|rar|msg|eml|txt|csv|ics|vcf|html?)($|\s|\()/i;
      const typeRe = /(\boutlook\s*item\b|\be-?mail-?nachricht\b|\bnachrichtenelement\b|\bmailnachricht\b)/i;
      const containers = Array.from(pane.querySelectorAll('[aria-label*="nlagen" i], [aria-label*="nlage" i], [aria-label*="ttachment" i]'));
      const chipRoots = [];
      const seenEls = new Set();
      const labelOf = (el) => [el.getAttribute && el.getAttribute('aria-label'), el.getAttribute && el.getAttribute('title'), el.innerText].filter(Boolean).join(' ');
      const acceptChip = (el, strict) => {
        if (!el || seenEls.has(el)) return;
        const label = labelOf(el);
        if (!label.trim()) return;
        // Innerhalb eines Anlagen-Kontainers reicht ein nichtleeres Label — das
        // deckt .msg-"Outlook Item"-Karten ohne Dateiendung ab (gbrain #7471).
        // Im Fallback weiter draussen gelten Dateiendung ODER Typ-Stichwort.
        if (strict && !extRe.test(label) && !typeRe.test(label)) return;
        seenEls.add(el);
        chipRoots.push(el);
      };
      for (const c of containers) {
        c.querySelectorAll('[role="listitem"], [role="option"], [role="button"], button, a').forEach((el) => acceptChip(el, false));
      }
      if (chipRoots.length === 0) {
        pane.querySelectorAll('[role="listitem"], [role="button"], button, a, span[title], div[title]').forEach((el) => acceptChip(el, true));
      }
      return chipRoots;
    });
    const chipProps = await chipArray.getProperties();
    await chipArray.dispose();

    // Kaskade: exaktes Fluent-Download-Icon -> Label beginnt mit
    // "Herunterladen"/"Download" -> Substring-Match. ("Alle herunterladen"
    // als ZIP-Fallback bewusst NICHT Teil der Kaskade — pro-Anhang-Dateien
    // sind das Ziel.)
    const TRIGGER_SELS = [
      '[data-icon-name="Download"]',
      'button[aria-label^="Herunterladen" i], button[aria-label^="Download" i], [role="button"][aria-label^="Herunterladen" i], [role="button"][aria-label^="Download" i]',
      'button[aria-label*="erunterladen" i], button[aria-label*="ownload" i]'
    ];

    for (const [, chipHandle] of chipProps) {
      let name = 'anhang';
      try {
        name = await chipHandle.evaluate((el) => {
          const extRe = /\.(png|jpe?g|gif|bmp|webp|heic|tif|tiff|pdf|docx?|xlsx?|pptx?|one|zip|7z|rar|msg|eml|txt|csv|ics|vcf|html?)($|\s|\()/i;
      const typeRe = /(\boutlook\s*item\b|\be-?mail-?nachricht\b|\bnachrichtenelement\b|\bmailnachricht\b)/i;
          const label = [el.getAttribute('aria-label'), el.getAttribute('title'), el.innerText].filter(Boolean).join(' ');
          const line = label.split(/\n|\s{2,}/).map(function (s) { return s.trim(); }).filter(Boolean).find(function (s) { return extRe.test(s) || typeRe.test(s); }) || label.trim();
          return line.replace(/[\s,]*(Outlook Item|E-Mail-Nachricht|Herunterladen|Download)\s*$/i, '').trim() || 'anhang';
        });
      } catch (e) {
        // Chip evtl. zwischen Scan und Klick rerendert -> generischer Name
      }

      try {
        await chipHandle.scrollIntoViewIfNeeded({ timeout: 5000 });
      } catch (e) {
        // nicht kritisch
      }

      let trigger = null;
      let previewOpened = false;
      for (const sel of TRIGGER_SELS) {
        trigger = await chipHandle.$(sel).catch(() => null);
        if (trigger) break;
      }
      if (!trigger) {
        // Download-Button erscheint in OWA teils erst beim Hover
        await chipHandle.hover({ timeout: 4000 }).catch(() => {});
        for (const sel of TRIGGER_SELS) {
          trigger = await chipHandle.$(sel).catch(() => null);
          if (trigger) break;
        }
      }
      if (!trigger) {
        // Chip-Klick oeffnet die Vorschau; Download-Button dort suchen
        await chipHandle.click({ timeout: 5000 }).catch(() => {});
        await page.waitForTimeout(1800);
        previewOpened = true;
        for (const sel of TRIGGER_SELS) {
          trigger = await page.$('[role="dialog"] ' + sel).catch(() => null);
          if (trigger) break;
          trigger = await page.$(sel).catch(() => null);
          if (trigger) break;
        }
      }

      if (!trigger) {
        failed.push({ name: name, error: 'Kein Download-Trigger gefunden (Chip + Vorschau)' });
        await page.keyboard.press('Escape').catch(() => {});
        continue;
      }

      try {
        const dlPromise = page.waitForEvent('download', { timeout: 30000 });
        await trigger.click({ timeout: 8000 });
        const download = await dlPromise;
        const suggested = download.suggestedFilename() || name;
        const target = uniqueTarget(suggested);
        await download.saveAs(target);
        saved.push({
          file: target,
          sizeBytes: fs.existsSync(target) ? fs.statSync(target).size : null,
          source: previewOpened ? 'chip-vorschau-download' : 'chip-download'
        });
        // Vorschau/Overlay schliessen, damit der naechste Chip klickbar bleibt
        await page.keyboard.press('Escape').catch(() => {});
        await page.waitForTimeout(600);
      } catch (e) {
        failed.push({ name: name, error: String((e && e.message) || e).slice(0, 200) });
        await page.keyboard.press('Escape').catch(() => {});
      }
    }
    for (const [, h] of chipProps) {
      h.dispose().catch(() => {});
    }

    // --- Weg B (Fallback): Inline-Bilder aus dem Body extrahieren ---
    const inlineCount = scan.inlineImages ? scan.inlineImages.length : 0;
    let inlineSkippedReason = null;
    if (includeInlineImages && inlineCount > 0 && saved.length === 0) {
      const imgs = await page.evaluate(async () => {
        const pane = document.querySelector('div[role="main"], div[aria-label="Reading Pane"], div[aria-label="Lesebereich"], div.ItemPartView');
        if (!pane) return [];
        const out = [];
        for (const img of pane.querySelectorAll('img')) {
          const src = img.getAttribute('src') || '';
          if (src.indexOf('blob:') !== 0 && src.indexOf('data:image/') !== 0) continue;
          const r = img.getBoundingClientRect();
          if (r.width < 100 || r.height < 60) continue;
          try {
            let b64 = null;
            let mime = 'image/png';
            if (src.indexOf('data:') === 0) {
              const m = src.match(/^data:([^;]+);base64,([\s\S]*)$/);
              if (m) { mime = m[1]; b64 = m[2]; }
            } else {
              const resp = await fetch(src);
              const blob = await resp.blob();
              mime = blob.type || 'image/png';
              const buf = new Uint8Array(await blob.arrayBuffer());
              let bin = '';
              for (let i = 0; i < buf.length; i += 0x8000) bin += String.fromCharCode.apply(null, buf.subarray(i, i + 0x8000));
              b64 = btoa(bin);
            }
            if (b64) out.push({ b64: b64, mime: mime, alt: img.getAttribute('alt') || null });
          } catch (e) {
            out.push({ error: String(e).slice(0, 120) });
          }
        }
        return out;
      });
      let idx = 0;
      for (const im of imgs) {
        idx += 1;
        if (!im.b64) {
          failed.push({ name: 'inline-bild-' + idx, error: im.error || 'kein base64' });
          continue;
        }
        const ext = (String(im.mime || 'image/png').split('/')[1] || 'png').replace('jpeg', 'jpg').replace(/[^a-z0-9]/gi, '') || 'png';
        const baseName = (im.alt && /\.(png|jpe?g|gif|webp)$/i.test(im.alt)) ? im.alt : 'inline-bild-' + idx + '.' + ext;
        const target = uniqueTarget(baseName);
        fs.writeFileSync(target, Buffer.from(im.b64, 'base64'));
        saved.push({ file: target, sizeBytes: fs.statSync(target).size, source: 'inline-image' });
      }
    } else if (inlineCount > 0 && saved.length > 0) {
      inlineSkippedReason = 'Chips erfolgreich gespeichert — Inline-Bilder uebersprungen (Duplikatsvermeidung)';
    }

    return {
      success: saved.length > 0,
      target: id || ('index:' + index),
      destDir: targetDir,
      savedCount: saved.length,
      saved: saved,
      failed: failed,
      inlineSkippedReason: inlineSkippedReason,
      scanned: { chips: scan.chips, inlineImageCount: inlineCount }
    };
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

    // DISCARD-STRATEGIEN (Befund 15.10.2026): Der frühere Einmal-Pfad (Kontextmenü
    // "Verwerfen", sonst Toolbar-Button) löschte Reply-Entwürfe nicht zuverlässig —
    // der Entwurf blieb in der Liste, weitere Versuche aktualisierten nur den
    // Zeitstempel (OWA-Autosave beim erneuten Öffnen). Bewährte Wege aus der
    // Live-Session (delete_junk_keyboard.mjs / discard_junk_drafts.mjs):
    //   1) Zeile selektieren + ENTF-Taste (Delete-Key) — zuverlässigster Weg
    //   2) Toolbar-Button "Löschen/Discard" im geöffneten Entwurf
    //   3) Kontextmenü (Rechtsklick) -> "Löschen/Verwerfen"
    // Nach JEDEM Versuch: Bestätigungsdialog abfangen, Undo-Toast auslaufen lassen,
    // Liste NEU LADEN (navigateFolder) und erst dann verifizieren — der DOM-Stand
    // vor dem Reload ist kein Beleg (OWA rendert die Zeile teils nach).
    const confirmDialog = async () => this._confirmDestructiveDialog(page);

    const strategies = [
      { name: 'delete-key', run: async () => {
          await targetRow.click();
          await page.waitForTimeout(1500);
          await page.keyboard.press('Delete');
        } },
      { name: 'toolbar-button', run: async () => {
          await targetRow.click();
          await page.waitForTimeout(2000);
          const toolbarBtn = await page.waitForSelector(
            'button[aria-label*="Löschen" i], button[aria-label*="Delete" i], ' +
            'button[aria-label*="Verwerfen" i], button[aria-label*="Discard" i], ' +
            'button[title*="Löschen" i], button[title*="Delete" i]',
            { timeout: 6000 }
          ).catch(() => null);
          if (!toolbarBtn) return false;
          await toolbarBtn.click();
          return true;
        } },
      { name: 'context-menu', run: async () => {
          await targetRow.click({ button: 'right' });
          await page.waitForTimeout(600);
          const discardItem = await page.waitForSelector(
            '[role="menuitem"]:has-text("Löschen"), [role="menuitem"]:has-text("Verwerfen"), ' +
            '[role="menuitem"]:has-text("Delete"), [role="menuitem"]:has-text("Discard"), ' +
            '[role="menuitem"][aria-label*="Löschen" i], [role="menuitem"][aria-label*="Verwerfen" i], ' +
            '[role="menuitem"][aria-label*="Delete" i], [role="menuitem"][aria-label*="Discard" i]',
            { timeout: 5000 }
          ).catch(() => null);
          if (!discardItem) {
            await page.keyboard.press('Escape').catch(() => null);
            return false;
          }
          await discardItem.click();
          return true;
        } }
    ];

    const attempted = [];
    let gone = false;
    for (const strategy of strategies) {
      let triggered = false;
      try {
        triggered = (await strategy.run()) !== false;
      } catch (e) {
        attempted.push(`${strategy.name}: FEHLER ${e && e.message ? e.message : e}`);
        continue;
      }
      if (!triggered) {
        attempted.push(`${strategy.name}: Auslöser nicht gefunden`);
        continue;
      }
      attempted.push(`${strategy.name}: ausgelöst`);

      // Bestätigungsdialoge abfangen (max. 3, niemals Abbrechen klicken)
      for (let i = 0; i < 3; i++) {
        await page.waitForTimeout(800);
        if (!(await confirmDialog())) break;
        await page.waitForTimeout(1200);
      }

      // Undo-Toast auslaufen lassen (NICHT anklicken — der Verwurf soll stehen bleiben)
      await page.waitForTimeout(1500);

      // VERIFIKATION NACH LISTEN-RELOAD: Der Zustand vor dem Reload ist kein Beleg.
      // Die frische Zeilen-Referenz dient zugleich der nächsten Strategie —
      // alte Handles sind nach dem Reload detached.
      await this.navigateFolder(folder || 'drafts');
      await page.waitForTimeout(2000);
      await page.waitForSelector('div[data-convid]', { timeout: 10000 }).catch(() => null);
      targetRow = null;
      const rowsAfter = await page.$$('div[data-convid]');
      for (const r of rowsAfter) {
        if (await rowMatches(r)) { targetRow = r; break; }
      }
      if (!targetRow) { gone = true; break; }
      attempted.push(`${strategy.name}: Entwurf nach Reload noch sichtbar`);
    }

    return {
      success: gone,
      verified: gone,
      query: targetQuery,
      folder: folder || 'drafts',
      attemptedStrategies: attempted,
      message: gone
        ? `Entwurf "${targetQuery}" wurde verworfen (nach Listen-Reload verifiziert, aus der Entwurfsliste entfernt).`
        : `Fehler: Entwurf "${targetQuery}" konnte NICHT verworfen werden — nach allen Strategien noch in der Entwurfsliste sichtbar. Versuche: ${attempted.join(' | ')}. Bitte manuell in Outlook prüfen (ggf. OWA-Autosave hat den Entwurf neu angelegt).`
    };
  }

  // Bestätigungsdialog abfangen (falls OWA nachfragt): Verwerfen/Löschen bestätigen,
  // niemals "Abbrechen/Cancel" klicken. Rueckgabe: true, wenn bestätigt wurde.
  async _confirmDestructiveDialog(page) {
    return page.evaluate(() => {
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
                  comb.includes('löschen') || comb.includes('delete') ||
                  comb === 'ok') &&
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
  }

  // Gemeinsame Lokalisierung einer Termineinladung (RSVP Accept/Decline):
  // Suchbegriff -> OWA-Suche; convid/Zeilen-ID -> Ordner-Scan; sonst erste Zeile.
  // Klickt die Einladung an und laedt den Lesebereich (Rueckgabe: Lesebereich-Handle).
  async _locateMeetingInvite({ targetId, searchQuery, folder } = {}) {
    const page = await this.getPage(true);

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

    return { page, readingPane };
  }

  async acceptMeeting({ email_id, id, query, subject, folder, send_response = true } = {}) {
    const targetId = email_id || id;
    const searchQuery = query || subject;

    // 1. Einladung lokalisieren + Lesebereich oeffnen (gemeinsamer RSVP-Helper)
    const { page, readingPane } = await this._locateMeetingInvite({ targetId, searchQuery, folder });
    const context = browserManager.getContext(this.realm);

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

  async declineMeeting({ email_id, id, query, subject, folder, send_response = true, comment } = {}) {
    const targetId = email_id || id;
    const searchQuery = query || subject;
    const hasComment = typeof comment === 'string' && comment.trim().length > 0;

    // 1. Einladung lokalisieren + Lesebereich oeffnen (gemeinsamer RSVP-Helper)
    const { page, readingPane } = await this._locateMeetingInvite({ targetId, searchQuery, folder });
    const context = browserManager.getContext(this.realm);

    // 3. "Absagen" (Decline) Button finden und klicken.
    // OWA zeigt bei Termineinladungen im Lesebereich eine Toolbar mit "Annehmen / Unter Vorbehalt / Absagen"
    // (deutsch) bzw. "Accept / Tentative / Decline" (englisch), oft als Dropdown unter "Jetzt antworten".
    const findDecline = async () => {
      // Präziser Selektor: OWA-Toolbar der Termineinladung nutzt aria-label "Decline the meeting" / "Absagen".
      const direct = await page.$('button[aria-label="Decline the meeting"], button[aria-label*="Decline the meeting"], button[aria-label="Absagen"], button[aria-label*="Absagen"]').catch(() => null);
      if (direct) return direct;
      const buttons = await page.$$('button[aria-label], button[title], [role="button"][aria-label]');
      for (const btn of buttons) {
        const ariaLabel = ((await btn.getAttribute('aria-label')) || '').toLowerCase();
        const title = ((await btn.getAttribute('title')) || '').toLowerCase();
        const text = ((await btn.innerText().catch(() => '')) || '').toLowerCase();
        const combined = `${ariaLabel} ${title} ${text}`;
        // "Absagen/Decline" ist die Ziel-Aktion; kombinierte Gruppen-Labels (Accept/Decline,
        // Tentative) sowie Annehmen-Varianten ausschließen.
        if ((combined.includes('absagen') || combined.includes('decline')) &&
            !combined.includes('annehmen') && !combined.includes('accept') &&
            !combined.includes('tentative') && !combined.includes('unter vorbehalt') &&
            !combined.includes('cancel') && !combined.includes('stornier')) {
          return btn;
        }
      }
      return null;
    };

    // 3a. Fallback: "Jetzt antworten" / "Respond" Dropdown öffnen und "Absagen" darin wählen
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

    let targetBtn = await findDecline();
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
          const isDecline =
            combined.includes('absagen') || combined.includes('decline') || combined.includes('ablehnen');
          const isNotAccept =
            !combined.includes('annehmen') && !combined.includes('accept') &&
            !combined.includes('unter vorbehalt') && !combined.includes('tentative');
          if (isDecline && isNotAccept) {
            targetBtn = item;
            break;
          }
        }
      }
    }

    if (!targetBtn) {
      throw new Error('"Absagen"-Button für die Termineinladung in Outlook Web nicht gefunden.');
    }

    // 4. Decline klicken (aktiver Tab/Page berücksichtigen)
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

    // 5. Bestätigungsdialog (falls vorhanden) behandeln:
    //    - mit Antworttext: "Antwort bearbeiten" -> Compose-Editor befüllen -> senden
    //    - sonst send_response=true:  "Antwort senden" / "Send the response"
    //    - sonst send_response=false: "Keine Antwort senden" / "Do not send a response"
    if (hasComment) {
      const editBtn = await activePage.$('button:has-text("Antwort bearbeiten"), button:has-text("Edit the response")').catch(() => null);
      if (editBtn) {
        await editBtn.click();
        await activePage.waitForTimeout(1500);
      }

      // Compose-Editor der Ablehnungs-Antwort finden (echter contenteditable-Knoten)
      const editor = await activePage.waitForSelector(
        'div[role="textbox"][contenteditable="true"][aria-label*="Message body" i], ' +
        'div[role="textbox"][contenteditable="true"][aria-label*="Nachrichtentext" i], ' +
        '[contenteditable="true"][aria-label*="Message body" i], ' +
        '[contenteditable="true"][aria-label*="Nachrichtentext" i]',
        { timeout: 15000 }
      ).catch(() => null);

      if (!editor) {
        throw new Error('Compose-Editor für den Ablehnungs-Kommentar konnte nicht gefunden werden.');
      }
      await editor.click();
      await activePage.waitForTimeout(300);
      await this._injectBodyText(activePage, editor, comment.trim());

      // Antwort absenden: Senden-Button oder Strg+Enter
      let sent = false;
      const sendBtn = await activePage.$('button[aria-label="Senden"], button[aria-label="Send"], button:has-text("Senden"), button:has-text("Send")').catch(() => null);
      if (sendBtn) { await sendBtn.click(); sent = true; }
      if (!sent) {
        await activePage.keyboard.press('Control+Enter');
        sent = true;
      }
      await activePage.waitForTimeout(2000);
    } else {
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
        if (send_response) {
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
        } else {
          // Ohne Antwortübermittlung: "Keine Antwort senden" / "Do not send a response"
          const noSend = await activePage.$('button:has-text("Keine Antwort senden"), button:has-text("Do not send a response")').catch(() => null);
          if (noSend) { await noSend.click(); await activePage.waitForTimeout(1200); }
        }
      }
    }

    await activePage.waitForTimeout(1500);

    // 6. Ergebnis prüfen: Status der Einladung
    // Nach dem Decline verschwindet die Einladung i.d.R. aus dem Posteingang (Outlook-Flow).
    // Daher gilt die Absage auch dann als erfolgreich, wenn die convid nicht mehr in der
    // aktuell sichtbaren Nachrichtenliste erscheint.
    const statusText = await readingPane.innerText().catch(() => '');
    const statusDeclined = statusText.toLowerCase().includes('abgesagt') ||
                           statusText.toLowerCase().includes('declined') ||
                           statusText.toLowerCase().includes('you declined') ||
                           statusText.toLowerCase().includes('nicht teil');

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

    const declined = statusDeclined || disappearedFromList;

    return {
      success: true,
      declined: declined,
      notification: declined ? undefined : 'Decline ausgeführt; Status konnte nicht eindeutig als "abgesagt" verifiziert werden.',
      message: declined && statusDeclined
        ? 'Termineinladung erfolgreich abgesagt. Die Gegenseite wurde benachrichtigt, der Termin wird nicht eingetragen.'
        : declined
          ? 'Termineinladung erfolgreich abgesagt und aus dem Posteingang entfernt. Die Gegenseite wurde benachrichtigt.'
          : 'Termineinladung wurde abgesagt. Bitte prüfe den Status in Outlook.'
    };
  }

  async replyDraft({ email_id, id, query, subject, body, reply_all = false, replyAll, folder, expected_from, expectedTo, expected_to, expectedSubject, expected_subject, expectedDate, expected_date } = {}) {
    if (!body || typeof body !== 'string' || body.trim().length === 0) {
      throw new Error('Parameter "body" ist erforderlich und darf nicht leer sein.');
    }

    const targetId = email_id || id;
    const searchQuery = query || subject;
    const isReplyAll = Boolean(reply_all ?? replyAll ?? false);
    // Alias-Normalisierung der Assertions-Parameter (camelCase und snake_case)
    const expFrom = expected_from;
    const expTo = expected_to || expectedTo;
    const expSubject = expected_subject || expectedSubject;
    const expDate = expected_date || expectedDate;

    const page = await this.getPage(true);
    const context = browserManager.getContext(this.realm);

    // Vorschau der getroffenen Ziel-Mail (Transparenz: welche Thread-Nachricht wurde beantwortet?)
    let selectedMailPreview = null;
    // Volltext der gewaehlten Listenzeile — Beweisbasis der Reply-Ziel-Assertion
    // (Absender + Betreff + Datum stehen in der Zeile, Befund 15.10.2026)
    let selectedRowText = '';

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
      selectedRowText = ((await selectedRow.innerText().catch(() => '')) || '').slice(0, 800);
      selectedMailPreview = selectedRowText
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
      selectedRowText = ((await targetRow.innerText().catch(() => '')) || '').slice(0, 800);
      selectedMailPreview = selectedRowText
        .split('\n').map(l => l.trim()).filter(Boolean).slice(0, 3).join(' | ').slice(0, 200);
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

    let replyTargetAssertion = null;
    // REPLY-ZIEL-ASSERTION (Sicherheitsgate, Befund 15.10.2026): Bevor geantwortet
    // wird, muss die gewählte Quellmail zu den erwarteten Werten passen — ein
    // Fehl-Reply (Prignitz/Isenberg statt UHK) erzeugt sonst einen Draft im
    // falschen Thread mit falschem Text. Beweisbasis: Listenzeile + Lesebereich-
    // Header (Absender-Chips, "An:"-Zeile, Betreff-Heading).
    if (expFrom || expTo || expSubject || expDate) {
      const headerEvidence = await page.evaluate(() => {
        const pane = document.querySelector('div[role="main"], div[aria-label="Lesebereich"], div[aria-label="Reading Pane"], div.ItemPartView');
        if (!pane) return { senderHints: '', recipientHints: '', paneHead: '' };
        // Strings aus Title-/Aria-Attributen: OWA legt E-Mail-Adressen in
        // Persona-Chips dort ab, nicht immer im sichtbaren Text.
        const attribs = [];
        for (const el of pane.querySelectorAll('[title], [aria-label]')) {
          attribs.push(el.getAttribute('title') || '', el.getAttribute('aria-label') || '');
        }
        const withAddress = attribs.filter((s) => s && s.includes('@')).join(' ; ');
        const paneText = (pane.innerText || '').slice(0, 1500);
        return {
          senderHints: `${withAddress} ; ${paneText}`,
          recipientHints: withAddress,
          paneHead: paneText
        };
      }).catch(() => ({ senderHints: '', recipientHints: '', paneHead: '' }));

      const assertion = resolveReplyTargetAssertion({
        evidence: `${selectedRowText}\n${headerEvidence.paneHead}`,
        senderEvidence: `${selectedRowText}\n${headerEvidence.senderHints}`,
        recipientEvidence: headerEvidence.recipientHints,
        expectedFrom: expFrom,
        expectedTo: expTo,
        expectedSubject: expSubject,
        expectedDate: expDate,
        rowTimestamp: selectedRowText ? parseMailRowTimestamp(selectedRowText) : null
      });

      if (!assertion.ok) {
        throw new Error(
          `REPLY-ZIEL-PRÜFUNG FEHLGESCHLAGEN (${assertion.reason}): Die gewählte Quellmail passt ` +
          `nicht zu den erwarteten Werten. Es wurde KEIN Antwort-Entwurf erstellt. ` +
          `Gefundene Quellmail: "${selectedMailPreview || headerEvidence.paneHead.slice(0, 200) || '(kein Zeilentext verfügbar)'}". ` +
          `Prüfung: ${JSON.stringify(assertion.details)}. Bitte Zielpost/expected_*-Parameter prüfen ` +
          `oder die Quellmail per email_id eindeutig adressieren.`
        );
      }
      replyTargetAssertion = {
        checked: assertion.details,
        unverified: assertion.unverified
      };
    }

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
    //
    // LEBT-BEFUND 09.10.2026: Es gibt im Compose-Fenner MEHRERE Kandidaten mit
    // aria-label "Message body": einen echten Editor
    // (role="textbox" + contenteditable="true") und einen Wrapper ohne
    // contenteditable. Ein selektor, der zuerst `div[aria-label="Message body"]`
    // trifft, landet auf dem Wrapper - dort ist innerText CSS-Text statt
    // Nachrichtentext, und jede Verifikation ist wertlos.
    //
    // Deshalb wird der Editor als ECHTE EIGENSCHAFT bestimmt: role="textbox"
    // ODER contenteditable="true", jeweils mit Message-body-/Nachrichtentext-
    // Label. Der Wrapper faellt dadurch durch.
    const editor = await activePage.waitForSelector(
      'div[role="textbox"][contenteditable="true"][aria-label*="Message body" i], ' +
      'div[role="textbox"][contenteditable="true"][aria-label*="Nachrichtentext" i], ' +
      '[contenteditable="true"][aria-label*="Message body" i], ' +
      '[contenteditable="true"][aria-label*="Nachrichtentext" i], ' +
      'div[role="textbox"][aria-label*="Message body" i], ' +
      'div[role="textbox"][aria-label*="Nachrichtentext" i], ' +
      'div[aria-label="Message body"][contenteditable="true"], ' +
      'div[aria-label="Nachrichtentext"][contenteditable="true"]',
      { timeout: 15000 }
    ).catch(() => null);

    if (!editor) {
      throw new Error('Compose-Editor für den Antwort-Entwurf konnte nicht gefunden werden.');
    }

    // EIGENSCHAFTS-PRUEFUNG: Der Editor MUSS ein echter contenteditable-Knoten
    // sein. Ein Wrapper wuerde jede Textpruefung verfaelschen - deshalb wird
    // hier hart abgebrochen statt einen unbrauchbaren Knoten zu verwenden.
    const editorKind = await editor.evaluate((el) => ({
      contentEditable: el.getAttribute && el.getAttribute('contenteditable'),
      role: (el.getAttribute && el.getAttribute('role')) || '',
      tag: el.tagName
    }));
    if (!isRealComposeEditor(editorKind)) {
      throw new Error(
        `Gefundener Compose-Knoten ist kein echter Editor ` +
        `(contenteditable="${editorKind.contentEditable}", role="${editorKind.role}"). ` +
        `Eine Textpruefung waere wertlos - es wird nichts getippt und nichts gespeichert.`
      );
    }

    // 6. TEXTUEBERNAHME (Problem 1, Playwright-Serialisierung) — LEBT-BEFUND
    // 09.10.2026: Die frueher hier eingesetzte reine DOM-Mutation
    // (insertBefore + dispatchEvent('input')) fuehrt bei OWA NICHT zur
    // Serialisierung. Live-Beweis: Draft wurde gespeichert, der gespeicherte
    // Koerper enthielt nur den zitierten Thread, der eingefuegte Text fehlte
    // vollstaendig. Grund: OWA serialisiert aus dem internen Editor-Zustand
    // (eigenes contenteditable-Modell), nicht aus dem DOM. Ein
    // DOM-insertBefore ist fuer dieses Modell unsichtbar — der anschliessende
    // innerText-Check bestaetigte deshalb nur die eigene Mutation und meldete
    // faelschlich Erfolg ("dom-insertion").
    //
    // KORREKTER WEG: Der Text wird autokorrektur-frei eingefügt (CDP
    // Input.insertText oder Clipboard-Paste — s. _injectBodyText). Beide Wege
    // erzeugen Input-Events, die das OWA-Editor-Modell verarbeitet, und sind
    // damit serialisierbar — OHNE den Autokorrektur-Pfad des zeichenweisen
    // Tippens (Befund 15.10.2026: "dem"->"Dem" nach "bzw. "). Die Verifikation
    // liest den Editor NACH dem Einfügen; der entscheidende Nachweis ist aber
    // das Wiederoeffnen des gespeicherten Entwurfs (Schritt 8) — erst der
    // persistierte Zustand gilt als Beleg, nicht der flüchtige Editor.
    //
    // REIHENFOLGE (bewusst, verhaltnis-richtig):
    //   a) Inhalt VOR dem Einfügen merken (Beweis fuer die Unversehrtheit von
    //      Verlauf/Signatur),
    //   b) Cursor an den Anfang des Editors setzen (additiv, KEIN Control+A —
    //      das wuerde Verlauf und Signatur markieren und beim Einfügen
    //      ueberschreiben, also genau das zerstoeren, was ein Reply bewahrt),
    //   c) Text einfügen (insertText, sonst Clipboard),
    //   d) Verifikation: (i) Text im Editor, (ii) Bestand noch vorhanden.
    const preInsertText = await editor.evaluate((el) => (el.innerText || el.textContent || '').trim());

    const insertionVerification = await editor.evaluate((el) => ({
      contentEditable: el.getAttribute && el.getAttribute('contenteditable'),
      role: el.getAttribute && el.getAttribute('role'),
      targetTag: el.tagName,
      preLength: (el.innerText || el.textContent || '').length
    }));

    // Caret an den Anfang des Editor-Inhalts setzen (rein additiv, keine
    // Selektion, kein Loeschen des Bestands).
    const caretPlaced = await editor.evaluate((el) => {
      try {
        const target = el.firstChild || el;
        const range = document.createRange();
        range.setStart(target, 0);
        range.collapse(true);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        el.focus();
        return true;
      } catch (e) {
        return false;
      }
    });
    if (!caretPlaced) {
      throw new Error(
        `Antwort-Text konnte nicht in den Editor übernommen werden: der Cursor ` +
        `konnte nicht gesetzt werden. Es wurde nichts gespeichert und nichts überschrieben.`
      );
    }

    // TEXTUEBERNAHME autokorrektur-frei (Befund 15.10.2026): insertText/Clipboard
    // statt zeichenweisem keyboard.type — OWA kapitalisierte sonst Wörter nach
    // "bzw. " ("dem"->"Dem") und Absatzanfänge ("kurzes"->"Kurzes").
    const injection = await this._injectBodyText(activePage, editor, body);

    // Verifikation NACH dem Einfügen (flüchtig — der starke Beleg folgt in Schritt 8
    // durch Wiederoeffnen des gespeicherten Entwurfs).
    const editorText = injection.editorText;
    if (!injection.method) {
      throw new Error(
        `Antwort-Text konnte NICHT in den Editor übernommen werden (insertText- und ` +
        `Clipboard-Weg fehlgeschlagen) — es wird nicht gespeichert, um keinen leeren ` +
        `oder verfaelschten Entwurf anzulegen. Diagnostik: ${JSON.stringify(insertionVerification)} ` +
        `Soll/Ist: ${JSON.stringify(buildBodySollIstDiff(body, editorText))}.`
      );
    }

    // Unversehrtheits-Beweis: Der Bestand (Verlauf/Signatur) muss nach dem Einfügen
    // noch vorhanden sein. Fehlt er, hat die Injection überschrieben — dann ist
    // das Entwurfs-Ergebnis verfaelscht und wird NICHT gespeichert.
    if (!isHistoryPreserved(preInsertText, editorText)) {
      throw new Error(
        `ABBRUCH: Beim Einfügen des Antwort-Texts ist der vorhandene E-Mail-Verlauf bzw. die ` +
        `Signatur verloren gegangen. Der Entwurf wurde NICHT gespeichert, damit kein ` +
        `Inhaltsverlust entsteht. Bitte den Entwurf manuell neu anlegen.`
      );
    }

    await activePage.waitForTimeout(500);

    // 7. Save draft
    await activePage.keyboard.press('Control+s');
    await activePage.waitForTimeout(3500);

    // bodyInsertSource kommt aus der Injection (cdp-insertText | clipboard-paste).
    const bodyInsertSource = injection.method;

    // 8. PERSISTIERTER NACHWEIS — gemeinsame Methode (s. verifyPersistedDraftBody).
    // Nur der Zustand NACH dem Speichern gilt als Beleg; der Entwurf bleibt bei
    // Fehlschlag zur Sichtung liegen und wird NICHT automatisch geloescht.
    const persisted = await this.verifyPersistedDraftBody(body);
    const persistedBodyVerified = persisted.verified;
    const persistedDetail = persisted.detail;

    // fuzzy = Abweichungen (z.B. Reste von Autokorrektur) — im Return als diff melden
    const bodyMatchMeta = persisted.match === 'fuzzy'
      ? { bodyMatch: 'fuzzy', bodySollIst: persisted.diff }
      : {};

    if (persistedBodyVerified !== true) {
      return {
        success: false,
        action: isReplyAll ? 'reply_all' : 'reply',
        subject: emailSubject || undefined,
        targetMailPreview: selectedMailPreview || undefined,
        replyTargetAssertion: replyTargetAssertion || undefined,
        preview: body.length > 200 ? body.substring(0, 200) + '...' : body,
        bodyInsertSource,
        persistedBodyVerified: persistedBodyVerified === null ? 'unverified' : false,
        folder: 'Entwürfe (Drafts)',
        message:
          `WARNUNG: Der Antwort-Entwurf wurde gespeichert, sein Text ist im gespeicherten ` +
          `Entwurf aber NICHT nachweisbar. Der Entwurf bleibt zur Sichtung in "Entwürfe" ` +
          `liegen und wurde NICHT automatisch gelöscht. Bitte manuell prüfen und ` +
          `ggf. neu anlegen. Grund: ${persistedDetail} | Soll/Ist: ` +
          `${JSON.stringify(persisted.diff || buildBodySollIstDiff(body, ''))}`,
        diagnostics: { insertionVerification, persistedDetail }
      };
    }

    return {
      success: true,
      action: isReplyAll ? 'reply_all' : 'reply',
      subject: emailSubject || undefined,
      targetMailPreview: selectedMailPreview || undefined,
      replyTargetAssertion: replyTargetAssertion || undefined,
      preview: body.length > 200 ? body.substring(0, 200) + '...' : body,
      bodyInsertSource,
      ...bodyMatchMeta,
      persistedBodyVerified: true,
      folder: 'Entwürfe (Drafts)',
      message:
        `Antwort-Entwurf (${isReplyAll ? 'Allen antworten' : 'Antworten'}) erstellt, in ` +
        `"Entwürfe" gespeichert und im gespeicherten Zustand verifiziert ` +
        `(Textübernahme: ${bodyInsertSource}${persisted.match === 'fuzzy' ? ', mit Abweichungen — siehe bodySollIst' : ''}). ` +
        `Verlauf und Signatur blieben erhalten.`
    };
  }

  // PERSISTIERTER NACHWEIS des gespeicherten Entwurfs (Problem 1, Live-Befund
  // 09.10.2026). Der fluechtige Editor-Inhalt ist KEIN Beweis fuer die
  // Speicherung — OWA serialisiert aus seinem internen Zustand, nicht aus dem
  // DOM. Belastbar ist nur der ZUSTAND NACH dem Speichern: den Entwurf erneut
  // oeffnen und den Koerper daraus lesen.
  //
  // Wird von replyDraft UND createDraft genutzt (gleicher Serialisierungs-Weg,
  // gleicher Beweisstandard). Der Entwurf wird hier NIE geloescht — ein
  // fehlgeschlagener Nachweis bleibt zur manuellen Sichtung liegen.
  //
  // Rueckgabe: { verified: boolean|null, detail: string }.
  // verified=null bedeutet "nicht pruefbar" (z.B. Draft nicht auffindbar).
  async verifyPersistedDraftBody(expectedBody, { withRecipients = false } = {}) {
    try {
      await this.navigateFolder('drafts');
      await new Promise((r) => setTimeout(r, 3000));
      const draftPage = await this.getPage(true);

      // Obersten [Draft] oeffnen (der juengste ist der soeben gespeicherte)
      const draftId = await draftPage.evaluate(() => {
        const rows = Array.from(document.querySelectorAll('div[data-convid]'));
        for (const r of rows) {
          const t = (r.innerText || '').replace(/\s+/g, ' ');
          if (/\[Draft\]/i.test(t)) {
            const id = r.getAttribute('data-convid') || '';
            r.click();
            return id;
          }
        }
        return null;
      });

      if (!draftId) {
        return { verified: null, detail: 'Draft nach dem Speichern nicht in der Entwurfsliste auffindbar.' };
      }

      await new Promise((r) => setTimeout(r, 3500));
      const persisted = await draftPage.evaluate(() => {
        // Gleiche Editor-Identitaet wie beim Compose: role=textbox ODER
        // contenteditable=true, damit nicht der Wrapper ohne echten Editor
        // gelesen wird.
        const sel = '[contenteditable="true"][aria-label*="Message body" i], ' +
          '[contenteditable="true"][aria-label*="Nachrichtentext" i], ' +
          'div[role="textbox"][aria-label*="Message body" i], ' +
          'div[role="textbox"][aria-label*="Nachrichtentext" i]';
        const ed = document.querySelector(sel);
        if (ed) return { found: true, text: (ed.innerText || ed.textContent || '').trim() };
        const pane = document.querySelector('div[role="main"]');
        return { found: false, text: pane ? (pane.innerText || '') : (document.body.innerText || '') };
      });

      // P4 (Befund 15.10.2026): Persistierte Empfänger mitlesen — der CC-DeepLink
      // (&cc=) wird von OWA teils nicht übernommen.
      let recipients = null;
      if (withRecipients) {
        recipients = await this._extractRecipientsFromCompose(draftPage);
      }

      // Dreistufiger Nachweis: 'exact' = sauber, 'fuzzy' = Abweichung (z.B.
      // Autokorrektur-Kapitalisierung) — zählt als verifiziert, MUSS aber im
      // Return als diff gemeldet werden. 'none' = nicht nachweisbar.
      const cls = classifyBodyMatch(persisted.text, expectedBody);
      return {
        verified: cls !== 'none',
        match: cls,
        diff: buildBodySollIstDiff(expectedBody, persisted.text),
        detail: `Persistierter Entwurfstext: ${persisted.text.length} Zeichen, Editor gefunden: ${persisted.found}, Match: ${cls}`,
        ...(recipients ? { recipients } : {})
      };
    } catch (e) {
      return { verified: null, detail: `Nachweis nicht durchfuehrbar: ${e && e.message ? e.message : e}` };
    }
  }

  // P4: Empfänger-Chips des aktuell offenen Compose/Drafts lesen. OWA legt die
  // E-Mail-Adressen in title-/aria-label-Attributen der Chips ab; die Zuordnung
  // zu "An"/"Cc" läuft über die aria-labels der übergeordneten Wells. Unklar
  // zuordenbare Chips zählen für beide Wells (konservativ — "CC fehlt" wird
  // dann nicht fälschlich gemeldet, wenn die Adresse nur schwach gelabelt ist).
  async _extractRecipientsFromCompose(page) {
    try {
      return await page.evaluate(() => {
        const out = { to: [], cc: [] };
        const push = (well, label) => {
          if (!label) return;
          if (well === 'cc') out.cc.push(label);
          else if (well === 'to') out.to.push(label);
          else { out.to.push(label); out.cc.push(label); }
        };
        const wellOf = (el) => {
          let cur = el;
          for (let i = 0; i < 8 && cur; i++) {
            const al = (cur.getAttribute && cur.getAttribute('aria-label')) || '';
            if (/^(an|to)\b/i.test(al.trim())) return 'to';
            if (/^(cc|ccempf|kopie|carbon)/i.test(al.trim())) return 'cc';
            cur = cur.parentElement;
          }
          return null;
        };
        // Chips: title/aria-label enthält eine E-Mail-Adresse
        for (const chip of document.querySelectorAll('[title*="@"], [aria-label*="@"]')) {
          const label = (chip.getAttribute('title') || chip.getAttribute('aria-label') || '').trim();
          push(wellOf(chip), label);
        }
        // Frisch getippte, noch nicht als Chip gerenderte Werte in Eingabefeldern
        for (const inp of document.querySelectorAll('input[aria-label]')) {
          const al = (inp.getAttribute('aria-label') || '').trim();
          const val = (inp.value || '').trim();
          if (!val) continue;
          if (/^(an|to)\b/i.test(al)) push('to', val);
          else if (/^(cc|kopie)/i.test(al)) push('cc', val);
        }
        return out;
      });
    } catch (e) {
      return { to: [], cc: [] };
    }
  }

  // P4: Persistierte Empfänger des obersten [Draft]-Entwurfs lesen (Entwurf
  // frisch öffnen — der Zustand nach dem Speichern ist der Beleg).
  async _readPersistedDraftRecipients() {
    try {
      const page = await this.getPage(true);
      await this.navigateFolder('drafts');
      await page.waitForTimeout(2500);
      await page.waitForSelector('div[data-convid]', { timeout: 10000 }).catch(() => null);
      const opened = await page.evaluate(() => {
        const rows = Array.from(document.querySelectorAll('div[data-convid]'));
        for (const r of rows) {
          const t = (r.innerText || '').replace(/\s+/g, ' ');
          if (/\[Draft\]/i.test(t)) {
            r.click();
            return true;
          }
        }
        return false;
      });
      if (!opened) {
        return { found: false, recipients: { to: [], cc: [] } };
      }
      await page.waitForTimeout(3000);
      const recipients = await this._extractRecipientsFromCompose(page);
      return { found: true, recipients };
    } catch (e) {
      return { found: false, recipients: { to: [], cc: [] } };
    }
  }

  // P4: Fehlende CC-Adresse im offenen Draft nachziehen (UI-Chips) und speichern.
  // Läuft auf der Seite, die verifyPersistedDraftBody mit dem geöffneten Entwurf
  // hinterlassen hat. Rückgabe: { repaired: boolean, reason?: string }.
  async _repairDraftCc(ccStr) {
    try {
      const page = await this.getPage(true);
      // Cc-Zeile ggf. erst einblenden (OWA: "Cc verwenden"-Button im eingeklappten Zustand)
      const ccReveal = await page.$(
        'button[aria-label*="Cc verwenden" i], button[aria-label*="Show Cc" i], ' +
        'button[aria-label*="Cc" i][aria-expanded="false"], button[title*="Cc" i]'
      );
      if (ccReveal) {
        await ccReveal.click().catch(() => null);
        await page.waitForTimeout(700);
      }
      const ccInput = await page.waitForSelector(
        'input[aria-label^="Cc" i], div[aria-label^="Cc" i] input, input[id*="cc" i]',
        { timeout: 5000 }
      ).catch(() => null);
      if (!ccInput) {
        return { repaired: false, reason: 'Cc-Eingabefeld im geöffneten Entwurf nicht gefunden' };
      }
      await ccInput.click();
      await page.waitForTimeout(300);
      await page.keyboard.type(ccStr, { delay: 20 });
      await page.waitForTimeout(900); // Kontaktvorschläge laden lassen
      await page.keyboard.press('Enter'); // Chip bestätigen
      await page.waitForTimeout(600);
      await page.keyboard.press('Escape').catch(() => null); // Vorschlags-Popup schließen
      await page.waitForTimeout(300);
      await page.keyboard.press('Control+s');
      await page.waitForTimeout(3000);
      return { repaired: true };
    } catch (e) {
      return { repaired: false, reason: `CC-Reparatur fehlgeschlagen: ${e && e.message ? e.message : e}` };
    }
  }

  getLastPage() { return this.lastPage; }
}

export const owaClient = new OwaClient();


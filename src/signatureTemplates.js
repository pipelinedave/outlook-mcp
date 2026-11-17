// ===========================================================================
// A2 Realm-Signatur-Templates (17.11.2026, P1-1 / Drafts-Only-Protokoll):
// Es gibt ZWEI echte Signaturen (adesso-Konto + d-velop-Konto). Der frühere
// Ansatz hatte EIN hardcoded default_signature.html (d-velop-HTML) — im
// falschen Realm landete damit die FREMDE Signatur im Draft.
//
// Die Templates werden READ-ONLY aus den OWA-Settings je Postfach extrahiert
// (script: extract-signatures.mjs) und hier als Mapping Realm -> Signatur-HTML
// gehalten. Kein Secret: es ist Davids eigene, ohnehin ausgehende Signatur.
//
// Beweisstandard (Drafts-Only): David sendet Drafts selbst in OWA — der
// persistierte Draft ist exakt das, was gesendet wird. bodyHasSignature()
// prüft den PERSISTIERTEN Draft-Text gegen das Fingerprint des Realm-Templates.
// ===========================================================================

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEMPLATES_PATH = path.join(__dirname, 'signature-templates.json');

let cache = null;

/**
 * Lädt das Template-File (gecacht). Fehlt es oder ist invalid -> {}.
 * @returns {Record<string, {html: string, mailbox?: string, extractedAt?: string, notes?: string}>}
 */
export function loadSignatureTemplates() {
  if (cache) return cache;
  try {
    const raw = JSON.parse(fs.readFileSync(TEMPLATES_PATH, 'utf-8'));
    cache = raw && typeof raw === 'object' ? raw : {};
  } catch {
    cache = {};
  }
  return cache;
}

/** Test-Hilfe: Cache leeren (nach Template-Datei-Wechsel in Tests). */
export function resetSignatureTemplatesCacheForTest() {
  cache = null;
}

/**
 * Template für einen Realm (case-insensitive). null wenn nicht vorhanden.
 * @param {string|null} realm z.B. 'd-velop.de' | 'adesso.de'
 */
export function loadSignatureTemplate(realm) {
  if (!realm) return null;
  const all = loadSignatureTemplates();
  const key = Object.keys(all).find((k) => String(k).toLowerCase() === String(realm).toLowerCase());
  return key ? all[key] : null;
}

/**
 * Distinctive Textmarker aus dem Template-HTML ableiten (reine Textzeilen,
 * die eine echte Signatur eindeutig identifizieren — NICHT generische Gruss-
 * formeln). Kern: Firmen-/Domain-Marker des Templates.
 * @param {{html?: string, mailbox?: string}|null} template
 * @returns {string[]} Marker-Substrings (lowercase)
 */
export function signatureFingerprintMarkers(template) {
  if (!template || !template.html) return [];
  const markers = new Set();
  const mailbox = String(template.mailbox || '').toLowerCase();
  // NUR eindeutige Marker: die VOLLSTÄNDIGE Mailbox-Adresse und Host-Domains.
  // Der Local-Part allein ("david.hallmann") ist in BEIDEN Postfächern identisch
  // und würde die adesso-Signatur fälschlich im d-velop-Draft erkennen (und
  // umgekehrt) — Cross-Realm-False-Positive bewusst ausgeschlossen.
  if (mailbox.includes('@')) markers.add(mailbox);
  const mailto = template.html.match(/mailto:([\w.+-]+@[\w-]+(?:\.[\w-]+)+)/i);
  if (mailto) markers.add(mailto[1].toLowerCase());
  // Link-Domains in der Signatur (z.B. signatur.d-velop.de, www.adesso.de)
  for (const href of template.html.matchAll(/https?:\/\/([\w.-]+)/gi)) {
    const host = href[1].toLowerCase();
    if (host && !/^(go\.micro|aka\.ms|support)/.test(host)) markers.add(host);
  }
  return Array.from(markers).filter((m) => m.length >= 3);
}

/**
 * Ist die Signatur im (persistierten) Draft-Body nachweisbar?
 * @param {string} bodyText persistierter Draft-Text (innerText des Editors)
 * @param {{html?: string, mailbox?: string}|null} template Realm-Template (null = keine Basis)
 * @returns {true|false|null} null = keine Prüfbasis (kein Template)
 */
export function bodyHasSignature(bodyText, template) {
  if (!template || !template.html) return null;
  const markers = signatureFingerprintMarkers(template);
  if (markers.length === 0) return null;
  // Draft-Text ebenfalls NFD->NFC folden (A1) + lowercase; HTML-Entities auflösen
  const text = String(bodyText || '')
    .normalize('NFC')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .toLowerCase();
  if (!text.trim()) return false;
  return markers.some((m) => text.includes(m));
}

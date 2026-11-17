// ===========================================================================
// A2: Signatur-Templates READ-ONLY aus OWA extrahieren (17.11.2026).
//
// Wege (alle read-only): (1) Compose-Deeplink öffnen — OWA setzt die Konto-
// Signatur automatisch in den Message-Body; (2) Fallback "New mail"-Button.
// Die Signatur-HTML wird aus dem Message-Body-Editor gelesen (bevorzugt das
// #Signature-Div), das Compose-Fenster wird OHNE Speichern verworfen.
// OWA-Settings werden NICHT angetastet (Settings zeigte "No signature yet" —
// die Roaming-Signatur ist unkonfiguriert; die echte Signatur lebt im Compose).
//
// Usage: node extract-signatures.mjs [realm1 realm2 ...] (Default: beide Realms)
// Output: src/signature-templates.json (Mapping Realm -> { html, mailbox, ... })
// ===========================================================================
import { browserManager } from './src/browser.js';
import fs from 'fs';
import path from 'path';

const realms = process.argv.slice(2).length ? process.argv.slice(2) : ['d-velop.de', 'adesso.de'];
const OUT_PATH = path.join(process.cwd(), 'src', 'signature-templates.json');

const readBodyEditor = (page) => page.evaluate(() => {
  const sels = [
    'div[aria-label="Message body"][contenteditable="true"]',
    'div[role="textbox"][aria-label*="Message body" i]',
    'div[role="textbox"][aria-label*="Nachrichtentext" i]',
    '[contenteditable="true"][aria-label*="Message body" i]'
  ];
  let e = null;
  for (const sel of sels) { e = document.querySelector(sel); if (e) break; }
  if (!e) return null;
  const sigDiv = e.querySelector('#Signature, div[id^="Signature"]');
  const text = e.innerText || '';
  // Signatur-Nachweis: Personal-Marker im Text (keine leere/neue Mail)
  const hasSigText = /david\s+hallmann|adesso\s+se|d-velop|d\.velop/i.test(text);
  if (!hasSigText) return { found: true, hasSignature: false, html: '', text: text.slice(0, 200) };
  return {
    found: true,
    hasSignature: true,
    // Bevorzugt NUR das Signature-Div (keine trailing <br>-Zeilen des Bodys)
    html: sigDiv ? sigDiv.outerHTML : e.innerHTML,
    via: sigDiv ? '#Signature-div' : 'full-editor',
    text: text.replace(/\s+/g, ' ').slice(0, 200)
  };
});

const discardCompose = async (page) => {
  await page.keyboard.press('Escape').catch(() => {});
  await page.waitForTimeout(1200);
  return page.evaluate(() => {
    const btns = Array.from(document.querySelectorAll('button'));
    const hit = btns.find(b => /verwerfen|discard|nicht speichern|don.t save/i.test((b.innerText || '') + ' ' + (b.getAttribute('aria-label') || '')));
    if (hit) { hit.click(); return 'discarded'; }
    return 'no-prompt';
  });
};

const extractForRealm = async (realm) => {
  const entry = { realm };
  const { page } = await browserManager.ensureContext(true, realm);
  await page.goto(`https://outlook.office.com/mail/?realm=${encodeURIComponent(realm)}`, { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(5000);
  if (page.url().includes('login.microsoftonline.com')) {
    entry.error = 'not-logged-in';
    return entry;
  }

  // Weg 1: Compose-Deeplink (leer) — öffnet neues Compose mit Auto-Signatur
  const deeplink = this_realmDeeplink(realm);
  await page.goto(deeplink, { waitUntil: 'domcontentloaded' }).catch(() => null);
  await page.waitForTimeout(7000);
  let body = await readBodyEditor(page);

  // Weg 2: "New mail"-Button (aria-label variiert je Tenant-Sprache)
  if (!body || !body.hasSignature) {
    await page.goto(`https://outlook.office.com/mail/?realm=${encodeURIComponent(realm)}`, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(5000);
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button, [role="button"]'));
      const hit = btns.find(b => {
        const blob = `${b.getAttribute('aria-label') || ''} ${b.getAttribute('title') || ''} ${b.getAttribute('data-tid') || ''} ${b.innerText || ''}`;
        return /new\s*mail|neue\s*e-?mail|^neu$/i.test(blob) && !/kategorie|category|ordner/i.test(blob);
      });
      if (hit) hit.click();
    });
    await page.waitForTimeout(7000);
    body = await readBodyEditor(page);
  }

  if (!body || !body.found) {
    entry.error = 'Message-Body-Editor nicht gefunden (Compose öffnete nicht)';
  } else if (!body.hasSignature) {
    entry.error = `Compose ohne Auto-Signatur geöffnet (Text: ${body.text})`;
  } else {
    entry.mailbox = realm;
    entry.html = body.html;
    entry.via = body.via || 'compose-editor';
    entry.extractedAt = new Date().toISOString();
    entry.textPreview = body.text;
  }
  // Read-only: Compose verwerfen, NICHTS speichern
  entry.discarded = await discardCompose(page);
  return entry;
};

const this_realmDeeplink = (realm) => (
  `https://outlook.office.com/mail/deeplink/compose?realm=${encodeURIComponent(realm)}`
);

const results = {};
for (const realm of realms) {
  console.error(`[extract-signatures] Realm ${realm}...`);
  let entry;
  try {
    entry = await extractForRealm(realm);
  } catch (e) {
    entry = { realm, error: String(e && e.message ? e.message : e) };
  }
  try { await browserManager.close(realm); } catch {}
  results[realm] = entry;
}

// Merge mit bestehendem File (frische Signatur-Einträge überschreiben, Fehler
// landen unter __errors.* und überschreiben KEINE existierenden Templates)
let merged = {};
try { merged = JSON.parse(fs.readFileSync(OUT_PATH, 'utf-8')); } catch {}
for (const [realm, entry] of Object.entries(results)) {
  if (entry.html) {
    merged[realm] = {
      html: entry.html,
      mailbox: entry.mailbox,
      extractedAt: entry.extractedAt,
      via: entry.via,
      textPreview: entry.textPreview
    };
  } else {
    merged[`__errors.${realm}`] = entry.error;
  }
}
fs.mkdirSync(path.dirname(OUT_PATH), { recursive: true });
fs.writeFileSync(OUT_PATH, JSON.stringify(merged, null, 2) + '\n');
console.log(JSON.stringify(results, (k, v) => (k === 'html' && typeof v === 'string' ? `<${v.length} Bytes>` : v), 2));

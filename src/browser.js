import { chromium } from 'playwright-core';
import path from 'path';
import fs from 'fs';
import { config } from './config.js';

/**
 * Realm-Multiplex (geteilter Daemon, 2026-10-19):
 * Frueher war BrowserManager ein Singleton mit EINER context/page — d.h. der Prozess
 * konnte nur EINEN Realm gleichzeitig offen halten (Realm-Wechsel schloss den Browser).
 * Fuer die HTTP-Daemon-Konsolidierung muessen mehrere Realms (adesso.de / d-velop.de)
 * PARALLEL offen sein, damit mehrere Agents ohne Kollision arbeiten koennen.
 *
 * Umstellung: Eine Map<realm, { context, page, isHeadless }> statt einzelner Felder.
 * ensureContext(realm) legt/behaelt den Browser NUR fuer diesen Realm an und schliesst
 * KEINEN anderen mehr. Alle Realm-Profil-Verzeichnisse (config.getProfileDirForRealm)
 * bleiben erhalten — die eingeloggten Mailbox-Sessions je Tenant werden weiter getrennt.
 */
class BrowserManager {
  constructor() {
    // realm -> { context, page, isHeadless, lastUsed }
    this.browsers = new Map();
    // Idle-Unload (2026-10-20): Browser je Realm nach N Minuten ohne Zugriff
    // schliessen. Profil bleibt erhalten -> naechster Call startet lazy neu.
    // 0 = deaktiviert.
    this.idleMinutes = Number.parseInt(process.env.OUTLOOK_MCP_BROWSER_IDLE_MIN ?? '30', 10);
    this.idleTimer = null;
    this.startIdleReaper();
  }

  startIdleReaper() {
    if (!Number.isFinite(this.idleMinutes) || this.idleMinutes <= 0) return;
    const idleMs = this.idleMinutes * 60_000;
    this.idleTimer = setInterval(() => { void this.reapIdle(idleMs); }, 60_000);
    if (typeof this.idleTimer.unref === 'function') this.idleTimer.unref();
  }

  async reapIdle(idleMs) {
    const now = Date.now();
    for (const [realm, e] of [...this.browsers.entries()]) {
      if (!e?.context) continue;
      if (now - (e.lastUsed || 0) < idleMs) continue;
      if (this.browsers.get(realm) !== e) continue;
      // Vor close austragen: eine parallele ensureContext darf den NEUEN Eintrag
      // nicht durch das nachgelagerte delete verlieren.
      this.browsers.delete(realm);
      try {
        await e.context.close();
      } catch (err) {}
      console.error(`[outlook-mcp] idle-unload: Browser fuer Realm '${realm}' nach ${this.idleMinutes} min Inaktivitaet geschlossen`);
    }
  }

  _entry(realm) {
    return this.browsers.get(realm) || null;
  }

  // Realm-scoped Getter fuer owa.js (statt direkter Singleton-Felder).
  getContext(realm) {
    const e = this._entry(realm);
    if (e) e.lastUsed = Date.now();
    return e ? e.context : null;
  }

  // Test-Hilfe: Reset des internen Zustands ohne echte Browser zu schliessen
  // (Tests mocken ensureContext; die alte `context = null`-Zuweisung geht mit
  // Getter-basiertem Design nicht mehr). Nur fuer Test-Setups gedacht.
  resetForTest() {
    this.browsers = new Map();
  }

  getPage(realm) {
    const e = this._entry(realm);
    if (e) e.lastUsed = Date.now();
    return e ? e.page : null;
  }

  // Abwaertskompatible Aliase: ohne realm = default-Realm (einzelner Browser im alten Stil).
  get context() {
    return this.getContext(config.defaultRealm);
  }

  get page() {
    return this.getPage(config.defaultRealm);
  }

  get currentRealm() {
    return config.defaultRealm;
  }

  get isHeadless() {
    const e = this._entry(config.defaultRealm);
    return e ? e.isHeadless : true;
  }

  async ensureContext(headless = true, realm) {
    const targetRealm = realm || config.defaultRealm;
    let entry = this._entry(targetRealm);

    // Stabilitäts-Regel: Einmal sichtbar (headless=false, z.B. Login-Fenster) gestartet,
    // bleibt der Browser sichtbar. Wir schliessen einen laufenden Kontext NUR, wenn ein
    // SICHTBARES Fenster gewünscht ist, aktuell aber headless läuft. Niemals von sichtbar
    // auf headless zurueckschalten — das wuerde z.B. ein offenes Login-Fenster mitten im
    // Login schliessen. Diese Regel gilt PRO REALM.
    if (entry && entry.context) {
      if (entry.isHeadless && !headless) {
        // Sichtbar gewuenscht, laeuft aber headless -> sichtbar neu oeffnen (nur dieser Realm)
        await this.close(targetRealm);
        entry = null;
      } else {
        try {
          // Session-Gesundheits-Gate: Eine offene Seite, die auf die Microsoft-
          // Login-Seite umgeleitet wurde (Session abgelaufen), ist NICHT gesund.
          if (entry.page && !entry.page.isClosed()) {
            const u = entry.page.url();
            const loggedOut = u.includes('login.microsoftonline.com') ||
              u.includes('login.live.com') ||
              u.includes('error/login') ||
              (u.includes('outlook') && u.includes('sampleresponse')) === false &&
              (u === 'about:blank' || u.includes('error'));
            if (loggedOut) {
              await this.close(targetRealm);
              entry = null;
            } else {
              entry.lastUsed = Date.now(); // Idle-Timer zuruecksetzen
              return { context: entry.context, page: entry.page };
            }
          }
        } catch (e) {
          // Seite nicht mehr abfragbar -> diesen Realm neu aufbauen
          this.browsers.delete(targetRealm);
          entry = null;
        }
      }
    }

    const profileDir = config.getProfileDirForRealm(targetRealm);
    if (!fs.existsSync(profileDir)) {
      fs.mkdirSync(profileDir, { recursive: true });
    }

    // Stale-Chromium-Dateien gruendlich bereinigen: Bei abgestuerzten/versehentlich
    // beendeten Prozessen (Profil-Kollision mehrerer Sessions, Fact #2672) bleiben
    // tote Locks/Port-Dateien zurueck. Nur Dateien entfernen — NIE Prozesse killen.
    const cleanStaleFiles = () => {
      const relics = ['SingletonLock', 'SingletonSocket', 'SingletonCookie', 'DevToolsActivePort', 'lockfile', 'parent.lock'];
      for (const file of relics) {
        const p = path.join(profileDir, file);
        try {
          if (fs.existsSync(p)) fs.unlinkSync(p);
        } catch (e) {}
      }
    };
    cleanStaleFiles();

    const chromeExecutable = config.chromePath;
    if (!chromeExecutable) {
      throw new Error(
        'Kein Chrome/Chromium/Edge gefunden. Bitte installiere Google Chrome / Microsoft Edge oder setze OUTLOOK_MCP_CHROME_PATH.'
      );
    }

    // Bei sichtbaren Fenstern (headless=false) die WSLg-Anzeige-Umgebung HART setzen.
    const wslgEnv = headless
      ? {
          DISPLAY: process.env.DISPLAY || ':0',
          WAYLAND_DISPLAY: process.env.WAYLAND_DISPLAY || 'wayland-0',
          XDG_RUNTIME_DIR: process.env.XDG_RUNTIME_DIR || '/mnt/wslg/runtime-dir'
        }
      : {
          DISPLAY: process.env.DISPLAY || ':0',
          WAYLAND_DISPLAY: 'wayland-0',
          XDG_RUNTIME_DIR: '/mnt/wslg/runtime-dir',
          PULSE_SERVER: process.env.PULSE_SERVER || '/mnt/wslg/PulseServer'
        };

    // Launch mit Retry bei kurzzeitigen Lock-Kollisionen.
    let launched = false;
    let lastError = null;
    let context = null;
    for (let attempt = 0; attempt < 2 && !launched; attempt++) {
      if (attempt > 0) {
        cleanStaleFiles();
        await new Promise(r => setTimeout(r, 800));
      }
      try {
        context = await chromium.launchPersistentContext(profileDir, {
          headless: headless,
          executablePath: chromeExecutable,
          acceptDownloads: true,
          viewport: { width: 1280, height: 800 },
          userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36 Edg/133.0.0.0',
          env: {
            ...process.env,
            ...wslgEnv
          },
          args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-blink-features=AutomationControlled',
            '--ozone-platform=wayland',
            '--disable-dev-shm-usage'
          ]
        });
        launched = true;
      } catch (e) {
        lastError = e;
        context = null;
      }
    }

    if (!launched || !context) {
      throw new Error(`Browser-Start fehlgeschlagen (auch nach Retry): ${lastError ? lastError.message : 'unbekannter Fehler'}`);
    }

    const pages = context.pages();
    const page = pages.length > 0 ? pages[0] : await context.newPage();
    this.browsers.set(targetRealm, { context, page, isHeadless: headless, lastUsed: Date.now() });
    return { context, page };
  }

  async openLoginWindow(realm = config.defaultRealm) {
    await this.close(realm);
    const { page } = await this.ensureContext(false, realm);
    const loginUrl = realm
      ? `https://outlook.office.com/mail/?realm=${encodeURIComponent(realm)}`
      : 'https://outlook.office.com/mail/';
    await page.goto(loginUrl, { waitUntil: 'domcontentloaded' });
    return page;
  }

  // Schliesst einen bestimmten Realm (Standard) oder alle, wenn kein realm angegeben.
  async close(realm) {
    if (realm) {
      const entry = this._entry(realm);
      if (entry && entry.context) {
        try {
          await entry.context.close();
        } catch (e) {}
        this.browsers.delete(realm);
      }
      return;
    }
    for (const [r, e] of [...this.browsers.entries()]) {
      if (e.context) {
        try {
          await e.context.close();
        } catch (err) {}
      }
      this.browsers.delete(r);
    }
  }
}

export const browserManager = new BrowserManager();

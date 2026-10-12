import { chromium } from 'playwright-core';
import path from 'path';
import fs from 'fs';
import { config } from './config.js';

class BrowserManager {
  constructor() {
    this.context = null;
    this.page = null;
    this.isHeadless = true;
    this.currentRealm = null;
  }

  async ensureContext(headless = true, realm) {
    const targetRealm = realm || config.defaultRealm;
    if (this.context && this.currentRealm !== targetRealm) {
      await this.close();
    }
    this.currentRealm = targetRealm;

    // Stabilitäts-Regel: Einmal sichtbar (headless=false, z.B. Login-Fenster) gestartet,
    // bleibt der Browser sichtbar. Wir schließen einen laufenden Kontext NUR, wenn ein
    // SICHTBARES Fenster gewünscht ist, aktuell aber headless läuft. Niemals von sichtbar
    // auf headless zurückschalten — das würde z.B. das sichtbare Login-Fenster mitten im
    // Login schließen und "das Fenster verschwindet" verursachen.
    if (this.context) {
      if (this.isHeadless && !headless) {
        // Sichtbar gewünscht, läuft aber headless -> sichtbar neu öffnen
        await this.close();
      } else {
        try {
          // Session-Gesundheits-Gate: Eine offene Seite, die auf die Microsoft-
          // Login-Seite umgeleitet wurde (Session abgelaufen), ist NICHT gesund.
          // Wir bauen den Context dann sauber neu, statt auf einer toten Seite weiter
          // zu arbeiten (verhindert stillschweigend fehlschlagende Tool-Calls).
          if (this.page && !this.page.isClosed()) {
            const u = this.page.url();
            const loggedOut = u.includes('login.microsoftonline.com') ||
              u.includes('login.live.com') ||
              u.includes('error/login') ||
              (u.includes('outlook') && u.includes('sampleresponse')) === false &&
              (u === 'about:blank' || u.includes('error'));
            if (loggedOut) {
              await this.close();
            } else {
              return { context: this.context, page: this.page };
            }
          }
        } catch (e) {
          // Seite nicht mehr abfragbar -> neu aufbauen
          this.context = null;
          this.page = null;
        }
      }
    }

    const profileDir = config.getProfileDirForRealm(targetRealm);
    if (!fs.existsSync(profileDir)) {
      fs.mkdirSync(profileDir, { recursive: true });
    }

    // Stale-Chromium-Dateien gründlich bereinigen: Bei abgestürzten/versehentlich
    // beendeten Prozessen (Profil-Kollision mehrerer Sessions, Fact #2672) bleiben
    // tote Locks/Port-Dateien zurück, die einen sauberen Neustart blockieren können.
    // Nur Dateien entfernen — NIEMALS verwaiste Chrome-Prozesse killen (könnten zu
    // anderen Sessions gehören).
    const cleanStaleFiles = () => {
      const relics = ['SingletonLock', 'SingletonSocket', 'SingletonCookie', 'DevToolsActivePort', 'lockfile', 'parent.lock'];
      for (const file of relics) {
        const p = path.join(profileDir, file);
        try {
          if (fs.existsSync(p)) fs.unlinkSync(p);
        } catch (e) {
          // Datei evtl. von laufendem Prozess gesperrt -> ignorieren (Chromium kümmert sich)
        }
      }
    };
    cleanStaleFiles();

    const chromeExecutable = config.chromePath;
    if (!chromeExecutable) {
      throw new Error(
        'Kein Chrome/Chromium/Edge gefunden. Bitte installiere Google Chrome / Microsoft Edge oder setze OUTLOOK_MCP_CHROME_PATH.'
      );
    }

    this.isHeadless = headless;
    // Bei sichtbaren Fenstern (headless=false) die WSLg-Anzeige-Umgebung HART setzen,
    // damit Chromium den Wayland-Socket findet und NICHT auf --ozone-platform=headless
    // zurückfällt. Fehlerquelle: Im Prozess ist XDG_RUNTIME_DIR oft /run/user/1000
    // (kein WSLg-Socket) -> ein bloßer ||-Fallback greift dann nicht.
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

    // Launch mit einem Retry: Bei kurzzeitigen Lock-Kollisionen (z.B. ein anderer
    // MCP-Prozess schließt gerade), ein zweiter, bereinigter Versuch. Sonst Fehler
    // transparent durchreichen.
    let launched = false;
    let lastError = null;
    for (let attempt = 0; attempt < 2 && !launched; attempt++) {
      if (attempt > 0) {
        cleanStaleFiles();
        await new Promise(r => setTimeout(r, 800));
      }
      try {
        this.context = await chromium.launchPersistentContext(profileDir, {
          headless: headless,
          executablePath: chromeExecutable,
          // Download-Handling (Tool-Luecke Anhänge, 12.10.2026): Ohne acceptDownloads
          // verwarf Chromium Downloads sofort (gbrain #7330: "kein Download-Handling").
          // Mit true feuert Playwright 'download'-Events, die saveAttachments per
          // saveAs() einfängt.
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
        this.context = null;
      }
    }

    if (!launched) {
      throw new Error(`Browser-Start fehlgeschlagen (auch nach Retry): ${lastError ? lastError.message : 'unbekannter Fehler'}`);
    }

    const pages = this.context.pages();
    this.page = pages.length > 0 ? pages[0] : await this.context.newPage();
    return { context: this.context, page: this.page };
  }

  async openLoginWindow(realm = config.defaultRealm) {
    await this.close();
    const { page } = await this.ensureContext(false, realm);
    const loginUrl = realm
      ? `https://outlook.office.com/mail/?realm=${encodeURIComponent(realm)}`
      : 'https://outlook.office.com/mail/';
    await page.goto(loginUrl, { waitUntil: 'domcontentloaded' });
    return page;
  }

  async close() {
    if (this.context) {
      try {
        await this.context.close();
      } catch (e) {}
      this.context = null;
      this.page = null;
    }
  }
}

export const browserManager = new BrowserManager();

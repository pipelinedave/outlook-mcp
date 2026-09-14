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

    if (this.context) {
      if (this.isHeadless !== headless) {
        await this.close();
      } else {
        try {
          if (this.page && !this.page.isClosed()) {
            return { context: this.context, page: this.page };
          }
        } catch (e) {}
      }
    }

    const profileDir = config.getProfileDirForRealm(targetRealm);
    if (!fs.existsSync(profileDir)) {
      fs.mkdirSync(profileDir, { recursive: true });
    }

    // Clean stale lock files if any
    try {
      const files = ['SingletonLock', 'SingletonSocket', 'SingletonCookie'];
      for (const file of files) {
        const p = path.join(profileDir, file);
        if (fs.existsSync(p)) {
          fs.unlinkSync(p);
        }
      }
    } catch (e) {}

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

    this.context = await chromium.launchPersistentContext(profileDir, {
      headless: headless,
      executablePath: chromeExecutable,
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
        '--ozone-platform=wayland'
      ]
    });

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

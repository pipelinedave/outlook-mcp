import path from 'path';
import os from 'os';
import fs from 'fs';

const HOME = os.homedir();

// Chrome/Edge-Executable automatisch erkennen (WSL/Linux, macOS, Windows).
// Über per-Env OUTLOOK_MCP_CHROME_PATH überschreibbar.
function detectChromePath() {
  if (process.env.OUTLOOK_MCP_CHROME_PATH) return process.env.OUTLOOK_MCP_CHROME_PATH;

  const candidates = [
    // Playwright-Cache (Linux, gängige Versionen)
    ...expandPlaywrightCache(),
    // System-Chrome / Chromium (Linux)
    '/usr/bin/google-chrome',
    '/usr/bin/google-chrome-stable',
    '/usr/bin/chromium',
    '/usr/bin/chromium-browser',
    '/usr/bin/microsoft-edge',
    '/usr/bin/microsoft-edge-stable',
    // macOS
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Chromium.app/Contents/MacOS/Chromium',
    '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    // Windows
    'C:\\\\Program Files\\\\Google\\\\Chrome\\\\Application\\\\chrome.exe',
    'C:\\\\Program Files (x86)\\\\Google\\\\Chrome\\\\Application\\\\chrome.exe',
    'C:\\\\Program Files (x86)\\\\Microsoft\\\\Edge\\\\Application\\\\msedge.exe',
    'C:\\\\Program Files\\\\Microsoft\\\\Edge\\\\Application\\\\msedge.exe'
  ];

  for (const c of candidates) {
    if (fs.existsSync(c)) return c;
  }
  return null;
}

function expandPlaywrightCache() {
  const base = path.join(HOME, '.cache/ms-playwright');
  const results = [];
  try {
    if (fs.existsSync(base)) {
      const entries = fs.readdirSync(base);
      for (const e of entries) {
        const p = path.join(base, e, 'chrome-linux64', 'chrome');
        if (fs.existsSync(p)) results.push(p);
      }
    }
  } catch (e) {}
  return results;
}

function getProfileDir() {
  return process.env.OUTLOOK_MCP_PROFILE_DIR || path.join(HOME, '.outlook-browser-profile');
}

function getDefaultRealm() {
  return process.env.OUTLOOK_MCP_REALM || 'd-velop.de';
}

function getDefaultHeadless() {
  const v = process.env.OUTLOOK_MCP_HEADLESS;
  if (v === undefined) return true;
  return String(v).toLowerCase() !== 'false' && v !== '0';
}

const detectedChromePath = detectChromePath();
if (!detectedChromePath) {
  console.warn('[outlook-mcp] Kein Chrome/Chromium/Edge automatisch gefunden. Setze OUTLOOK_MCP_CHROME_PATH, falls Playwright keinen eigenen Browser mitbringt.');
}

export const config = {
  get chromePath() {
    return process.env.OUTLOOK_MCP_CHROME_PATH || detectedChromePath || undefined;
  },
  profileDir: getProfileDir(),
  defaultRealm: getDefaultRealm(),
  defaultHeadless: getDefaultHeadless()
};

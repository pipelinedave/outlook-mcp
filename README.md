# Outlook MCP Server

Microsoft Outlook Web (OWA) MCP-Server auf Basis einer **Playwright Browser-Bridge**.
Steuert den echten Outlook Webclient im Chromium, Chrome oder Edge über CDP – keine
komplexen Azure AD App Registrierungen oder Graph-API Admin-Consents erforderlich.
Funktioniert direkt mit echten Browser-Sessions und ohne Cloud-Relay.

Ermöglicht LLM-Agenten (Antigravity, opencode, Claude Desktop, Cursor, …) den Lese-
und Entwurfszugriff auf Microsoft Outlook: E-Mails suchen, Posteingang und Ordner
durchsuchen, E-Mails als Markdown inklusive Teams-Besprechungslinks auslesen sowie
neue E-Mail-Entwürfe sicher anlegen.

> **Hinweis**: Dieses Repository ist vollständig organisationsneutral. Es enthält keine
> firmenspezifischen Accounts oder Zugangsdaten. Der gewünschte Tenant/Realm (z.B.
> `adesso.de`, `d-velop.de` oder eine beliebige Microsoft 365 Domain) kann frei über
> Umgebungsvariablen oder direkt im Login-Tool angegeben werden.

---

## Features

- **14 leistungsfähige MCP-Tools**:
  - `outlook_status`: Prüft Verbindungs- und Anmeldestatus von Outlook Web.
  - `outlook_login`: Öffnet ein sichtbares Browserfenster für den einmaligen interaktiven Login (inkl. SSO / MFA).
  - `outlook_search`: Volltextsuche über das OWA-Suchfeld (nach Kunden, Tickets, Absendern, Stichwörtern).
  - `outlook_get_email`: Liest eine E-Mail vollständig im Lesebereich aus (konvertiert in Markdown, extrahiert Teams-Links und Metadaten).
  - `outlook_list_recent`: Schnelle Übersicht über die neuesten E-Mails im Posteingang.
  - `outlook_create_draft`: Legt neue E-Mail-Entwürfe im Ordner „Entwürfe“ (Drafts) an (per OWA DeepLink & Ctrl+S).
  - `outlook_reply_draft`: Erstellt einen nativen Antwort-Entwurf (Reply / Reply All) direkt auf eine bestehende E-Mail. Behält vollständigen E-Mail-Verlauf (Thread/Zitat) und die automatische Outlook-Signatur bei.
  - `outlook_list_folders`: Listet alle Navigationsordner (Posteingang, Gesendet, Archiv, Entwürfe etc.).
  - `outlook_navigate_folder`: Wechselt gezielt in einen bestimmten Ordner (z.B. Sent Items, Archiv).
  - `outlook_bulk_list`: Scroll-Enumeration zur Überwindung der OWA-Listen-Virtualisierung (erfasst alle Conversation-IDs).
  - `outlook_read_email`: Liest gezielt eine E-Mail anhand ihrer Conversation-ID aus.
  - `outlook_toggle_flag`: Toggelt bzw. entfernt das Flag (Kennzeichnung) einer E-Mail anhand ihrer Conversation-ID.
   - `outlook_archive_email`: Verschiebt eine E-Mail anhand ihrer Conversation-ID ins Archiv.
   - `outlook_accept_meeting`: Nimmt eine Termineinladung (Meeting Request) in Outlook Web an (RSVP "Annehmen"/Accept). Sendet die Annahme-Rückmeldung an die Gegenseite, trägt den Termin in den Firmenkalender ein und übernimmt den Outlook-typischen Archiv-Flow. Lokalisierung per Conversation-ID oder Suchbegriff.
   - `outlook_close`: Beendet die Hintergrund-Browserinstanz.
- **Sicherheits-Konzept (Draft-First)**: Es gibt kein Blindversenden von Mails. Der Server legt Entwürfe an, sodass der Anwender die Nachricht vor dem Absenden in Outlook sichten und freigeben kann.
- **Auto-Erkennung von Browsern**: Findet automatisch Chromium/Chrome/Edge unter Linux, WSL, macOS und Windows (oder via `OUTLOOK_MCP_CHROME_PATH`).
- **Persistentes lokales Profil**: Anmeldedaten und MFA-Tokens verbleiben sicher im lokalen Profil (`~/.outlook-browser-profile`).

---

## Voraussetzungen

- **Node.js ≥ 20**
- Ein installierter Browser: Google Chrome, Microsoft Edge oder Playwright Chromium (`npx playwright install chromium`).
- (Unter WSL/Linux) WSLg oder ein laufender X-Server für das einmalige Login-Fenster.

---

## Installation

```bash
git clone https://github.com/pipelinedave/outlook-mcp.git
cd outlook-mcp
npm install
```

---

## Konfiguration (Umgebungsvariablen)

Alle Umgebungsvariablen sind optional:

| Variable | Beschreibung | Standard |
|---|---|---|
| `OUTLOOK_MCP_REALM` | Standard-Tenant/Domain für OWA (z.B. `adesso.de`, `d-velop.de`) | `d-velop.de` |
| `OUTLOOK_MCP_CHROME_PATH` | Expliziter Pfad zur Browser-Executable | Automatische Erkennung |
| `OUTLOOK_MCP_PROFILE_DIR` | Speicherort für das Browser-Profil | `~/.outlook-browser-profile` |
| `OUTLOOK_MCP_HEADLESS` | Headless-Modus (`true` / `false`) | `true` |

---

## Einbindung in MCP-Clients

### Antigravity / Claude Desktop / Cursor / opencode

Füge den Server in deine MCP-Konfigurationsdatei (z.B. `mcp_config.json` oder `opencode.json`) ein:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/pfad/zu/outlook-mcp/index.js"],
      "env": {
        "OUTLOOK_MCP_REALM": "adesso.de"
      }
    }
  }
}
```

Oder unter WSL mit Node-Wrapper:

```bash
#!/bin/bash
export PATH="$HOME/.nvm/versions/node/v20.20.2/bin:$PATH"
export DISPLAY="${DISPLAY:-:0}"
export WAYLAND_DISPLAY="${WAYLAND_DISPLAY:-wayland-0}"
export XDG_RUNTIME_DIR="${XDG_RUNTIME_DIR:-/mnt/wslg/runtime-dir}"
export PULSE_SERVER="${PULSE_SERVER:-/mnt/wslg/PulseServer}"

exec node /pfad/zu/outlook-mcp/index.js "$@"
```

---

## Erste Schritte (Login)

1. Rufe das Tool `outlook_login` auf:
   ```json
   {
     "realm": "adesso.de"
   }
   ```
2. Es öffnet sich ein sichtbares Browserfenster. Melde dich dort wie gewohnt bei deinem Microsoft 365 Account an (inkl. MFA/Passkey).
3. Sobald dein Posteingang geladen ist, kannst du das Fenster schließen oder geöffnet lassen.
4. Prüfe mit `outlook_status` den Verbindungsstatus.
5. Das Profil wird dauerhaft lokal gespeichert – zukünftige MCP-Aufrufe laufen vollautomatisch im Hintergrund (Headless).

---

## Tests & Syntaxprüfung

```bash
npm run check       # Überprüft JS-Syntax aller Dateien
npm test            # Führt Unit-Tests via node:test aus
```

---

## Lizenz

MIT License – siehe [LICENSE](LICENSE).

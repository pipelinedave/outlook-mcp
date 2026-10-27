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

- **24 leistungsfähige MCP-Tools**:
  - `outlook_status`: Prüft Verbindungs- und Anmeldestatus von Outlook Web.
  - `outlook_login`: Öffnet ein sichtbares Browserfenster für den einmaligen interaktiven Login (inkl. SSO / MFA).
  - `outlook_set_realm`: Schaltet den aktiven Tenant/Realm um (Multi-Tenant, z.B. `d-velop.de` ↔ `adesso.de`). Jeder Realm nutzt ein eigenes persistentes Browser-Profil, damit die Mailbox-Sessions nicht vermischt werden.
  - `outlook_search`: Volltextsuche über das OWA-Suchfeld (nach Kunden, Tickets, Absendern, Stichwörtern).
  - `outlook_get_email`: Liest eine E-Mail vollständig im Lesebereich aus (konvertiert in Markdown, extrahiert Teams-Links und Metadaten).
  - `outlook_list_attachments`: **READ-ONLY** – listet die Anlagen einer E-Mail im Lesebereich auf (Anlagen-Chips mit Dateinamen + Inline-Bilder mit Größen). Vorstufe zum Speichern.
  - `outlook_save_attachments`: Speichert die Anlagen einer E-Mail auf die Festplatte. Anlagen-Chips per Hover-Download-Button (echtes Playwright-Download-Event, `acceptDownloads`), Inline-Bilder (`blob:`/`data:`-URIs im Body) als Fallback. `destDir` (unter dem Home-Verzeichnis) oder Default `~/projects/outlook-attachments/<Zeitstempel>/`. Kollisionsfreie Dateinamen, Diagnostik pro Anhang.
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
- **Firmenkalender-Tools** (für Blocker & Fokuszeiten im Outlook-Firmenkalender):
  - `outlook_create_calendar_event`: Legt einen Kalender-Eintrag (Standard `transparency: busy`) per OWA DeepLink an. Termin-Zeiten werden korrekt als Berlin-Lokalzeit interpretiert (keine UTC-Verschiebung).
  - `outlook_get_freebusy`: Fragt die Belegung in einem Zeitfenster ab. OHNE `attendees`: eigener Firmenkalender (Busy-Slots aus ARIA-Labels). MIT `attendees` (Array von E-Mails): Verfügbarkeit der angegebenen Teilnehmer über das OWA-GraphQL-Gateway (Free/Busy/Tentative/OOF + gemeinsame freie Slots) — zur kollisionsfreien Terminvorschlag-Findung.
  - `outlook_list_calendar_events`: **READ-ONLY** – listet ALLE Termine (Titel, Datum, Start/Ende, Ganztag, Status) des Firmenkalenders in einem Zeitraum. Basis für Termin-Briefings. Es wird nur gelesen, nichts verändert.
  - `outlook_get_calendar_event`: **READ-ONLY** – liest die Detail-Ansicht (Reading-Panel) eines Kalender-Events anhand seines Titels. Liefert den vollständigen Detail-Text für Termin-Briefings.
  - `outlook_delete_calendar_event`: Löscht ein Kalender-Event anhand seines Titels/Suchbegriffs (über die Wochenansicht + echten Playwright-Klick). Entfernt einzelne Events, verhindert versehentliches Löschen ganzer Serien.
  - `outlook_update_calendar_event`: Bearbeitet einen bestehenden Kalender-Eintrag (Titel/Zeit/Beschreibung) anhand seines Suchbegriffs.
- **Sicherheits-Konzept (Draft-First)**: Es gibt kein Blindversenden von Mails. Der Server legt Entwürfe an, sodass der Anwender die Nachricht vor dem Absenden in Outlook sichten und freigeben kann.
- **Verifizierte Reply-Textübernahme** (live verifiziert 09.10.2026): `outlook_reply_draft` tippt den Antworttext per echter Tastatureingabe in den Editor — OWA serialisiert aus seinem internen Editor-Modell, eine reine DOM-Manipulation (`insertBefore`) wird dabei **nicht** übernommen (live belegt: Entwurf gespeichert, Body leer). Nach dem Speichern wird der Entwurf erneut geöffnet und der persistierte Body geprüft (`persistedBodyVerified`). Schlägt dieser Nachweis fehl, meldet das Tool `success: false` und legt keinen scheinbar erfolgreichen, in Wahrheit leeren Entwurf an — der Entwurf bleibt zur Sichtung liegen. Ein Unversehrtheits-Vorher/Nachher-Vergleich bricht ab, falls Verlauf oder Signatur beim Tippen verloren gehen. Derselbe Standard gilt für `outlook_create_draft` (live im adesso-Realm verifiziert); beide nutzen die gemeinsame Nachweis-Methode.
- **Multi-Tenant (Realm pro Tool)**: Jedes Tool akzeptiert einen optionalen `realm`-Parameter (z.B. `adesso.de`, `d-velop.de`), um gezielt das Konto zu wechseln. Live verifiziert am 09.10.2026: beide Konten sind angemeldet und liefern nachweislich verschiedene Postfächer (`David.Hallmann@d-velop.de` vs. `David.Hallmann@adesso.de`) mit disjunkten Conversation-IDs und getrennten Profilen. Der Realm gilt bis zum nächsten Wechsel („sticky“), `outlook_set_realm` setzt ihn explizit. Ungültige Werte (Pfad-Trenner, Shell-Metazeichen, zu kurz) weist der Server ab.
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

## Firmenkalender-Blocker (Fokuszeiten)

Die Kalender-Tools bedienen ausschließlich den **Outlook-Firmenkalender** (OWA).
Nützlich, um Fokusblöcke/Abwesenheiten anzulegen, Kollisionen zu prüfen und Blocker zu pflegen bzw. zu entfernen.

```json
// 1) Vorab Kollisionsprüfung (Busy-Slots im Fenster auslesen)
{
  "name": "outlook_get_freebusy",
  "arguments": { "start": "2026-09-15T09:30:00", "end": "2026-09-15T11:00:00" }
}

// 2) Blocker anlegen (08:00–09:00 Berlin, busy)
{
  "name": "outlook_create_calendar_event",
  "arguments": { "subject": "Fokus: tiefe Arbeit", "start": "2026-09-15T08:00:00", "end": "2026-09-15T09:00:00" }
}

// 3) Blocker löschen (per Titel-Fragment)
{
  "name": "outlook_delete_calendar_event",
  "arguments": { "query": "Fokus: tiefe Arbeit" }
}
```

### Teilnehmer-Verfügbarkeit (attendees-Parameter)

Mit `attendees` fragt `outlook_get_freebusy` die Verfügbarkeit der ANGEGEBENEN
Personen ab (statt des eigenen Kalenders) — für kollisionsfreie Terminvorschläge:

```json
{
  "name": "outlook_get_freebusy",
  "arguments": {
    "attendees": ["silas.beckmann@d-velop.de", "mathis.kuenzel@d-velop.de"],
    "start": "2026-10-29T09:00:00",
    "end": "2026-10-29T17:00:00",
    "interval_minutes": 15
  }
}
```

Rückgabe: je Teilnehmer `busySlots` (Free/Busy/Tentative/OOF, UTC + Berlin-Lokalzeit),
`scheduleItems` (Betreff sofern nicht privat), `workingHours` — plus
`commonFreeSlots` (gemeinsame freie Intervalle ≥ Slot-Granularität).

Technik (Live-verifiziert 27.10.2026): Abfrage über das OWA-GraphQL-Gateway
(`POST /outlookgatewayb2/graphql`, Operation `GetSchedule`). Der Bearer-Token wird
beim Kalender-Boot aus dem Eigen-Traffic der App abgegriffen (JWT, ~24h gültig,
gecacht — Folgeaufrufe < 1s), `userId` kommt aus dem GetSchedule-Body bzw. den
JWT-Claims. Gesamtzeitbudget 40s (sicher unter dem MCP-Client-Timeout); Fehler
je Teilnehmer werden einzeln berichtet, statt den Call zu killen. Der Hinweis in
`note` warnt, wenn Teilnehmer fehlschlugen — deren Belegung ist dann NICHT enthalten.

> **Zeitzonen-Hinweis**: Zeiten ohne Offset (z.B. `2026-09-15T08:00:00`) werden als
> **Europe/Berlin-Lokalzeit** interpretiert (OWA interpretiert `startdt` ohne Offset als
> lokale Browserzeit). Ein expliziter Offset (`+02:00`) bzw. `Z` wird nach UTC-ISO normalisiert.

---

## READ-ONLY Kalender-Lese-Tools (Termin-Briefings)

Für den `otto`-Agenten stehen zwei **strikt lesende** Tools zur Verfügung, um
Termin-Briefings aus dem Outlook-Firmenkalender zu bauen. Sie verändern nichts –
es wird nur navigiert und gelesen.

```jsonc
// 1) Alle Termine eines Zeitraums listen (Titel, Datum, Start/Ende, Ganztag, Status)
{
  "name": "outlook_list_calendar_events",
  "arguments": { "start": "2026-10-02", "end": "2026-10-09" }
}
// -> { status, timespanText, weekVerified, dataQuality, events: [...], count }

// 2) Detail-Ansicht (Reading-Panel) eines einzelnen Events lesen
{
  "name": "outlook_get_calendar_event",
  "arguments": { "query": "Amt Kellinghusen Citrix Add-In" }
}
// -> { status, query, titleMatched, detailText }
```

Ehrlichkeits-Garantien:
- Die Zielwoche wird ausschließlich am **angezeigten Wochen-Header** verifiziert;
  falsche Wochen liefern bewusst keine Daten (`renderedWeekContains`).
- `dataQuality` ist nur `'verified'`, wenn die Woche verifiziert **und**
  Event-Knoten gelesen wurden – sonst wird die Liste ehrlich als unvollständig
  ausgewiesen.
- Pseudo-Slots (Arbeitsort, „(ohne Titel)“, URLs) und Duplikate werden entfernt.
- `outlook_get_calendar_event` öffnet das Event per Doppelklick im lesenden
  Reading-Panel – keine Schreiboperation.

### Anbindung an den `otto`-Agenten (otto.md)

Der `otto`-Agent hat in seiner Agent-Definition (`agents/otto.md`) bereits die
Permissive `outlook_*: allow` und erhält die neuen Tools damit **automatisch**,
sobald dieser Server bei ihm gemountet ist. Für Termin-Briefings sind die
Tool-Namen so zu ergänzen (analog `outlook_outlook_get_freebusy`):

```md
- `outlook_outlook_list_calendar_events` (READ-ONLY: alle Termine eines Zeitraums listen — Basis für Termin-Briefings)
- `outlook_outlook_get_calendar_event` (READ-ONLY: Detail-Ansicht eines Events per Titel lesen)
```

Ablauf für ein Briefing: `outlook_list_calendar_events(start, end)` liefert die
Terminübersicht, dann `outlook_get_calendar_event(query)` die Details zu den
relevanten Terminen. Beide sind strikt lesend – sie verändern den Kalender nie.

---

## Tests & Syntaxprüfung

```bash
npm run check       # Überprüft JS-Syntax aller Dateien
npm test            # Führt Unit-Tests via node:test aus
```

---

## Lizenz

MIT License – siehe [LICENSE](LICENSE).

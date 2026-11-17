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
  - `outlook_get_email`: Liest eine E-Mail vollständig im Lesebereich aus (konvertiert in Markdown, extrahiert Teams-Links und Metadaten). Bei Konversations-IDs wird die **neueste Thread-Mail** gerendert (die Mail, die Suche/Liste in der Preview getroffen haben) — die vollständige Konversations-Historie liegt als Metadatum `conversation.messages` bei (neueste zuerst, Drafts markiert); optionaler Parameter `message_index` rendert eine andere Historien-Nachricht.
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
   - `outlook_decline_meeting`: Lehnt eine Termineinladung (Meeting Request) in Outlook Web ab (RSVP "Absagen"/Decline). Sendet die Absage-Rückmeldung an die Gegenseite (optional mit Kommentar), trägt den Termin NICHT ein. Lokalisierung per Conversation-ID oder Suchbegriff.
   - `outlook_close`: Beendet die Hintergrund-Browserinstanz.
- **Firmenkalender-Tools** (für Blocker & Fokuszeiten im Outlook-Firmenkalender):
  - `outlook_create_calendar_event`: Legt einen Kalender-Eintrag (Standard `transparency: busy`) per OWA DeepLink an. Termin-Zeiten werden korrekt als Berlin-Lokalzeit interpretiert (keine UTC-Verschiebung).
  - `outlook_get_freebusy`: Fragt die Belegung in einem Zeitfenster ab. OHNE `attendees`: eigener Firmenkalender (Busy-Slots aus ARIA-Labels). MIT `attendees` (Array von E-Mails): Verfügbarkeit der angegebenen Teilnehmer über das OWA-GraphQL-Gateway (Free/Busy/Tentative/OOF + gemeinsame freie Slots) — zur kollisionsfreien Terminvorschlag-Findung.
  - `outlook_list_calendar_events`: **READ-ONLY** – listet ALLE Termine (Titel, Datum, Start/Ende, Ganztag, Status) des Firmenkalenders in einem Zeitraum. Basis für Termin-Briefings. Es wird nur gelesen, nichts verändert. Wochen-Navigation läuft über die OWA-Nav-Buttons mit **Label-Verriegelung** (die Buttons nennen ihre Zielwoche im ARIA-Label, z.B. "Go to next week 9–15 November, 2026 (Week 46)"); der `startdate`-URL-Parameter wird von OWA nicht zuverlässig befolgt (nur noch Optimierungsversuch). Header/Label-Konflikte (OWA-Boot-Flash) werden erkannt und über die Label-abgeleitete Woche aufgelöst.
  - `outlook_get_calendar_event`: **READ-ONLY** – liest die Detail-Ansicht (Reading-Panel) eines Kalender-Events anhand seines Titels. Liefert den vollständigen Detail-Text für Termin-Briefings. **`date`-Parameter (empfohlen):** ISO-Datum in der Zielwoche des Events — ohne `date` wird nur die AKTUELLE Woche durchsucht (Tool-Lücke 05.11.2026: Events in anderen Wochen waren unsichtbar).
  - `outlook_delete_calendar_event`: Löscht ein Kalender-Event anhand seines Titels/Suchbegriffs (über die Wochenansicht + echten Playwright-Klick). Entfernt einzelne Events, verhindert versehentliches Löschen ganzer Serien. **`date`-Parameter (empfohlen):** navigiert verifiziert zur Zielwoche des Events (gleiche Navigations-Primitive wie `list_calendar_events`: startdate-URL, dann verriegelte Nav-Button-Klicks); ohne `date` nur aktuelle Woche. Verify mit Re-Poll gegen OWA Render-Lag — ehrliche Semantik: nicht verifizierbar ⇒ `success: false`, nie ein falsches "gelöscht".
  - `outlook_update_calendar_event`: Bearbeitet einen bestehenden Kalender-Eintrag (Titel/Zeit/Beschreibung) anhand seines Suchbegriffs. **`date`-Parameter (empfohlen):** wie delete — ohne `date` nur aktuelle Woche.
- **Sicherheits-Konzept (Draft-First)**: Es gibt kein Blindversenden von Mails. Der Server legt Entwürfe an, sodass der Anwender die Nachricht vor dem Absenden in Outlook sichten und freigeben kann.
- **Verifizierte Reply-Textübernahme** (live verifiziert 09.10.2026): `outlook_reply_draft` tippt den Antworttext per echter Tastatureingabe in den Editor — OWA serialisiert aus seinem internen Editor-Modell, eine reine DOM-Manipulation (`insertBefore`) wird dabei **nicht** übernommen (live belegt: Entwurf gespeichert, Body leer). Nach dem Speichern wird der Entwurf erneut geöffnet und der persistierte Body geprüft (`persistedBodyVerified`). Schlägt dieser Nachweis fehl, meldet das Tool `success: false` und legt keinen scheinbar erfolgreichen, in Wahrheit leeren Entwurf an — der Entwurf bleibt zur Sichtung liegen. Ein Unversehrtheits-Vorher/Nachher-Vergleich bricht ab, falls Verlauf oder Signatur beim Tippen verloren gehen. Derselbe Standard gilt für `outlook_create_draft` (live im adesso-Realm verifiziert); beide nutzen die gemeinsame Nachweis-Methode.
- **Realm-Signatur-Templates (A2, 17.11.2026)**: Es gibt ZWEI echte Signaturen (adesso-Konto + d-velop-Konto). `src/signature-templates.json` hält die READ-ONLY aus dem jeweiligen Postfach extrahierte Signatur-HTML (Script: `node extract-signatures.mjs [realm...]` — öffnet Compose, liest die Auto-Signatur, verwirft ohne Speichern). Mapping Realm→Template: createDraft injiziert bei fehlender Auto-Signatur das Template DES Realms (erst ohne Template greift die alte `default_signature.html`); replyDraft assertiert am PERSISTIERTEN Draft (`signaturePresent`), dass die Signatur des Realms drin ist — mit Re-Poll (OWA rendert sie verzögert), ehrlich `unverified` ohne Template und `success:false` bei fehlender erwarteter Signatur. Marker sind bewusst NUR die volle Mailbox-Adresse + Host-Domains (der geteilte Local-Part `david.hallmann` erzeugte sonst Cross-Realm-False-Positives).
- **CC-Merge + Recipient-Readback (A1, 17.11.2026)**: `outlook_reply_draft` akzeptiert optional `cc` (Semikolon/Komma-getrennt) und merged sie nach dem Body-Insert in den CC-Well; nach dem Speichern werden die PERSISTIERTEN Empfänger mitgelesen (`recipients`) und gegen den Soll-Zustand verglichen. CC als "An" gelandet wird ehrlich gemeldet (`ccLocation: "to"`), fehlende CC werden via `_repairDraftCc` nachgezogen; bleibt der Nachweis aus → `success:false` (Drafts-Only: der persistierte Draft ist exakt das, was David sendet). `recipientListContains` foldet NFC/NFD (OWA-Chips kommen teils NFD-normalisiert zurück).
- **Perspektive: Graph-API statt OWA-UI (Entscheidung 17.11.2026, NICHT gebaut)**: OWA-/Teams-UI-Automatisierung bleibt der Pfad; ein Graph-API-Pfad (Send/Read via Microsoft Graph statt Browser-DOM) kann später andocken — Anknüpfungspunkt: Kollege Yannick Bülter hat im Teams-Setup bereits einen Graph-API-Token reverse engineered. Der Wechsel würde die Browser-Flows (getPage/Realm-Multiplex/DOM-Selektoren) ersetzen, nicht die Beweisstandard-Schicht (persistenter Draft-Readback, Empfänger-/Signatur-Assertion bleiben unverändert als Verifikationsziel).
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

### Determinismus-Garantien (Runde 2, Live-Befunde 05.11.2026)

- **`outlook_get_email` mit ID/Conversation-ID rendert NIE eine andere Mail:** Ist die ID
  in der aktuell gerenderten Liste nicht enthalten, läuft eine begrenzte Scroll-Enumeration;
  danach ein ehrlicher Fehler mit Recovery-Hinweis (`outlook_search`) — niemals ein
  `rows[0]`-Fallback (Live-Befund: 3 identische Calls, 3 verschiedene Fremdmails).
- **Realm-Lock:** Pro Tenant-Realm läuft ein Tool-Call atomar (Promise-Kette in
  `src/realm-lock.js`); parallele Agenten/Sessions auf dem geteilten Daemon-Browser
  interleaven nicht mehr in fremde UI-Flows. Verschiedene Realms bleiben parallel.
- **Kalender-Verifikation:** Monats-/Datum-Picker-Labels (`"November 2026, select to
  change the month"` → Phantom-Woche aus „November 20(26)") sind nie eine Wochen-/Datums-
  Quelle; vor dem Wochen-Nachweis re-pollt der Flow auf einen GELESENEN Header — eine
  korrekt gerenderte Woche liefert daher `verified` bzw. `verified-empty`, nie
  `unverified` bei korrekter Navigation.

---

## Lizenz

MIT License – siehe [LICENSE](LICENSE).

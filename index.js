#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema
} from '@modelcontextprotocol/sdk/types.js';
import { browserManager } from './src/browser.js';
import { owaClient } from './src/owa.js';
import { config } from './src/config.js';
import { withRealmLock } from './src/realm-lock.js';

const server = new Server(
  {
    name: 'outlook-mcp',
    version: '1.0.0'
  },
  {
    capabilities: {
      tools: {}
    }
  }
);

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: 'outlook_status',
        description: 'Prüft den Status der Outlook Web Verbindung und ob du eingeloggt bist.',
        inputSchema: {
          type: 'object',
          properties: {
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          }
        }
      },
      {
        name: 'outlook_set_realm',
        description: `Schaltet den aktiven Tenant/Realm der Bridge um (Multi-Tenant). Wechselt bei Bedarf das persistentes Browser-Profil (pro Realm ein eigenes Profil, damit die Mailbox-Sessions nicht vermischt werden). Bekannte Realms: Standard "${config.defaultRealm}", z.B. adesso.de oder d-velop.de. Nach dem Wechsel outlook_status aufrufen, um die Mailbox-Zuordnung zu prüfen.`,
        inputSchema: {
          type: 'object',
          properties: {
            realm: {
              type: 'string',
              description: `Tenant/Domain (Standard: ${config.defaultRealm}, z.B. adesso.de oder d-velop.de)`
            }
          },
          required: ['realm']
        }
      },
      {
        name: 'outlook_login',
        description: 'Öffnet ein sichtbares Browserfenster (Edge/Chromium via WSLg/Desktop), damit du dich einmalig bei Microsoft 365 / OWA anmelden kannst.',
        inputSchema: {
          type: 'object',
          properties: {
            realm: {
              type: 'string',
              description: `Tenant/Domain (Standard: ${config.defaultRealm}, z.B. adesso.de oder d-velop.de)`,
              default: config.defaultRealm
            }
          }
        }
      },
      {
        name: 'outlook_search',
        description: 'Sucht E-Mails in Outlook Web anhand eines Suchbegriffs (z.B. Kundenname, Ticketnummer, Absender). Bei leerer OWA-Suche automatischer Fallback: Datumsfenster-Enumeration im Sent-Ordner mit Empfänger-Filter; wenn weiterhin nichts gefunden: klarer NOT FOUND-Return.',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Suchbegriff (z.B. "Fit-Connect", "Bülter", "Rechnung")'
            },
            limit: {
              type: 'number',
              description: 'Maximale Anzahl an Ergebnissen (Standard: 10)',
              default: 10
            },
            folder: {
              type: 'string',
              description: 'Optional: Ordner-Scoping. "sentitems" erzwingt die Sent-Enumeration (Fallback für Mail-Indexierungslücken versendeter Mails).'
            },
            to: {
              type: 'string',
              description: 'Optional: Empfänger-Adresse oder -Name für den Sent-Fallback-Filter (nur relevant mit folder="sentitems" oder wenn OWA-Suche leer bleibt).'
            },
            date_from: {
              type: 'string',
              description: 'Optional: Fensterstart für die Sent-Enumeration ("YYYY-MM-DD" oder "DD.MM.YYYY").'
            },
            date_to: {
              type: 'string',
              description: 'Optional: Fensterende für die Sent-Enumeration ("YYYY-MM-DD" oder "DD.MM.YYYY").'
            },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de") für den gezielten Kontenwechsel.' }
          },
          required: ['query']
        }
      },
      {
        name: 'outlook_get_email',
        description: 'Liest eine E-Mail im Lesebereich vollständig aus (Markdown, Teams-Links, Metadaten). Standardmäßig die oberste E-Mail (Index 0). Bei Konversations-IDs wird die NEUESTE Thread-Mail gerendert (die Mail, die die Suche/Liste in der Preview getroffen hat); die vollständige Konversations-Historie liegt als Metadatum "conversation.messages" bei (message_index wählt eine andere Nachricht, 0 = neueste).',
        inputSchema: {
          type: 'object',
          properties: {
            index: {
              type: 'number',
              description: '0-basierter Index der E-Mail in der aktuellen Liste (Standard: 0)',
              default: 0
            },
            id: {
              type: 'string',
              description: 'Optionale ID oder Conversation-ID der E-Mail'
            },
            message_index: {
              type: 'number',
              description: 'Optional: 0-basierter Index in der Konversations-Historie, neueste zuerst (0 = neueste Thread-Mail = Standard). Nur bei Konversationen wirksam.'
            },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          }
        }
      },
      {
        name: 'outlook_list_attachments',
        description: 'LISTET die Anlagen einer E-Mail im Lesebereich auf (Dateinamen der Anlagen-Chips + Inline-Bilder mit Größen). Non-destruktive Vorstufe zu outlook_save_attachments.',
        inputSchema: {
          type: 'object',
          properties: {
            index: {
              type: 'number',
              description: '0-basierter Index der E-Mail in der aktuellen Liste (Standard: 0)',
              default: 0
            },
            id: {
              type: 'string',
              description: 'Optionale Conversation-ID (data-convid) der E-Mail — hat Vorrang vor index'
            },
            folder: {
              type: 'string',
              description: 'Optionaler Ausgangsordner (inbox/sentitems/archive/drafts), in dem die Conversation-ID gesucht wird'
            },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          }
        }
      },
      {
        name: 'outlook_save_attachments',
        description: 'Speichert die Anlagen der E-Mail auf die Festplatte: Anlagen-Chips per Hover-Download (Playwright-Download-Event), Inline-Bilder (blob/data-URIs) als Fallback. Gibt die gespeicherten Dateipfade zurück. destDir muss unter dem Home-Verzeichnis liegen.',
        inputSchema: {
          type: 'object',
          properties: {
            index: {
              type: 'number',
              description: '0-basierter Index der E-Mail in der aktuellen Liste (Standard: 0)',
              default: 0
            },
            id: {
              type: 'string',
              description: 'Optionale Conversation-ID (data-convid) der E-Mail — hat Vorrang vor index'
            },
            folder: {
              type: 'string',
              description: 'Optionaler Ausgangsordner (inbox/sentitems/archive/drafts), in dem die Conversation-ID gesucht wird'
            },
            destDir: {
              type: 'string',
              description: 'Zielordner (absolut oder ~/...; muss unter dem Home-Verzeichnis liegen). Default: ~/projects/outlook-attachments/<Zeitstempel>/'
            },
            includeInlineImages: {
              type: 'boolean',
              description: 'Inline-Bilder als Fallback extrahieren, wenn kein Chip-Download gelang (Standard: true; verhindert Duplikate, wenn Chips erfolgreich waren)',
              default: true
            },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          }
        }
      },
      {
        name: 'outlook_list_recent',
        description: 'Listet die neuesten E-Mails im Posteingang auf.',
        inputSchema: {
          type: 'object',
          properties: {
            limit: {
              type: 'number',
              description: 'Anzahl der E-Mails (Standard: 10)',
              default: 10
            },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          }
        }
      },
      {
        name: 'outlook_create_draft',
        description: 'Erstellt einen neuen E-Mail-Entwurf im Ordner "Entwürfe" (Drafts) in Outlook Web. Du musst die E-Mail danach nur noch in Outlook öffnen, kurz gegenlesen und absenden.',
        inputSchema: {
          type: 'object',
          properties: {
            to: {
              type: 'string',
              description: 'Empfänger-E-Mail-Adresse (oder mehrere getrennt durch Komma/Semikolon)'
            },
            subject: {
              type: 'string',
              description: 'Betreff der E-Mail'
            },
            body: {
              type: 'string',
              description: 'Inhalt / Text der E-Mail (Absätze werden automatisch übernommen)'
            },
            cc: {
              type: 'string',
              description: 'Optionale CC-Empfänger-Adresse(n)'
            },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          },
          required: ['to', 'subject', 'body']
        }
      },
      {
        name: 'outlook_reply_draft',
        description: 'Erstellt einen nativen Antwort-Entwurf (Reply / Reply All) auf eine vorhandene E-Mail in Outlook Web. Behält den vollständigen E-Mail-Verlauf (Thread/Zitat) und die automatische Outlook-Signatur bei. Speichert den Entwurf im Ordner "Entwürfe" (Drafts).',
        inputSchema: {
          type: 'object',
          properties: {
            email_id: {
              type: 'string',
              description: 'Optionale ID oder Conversation-ID (data-convid) der Ziel-E-Mail.'
            },
            id: {
              type: 'string',
              description: 'Alias für email_id.'
            },
            query: {
              type: 'string',
              description: 'Optionaler Suchbegriff oder Betreff zum Finden der Ziel-E-Mail (falls email_id nicht angegeben).'
            },
            subject: {
              type: 'string',
              description: 'Alias für query (Betreff der zu beantwortenden E-Mail).'
            },
            body: {
              type: 'string',
              description: 'Inhalt / Text der Antwort (wird oben über dem bisherigen Verlauf und vor der Signatur eingefügt).'
            },
            reply_all: {
              type: 'boolean',
              description: 'True für "Allen antworten" (Reply All), False für einfaches "Antworten" (Reply). Standard: false.',
              default: false
            },
            folder: {
              type: 'string',
              description: 'Optional: Ordner, in dem gesucht werden soll (Standard: inbox).'
            },
            expected_from: {
              type: 'string',
              description: 'Sicherheits-Assertion: Erwarteter Absender der Quellmail (Name oder E-Mail). Bei Abweichung wird KEIN Entwurf erstellt (verhindert Reply im falschen Thread).'
            },
            expected_to: {
              type: 'string',
              description: 'Optionale Sicherheits-Assertion: Erwarteter Haupt-Empfänger (To) der Quellmail.'
            },
            expected_subject: {
              type: 'string',
              description: 'Optionale Sicherheits-Assertion: Erwarteter Betreff der Quellmail.'
            },
            expected_date: {
              type: 'string',
              description: 'Optionale Sicherheits-Assertion: Erwartetes Datum der Quellmail ("YYYY-MM-DD" oder "DD.MM.YYYY").'
            },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          },
          required: ['body']
        }
      },
      {
        name: 'outlook_close',
        description: 'Schließt die Hintergrund-Browserinstanz.',
        inputSchema: {
          type: 'object',
          properties: {
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          }
        }
      },
      {
        name: 'outlook_list_folders',
        description: 'Listet die Ordner im Navigationsbereich (Inbox, Sent Items, Archive, Deleted Items, Drafts, Newsletter) mit ihren Namen und Item-Counts.',
        inputSchema: {
          type: 'object',
          properties: {
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          }
        }
      },
      {
        name: 'outlook_navigate_folder',
        description: 'Navigiert in einen Mail-Ordner (inbox|sentitems|archive|deleteditems|drafts|newsletter).',
        inputSchema: {
          type: 'object',
          properties: {
            folder: { type: 'string', description: 'Ordnerpfad (inbox, sentitems, archive, ...)' },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          },
          required: ['folder']
        }
      },
      {
        name: 'outlook_bulk_list',
        description: 'Listet ALLE E-Mails eines Ordners per Scroll-Enumeration (überwindet die Virtualisierung der Nachrichtenliste). Gibt conversation-IDs + Vorschau je Mail zurück.',
        inputSchema: {
          type: 'object',
          properties: {
            folder: { type: 'string', description: 'inbox|sentitems|archive|deleteditems|drafts|newsletter' },
            maxScrolls: { type: 'number', description: 'Maximale Scroll-Iterationen (Standard 140)' },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          },
          required: ['folder']
        }
      },
      {
        name: 'outlook_read_email',
        description: 'Liest eine E-Mail vollständig im Lesebereich (Markdown) anhand ihrer conversation-ID (data-convid aus bulk_list). Scrollt ggf. automatisch durch den Ordner, bis die conversation gefunden ist. Bei Timeout/nicht gefundener ID wird automatisch 1x per Suche recovert (korrigierte conversation-ID aus dem Suchtreffer).',
        inputSchema: {
          type: 'object',
          properties: {
            id: { type: 'string', description: 'conversation-ID (data-convid)' },
            folder: { type: 'string', description: 'Optional: Ordner, aus dem gelesen werden soll (inbox|sentitems|archive|...)' },
            query: { type: 'string', description: 'Optional: Suchbegriff (Betreff/Absender) für den automatischen Recovery-Retry, falls die conversation-ID nicht gefunden wird. Ohne query wird die ID selbst als Suchbegriff verwendet.' },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          },
          required: ['id']
        }
      },
      {
        name: 'outlook_toggle_flag',
        description: 'Toggelt bzw. entfernt das Flag (Kennzeichnung) einer E-Mail anhand ihrer conversation-ID (data-convid).',
        inputSchema: {
          type: 'object',
          properties: {
            id: { type: 'string', description: 'conversation-ID (data-convid)' },
            folder: { type: 'string', description: 'Optional: Ordner (Standard: flagged)' },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          },
          required: ['id']
        }
      },
      {
        name: 'outlook_archive_email',
        description: 'Verschiebt eine E-Mail anhand ihrer conversation-ID (data-convid) ins Archiv.',
        inputSchema: {
          type: 'object',
          properties: {
            id: { type: 'string', description: 'conversation-ID (data-convid)' },
            folder: { type: 'string', description: 'Optional: Ordner (Standard: inbox)' },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          },
          required: ['id']
        }
      },
      {
        name: 'outlook_discard_draft',
        description: 'Verwirft einen E-Mail-Entwurf im Ordner "Entwürfe" (Drafts) in Outlook Web (OWA-"Verwerfen"/Discard per echtem Klick). Nützlich, um veraltete oder doppelte Entwürfe aufzuräumen. Lokalisierung per conversation-ID (data-convid) oder Betreff-Fragment.',
        inputSchema: {
          type: 'object',
          properties: {
            id: { type: 'string', description: 'conversation-ID (data-convid) des Entwurfs' },
            query: { type: 'string', description: 'Betreff-Fragment zum Finden des Entwurfs (falls keine ID bekannt)' },
            folder: { type: 'string', description: 'Optional: Ordner (Standard: drafts/Entwürfe)' },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          }
        }
      },
      {
        name: 'outlook_accept_meeting',
        description: 'Nimmt eine Termineinladung (Meeting Request) in Outlook Web an (RSVP "Annehmen"/Accept). Sendet die Annahme-Rückmeldung an die Gegenseite, trägt den Termin in den Firmenkalender ein und übernimmt den Outlook-typischen Archiv-Flow. Lokalisierung per conversation-ID (data-convid) oder Suchbegriff.',
        inputSchema: {
          type: 'object',
          properties: {
            id: { type: 'string', description: 'conversation-ID (data-convid) der Einladung' },
            email_id: { type: 'string', description: 'Alias für id.' },
            query: { type: 'string', description: 'Optionaler Suchbegriff oder Betreff zum Finden der Einladung (z.B. "CMIS Schnittstelle").' },
            subject: { type: 'string', description: 'Alias für query (Betreff der Einladung).' },
            folder: { type: 'string', description: 'Optional: Ordner, in dem gesucht werden soll (Standard: inbox).' },
            send_response: { type: 'boolean', description: 'True (Standard): Antwortsendung an die Gegenseite. False: ohne Antwortübermittlung annehmen.', default: true },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          }
        }
      },
      {
        name: 'outlook_decline_meeting',
        description: 'Lehnt eine Termineinladung (Meeting Request) in Outlook Web ab (RSVP "Absagen"/Decline). Sendet die Absage-Rückmeldung an die Gegenseite (optional mit Kommentar), trägt den Termin NICHT in den Kalender ein. Lokalisierung per conversation-ID (data-convid) oder Suchbegriff.',
        inputSchema: {
          type: 'object',
          properties: {
            id: { type: 'string', description: 'conversation-ID (data-convid) der Einladung' },
            email_id: { type: 'string', description: 'Alias für id.' },
            query: { type: 'string', description: 'Optionaler Suchbegriff oder Betreff zum Finden der Einladung (z.B. "Feinkonzept").' },
            subject: { type: 'string', description: 'Alias für query (Betreff der Einladung).' },
            folder: { type: 'string', description: 'Optional: Ordner, in dem gesucht werden soll (Standard: inbox).' },
            comment: { type: 'string', description: 'Optionaler Antworttext an den Organisator (wird über "Antwort bearbeiten" in die Absage eingefügt).' },
            send_response: { type: 'boolean', description: 'True (Standard): Antwortsendung an die Gegenseite. False: ohne Antwortübermittlung ablehnen.', default: true },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          }
        }
      },
      {
        name: 'outlook_create_calendar_event',
        description: 'Legt einen Eintrag (Fokuszeit-Blocker) im Outlook-FIRMENKALENDER an: Titel, Start/Ende, optionale Beschreibung. Firmenblöcke gehören laut Kalender-Architektur AUSSCHLIESSLICH hierher (nicht Google). Empfohlen vorher outlook_get_freebusy für den Kollisions-Check. Default transparency=busy.',
        inputSchema: {
          type: 'object',
          properties: {
            subject: { type: 'string', description: 'Titel des Blockers, z.B. "Blocker für Vorbereitung ..."' },
            start: { type: 'string', description: 'Start ISO, z.B. "2026-09-15T09:30:00"' },
            end: { type: 'string', description: 'Ende ISO, z.B. "2026-09-15T11:00:00"' },
            body: { type: 'string', description: 'Optionale Beschreibung/Notizen' },
            is_all_day: { type: 'boolean', description: 'Ganztägig? (Standard: false)', default: false },
            timezone: { type: 'string', description: 'Optional, Default Europe/Berlin', default: 'Europe/Berlin' },
            transparency: { type: 'string', description: 'busy (nicht verfügbar, Standard) oder free', enum: ['busy', 'free'], default: 'busy' },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          },
          required: ['subject', 'start', 'end']
        }
      },
      {
        name: 'outlook_get_freebusy',
        description: 'Fragt die Kalender-Belegung in einem Zeitfenster ab. OHNE attendees: eigener FIRMENKALENDER (Busy-Slots aus der Wochenansicht). MIT attendees (Array von E-Mails): Verfügbarkeit der ANGEGEBENEN TEILNEHMER über das OWA-Gateway (Free/Busy/Tentative/OOF je Teilnehmer + gemeinsame freie Slots) — für kollisionsfreie Terminvorschläge. Empfohlen vor jedem outlook_create_calendar_event.',
        inputSchema: {
          type: 'object',
          properties: {
            start: { type: 'string', description: 'Start ISO, z.B. "2026-09-15" oder "2026-09-15T09:30:00"' },
            end: { type: 'string', description: 'Ende ISO, z.B. "2026-09-15" oder "2026-09-15T18:00:00"' },
            attendees: {
              type: 'array',
              items: { type: 'string' },
              description: 'Optional: E-Mail-Adressen der Teilnehmer, deren Verfügbarkeit geprüft werden soll, z.B. ["silas.beckmann@d-velop.de","mathis.kuenzel@d-velop.de"]. Ohne Angabe: eigener Kalender.',
              maxItems: 20
            },
            interval_minutes: { type: 'integer', description: 'Slot-Granularität in Minuten (nur mit attendees, Default 15, 5-60).', default: 15, minimum: 5, maximum: 60 },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          },
          required: ['start', 'end']
        }
      },
      {
        name: 'outlook_list_calendar_events',
        description: 'READ-ONLY: Listet ALLE Termine (Titel, Datum, Start/Ende, Ganztag, Status) des Outlook-FIRMENKALENDERS in einem Zeitraum. Basis für Termin-Briefings. Es wird nur gelesen/navigiert, nichts verändert.',
        inputSchema: {
          type: 'object',
          properties: {
            start: { type: 'string', description: 'Start ISO, z.B. "2026-09-15" oder "2026-09-15T09:30:00"' },
            end: { type: 'string', description: 'Ende ISO, z.B. "2026-10-04" oder "2026-10-04T18:00:00"' },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          },
          required: ['start', 'end']
        }
      },
      {
        name: 'outlook_get_calendar_event',
        description: 'READ-ONLY: Liest die Detail-Ansicht (Reading-Panel) eines Kalender-Event im Outlook-FIRMENKALENDER anhand seines Titels/Suchbegriffs. Liefert den vollständigen Detail-Text für Termin-Briefings. Es wird nur gelesen, nichts verändert.',
        inputSchema: {
          type: 'object',
          properties: {
            query: { type: 'string', description: 'Suchbegriff/Titel-Fragment des Events, z.B. "MKK Nachbereitung"' },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          },
          required: ['query']
        }
      },
      {
        name: 'outlook_delete_calendar_event',
        description: 'Löscht einen Eintrag aus dem Outlook-FIRMENKALENDER anhand seines Titels/Suchbegriffs (bzw. einer ID). Verwendet die OWA-Kalender-Suche zum Auffinden und entfernt das Event. Nützlich, um versehentlich angelegte Blocker/Test-Einträge zu bereinigen.',
        inputSchema: {
          type: 'object',
          properties: {
            query: { type: 'string', description: 'Suchbegriff/Titel-Fragment des zu löschenden Kalender-Events' },
            id: { type: 'string', description: 'Alternative: Kalender-Event-ID (falls bekannt)' },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          }
        }
      },
      {
        name: 'outlook_update_calendar_event',
        description: 'Aktualisiert einen bestehenden Eintrag im Outlook-FIRMENKALENDER (Titel/Zeit/Beschreibung) anhand seines Titels/Suchbegriffs. Nützlich, um Blocker nachzubearbeiten.',
        inputSchema: {
          type: 'object',
          properties: {
            query: { type: 'string', description: 'Suchbegriff/Titel-Fragment des zu ändernden Kalender-Events' },
            id: { type: 'string', description: 'Alternative: Kalender-Event-ID (falls bekannt)' },
            subject: { type: 'string', description: 'Neuer Titel' },
            start: { type: 'string', description: 'Neuer Start (ISO)' },
            end: { type: 'string', description: 'Neues Ende (ISO)' },
            body: { type: 'string', description: 'Neue Beschreibung' },
            is_all_day: { type: 'boolean', description: 'Ganztägig? (Standard: false)' },
            timezone: { type: 'string', description: 'Optional, Default Europe/Berlin' },
            transparency: { type: 'string', description: 'busy oder free', enum: ['busy', 'free'] },
            realm: { type: 'string', description: 'Optionaler Tenant/Realm (z.B. "adesso.de" oder "d-velop.de").' }
          }
        }
      }
    ]
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  // Realm-Lock (Runde 2, Live-Befund 05.11.2026): Pro Realm läuft EIN Browser
  // mit EINER Page, die sich alle Agenten/Sessions teilen. Ohne Serialisierung
  // interleaven parallele Tool-Calls ihre mehrstufigen UI-Flows (Suche -> Klick
  // -> Read, Navigation -> Header-Read) auf derselben Page — Folge waren
  // Fremdmails bei get_email und transiente Kalender-Snapshots. Jeder Call ist
  // jetzt ATOMAR je Realm; verschiedene Realms bleiben parallel.
  // Der Lock-Key wird VOR dem Setzen gelesen: expliziter args.realm (bereits
  // sanitär geprüft) oder der aktuelle Singleton-Realm — beides konsistente
  // Zustände, da der sperrende Call setRealm + Ablauf zusammen umschließt.
  const explicitRealm = (args && typeof args.realm === 'string' && args.realm.trim() &&
    /^[a-zA-Z0-9.-]{3,}$/.test(args.realm.trim()))
    ? args.realm.trim()
    : null;
  const realmKey = explicitRealm || owaClient.realm || config.defaultRealm;

  return withRealmLock(realmKey, async () => {
  // Multi-Tenant (Problem 2): Jedes Tool kann optional einen 'realm' übergeben,
  // um gezielt zwischen d-velop.de und adesso.de zu wechseln. Ein globaler
  // owaClient-Singleton trägt den aktuellen Realm; setRealm() lässt den
  // Browser-Kontext über ensureContext() ggf. auf das Pro-Realm-Profil wechseln
  // (config.getProfileDirForRealm). Ohne 'realm' bleibt der letzte/Default-Realm.
  if (args && typeof args.realm === 'string' && args.realm.trim()) {
    const realm = args.realm.trim();
    // Sanitärer Guard: kein Pfad-/Steuerzeichen, damit der Realm nur eine
    // plausible Tenancy-Domain steuert (Profilpfad-Sicherheit, vgl. config).
    if (!/^[a-zA-Z0-9.-]{3,}$/.test(realm)) {
      return {
        isError: true,
        content: [{ type: 'text', text: `Ungültiger realm-Wert "${realm}". Bitte eine Domäne wie "d-velop.de" oder "adesso.de" angeben.` }]
      };
    }
    owaClient.setRealm(realm);
  }

  try {
    switch (name) {
      case 'outlook_status': {
        const status = await owaClient.checkStatus();
        return {
          content: [{ type: 'text', text: JSON.stringify(status, null, 2) }]
        };
      }

      case 'outlook_set_realm': {
        // Der Realm wurde oben bereits global gesetzt und validiert.
        const realm = owaClient.realm;
        return {
          content: [
            {
              type: 'text',
              text: `Aktiver Tenant/Realm ist jetzt "${realm}". Browser-Profil: ${config.getProfileDirForRealm(realm)}. ` +
                `Hinweis: Beim ersten Wechsel muss die Mailbox in diesem Profil einmalig angemeldet sein (outlook_login mit realm="${realm}"). ` +
                `Mit outlook_status lässt sich prüfen, ob die aktive Mailbox zum Realm passt.`
            }
          ]
        };
      }

      case 'outlook_login': {
        const realm = args?.realm || config.defaultRealm;
        owaClient.setRealm(realm);
        await browserManager.openLoginWindow(realm);
        return {
          content: [
            {
              type: 'text',
              text: `Ein sichtbares Browserfenster für '${realm}' wurde geöffnet. Bitte melde dich dort einmalig an (inkl. MFA). Die Session wird im Profil dauerhaft gespeichert. Sobald du deinen Posteingang siehst, kannst du 'outlook_status' aufrufen.`
            }
          ]
        };
      }

      case 'outlook_create_draft': {
        const result = await owaClient.createDraft({
          to: args.to,
          subject: args.subject,
          body: args.body,
          cc: args?.cc
        });
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
        };
      }

      case 'outlook_reply_draft': {
        const result = await owaClient.replyDraft({
          email_id: args?.email_id || args?.id,
          query: args?.query || args?.subject,
          body: args?.body,
          reply_all: args?.reply_all ?? args?.replyAll ?? false,
          folder: args?.folder,
          expected_from: args?.expected_from,
          expected_to: args?.expected_to,
          expected_subject: args?.expected_subject,
          expected_date: args?.expected_date
        });
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
        };
      }

      case 'outlook_search': {
        const results = await owaClient.searchWithFallback(args.query, {
          limit: args.limit || 10,
          folder: args.folder,
          to: args.to,
          dateFrom: args.date_from,
          dateTo: args.date_to
        });
        return {
          content: [{ type: 'text', text: JSON.stringify(results, null, 2) }]
        };
      }

      case 'outlook_get_email': {
        const indexOrId = args?.id || (typeof args?.index === 'number' ? args.index : 0);
        const messageIndex = typeof args?.message_index === 'number' ? args.message_index : null;
        const email = await owaClient.getEmailDetails(indexOrId, { messageIndex });
        return {
          content: [{ type: 'text', text: JSON.stringify(email, null, 2) }]
        };
      }

      case 'outlook_list_recent': {
        const recent = await owaClient.listRecent(args?.limit || 10);
        return {
          content: [{ type: 'text', text: JSON.stringify(recent, null, 2) }]
        };
      }

      case 'outlook_close': {
        await browserManager.close();
        return {
          content: [{ type: 'text', text: 'Browserinstanz wurde geschlossen.' }]
        };
      }

      case 'outlook_list_folders': {
        const res = await owaClient.listFolders();
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_navigate_folder': {
        const res = await owaClient.navigateFolder(args.folder);
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_bulk_list': {
        const res = await owaClient.bulkList(args.folder, args.maxScrolls || 140);
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_read_email': {
        let res;
        try {
          res = await owaClient.readEmailByConvId(args.id, args.folder);
        } catch (err) {
          // Automatischer Recovery (genau 1 Retry, kein Loop): Bei Timeout oder nicht
          // gefundener conversation-ID per Suche nach Betreff/ID suchen, die korrigierte
          // conversation-ID aus dem ersten abweichenden Treffer übernehmen und erneut lesen.
          const msg = err?.message || '';
          const recoverable = /timeout/i.test(msg) || /nicht gefunden nach Scroll-Enumeration/i.test(msg);
          if (!recoverable) throw err;

          const searchQuery = args.query || args.id;
          let correctedId = null;
          try {
            const searchRes = await owaClient.searchEmails(searchQuery, 5);
            correctedId = searchRes?.results?.find(r => r.id && r.id !== args.id)?.id || null;
          } catch (searchErr) {
            // Suche selbst fehlgeschlagen -> Originalfehler transparent durchreichen
            throw err;
          }
          if (!correctedId) throw err;

          res = await owaClient.readEmailByConvId(correctedId, args.folder || 'inbox');
          res.recoveredViaSearch = true;
          res.originalId = args.id;
          res.correctedId = correctedId;
        }
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_list_attachments': {
        const res = await owaClient.listAttachments({
          index: typeof args?.index === 'number' ? args.index : 0,
          id: args?.id,
          folder: args?.folder
        });
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_save_attachments': {
        const res = await owaClient.saveAttachments({
          index: typeof args?.index === 'number' ? args.index : 0,
          id: args?.id,
          folder: args?.folder,
          destDir: args?.destDir,
          includeInlineImages: args?.includeInlineImages ?? true
        });
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_toggle_flag': {
        const res = await owaClient.toggleFlag(args.id, args.folder);
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_archive_email': {
        const res = await owaClient.archiveEmail(args.id, args.folder);
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_discard_draft': {
        const res = await owaClient.discardDraft({
          id: args?.id,
          query: args?.query,
          folder: args?.folder
        });
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_accept_meeting': {
        const res = await owaClient.acceptMeeting({
          email_id: args?.email_id || args?.id,
          id: args?.id,
          query: args?.query || args?.subject,
          subject: args?.subject,
          folder: args?.folder,
          send_response: args?.send_response ?? true
        });
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_decline_meeting': {
        const res = await owaClient.declineMeeting({
          email_id: args?.email_id || args?.id,
          id: args?.id,
          query: args?.query || args?.subject,
          subject: args?.subject,
          folder: args?.folder,
          comment: args?.comment,
          send_response: args?.send_response ?? true
        });
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_create_calendar_event': {
        const res = await owaClient.createCalendarEvent({
          subject: args.subject,
          start: args.start,
          end: args.end,
          body: args?.body,
          is_all_day: args?.is_all_day ?? false,
          timezone: args?.timezone || 'Europe/Berlin',
          transparency: args?.transparency || 'busy'
        });
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_get_freebusy': {
        // Zwei Modi: MIT attendees -> Teilnehmer-Verfügbarkeit über das
        // OWA-GraphQL-Gateway (Kollegen-Kalender). OHNE attendees -> bisheriges
        // Verhalten (eigener Kalender, Wochenansicht-DOM).
        const attendeeList = Array.isArray(args?.attendees) ? args.attendees : [];
        const res = attendeeList.length
          ? await owaClient.getAttendeeAvailability({
              attendees: attendeeList,
              start: args.start,
              end: args.end,
              intervalMinutes: args?.interval_minutes ?? 15
            })
          : await owaClient.getFreeBusy({ start: args.start, end: args.end });
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_list_calendar_events': {
        const res = await owaClient.listCalendarEvents({ start: args.start, end: args.end });
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_get_calendar_event': {
        const res = await owaClient.getCalendarEvent({ query: args.query });
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_delete_calendar_event': {
        const res = await owaClient.deleteCalendarEvent({ query: args?.query, id: args?.id });
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      case 'outlook_update_calendar_event': {
        const res = await owaClient.updateCalendarEvent({
          query: args?.query,
          id: args?.id,
          subject: args?.subject,
          start: args?.start,
          end: args?.end,
          body: args?.body,
          is_all_day: args?.is_all_day,
          timezone: args?.timezone || 'Europe/Berlin',
          transparency: args?.transparency
        });
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      default:
        throw new Error(`Unbekanntes Tool: ${name}`);
    }
  } catch (error) {
    // Runde 4: Diagnostik am Error-Objekt muss den Transport überleben —
    // sie wird an die Fehlermeldung angehängt, nicht verworfen.
    let text = `Fehler in ${name}: ${error.message}`;
    if (error.diagnostics) {
      const diagText = typeof error.diagnostics === 'string'
        ? error.diagnostics
        : JSON.stringify(error.diagnostics, null, 2);
      text += `\nDiagnostik:\n${diagText}`;
    }
    return {
      isError: true,
      content: [{ type: 'text', text }]
    };
  }
  }); // withRealmLock — Realm-Call atomar (siehe Kommentar oben)
});

async function run() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

run().catch((err) => {
  console.error('Fatal error in outlook-mcp:', err);
  process.exit(1);
});

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
          properties: {}
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
        description: 'Sucht E-Mails in Outlook Web anhand eines Suchbegriffs (z.B. Kundenname, Ticketnummer, Absender).',
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
            }
          },
          required: ['query']
        }
      },
      {
        name: 'outlook_get_email',
        description: 'Liest eine E-Mail im Lesebereich vollständig aus (Markdown, Teams-Links, Metadaten). Standardmäßig die oberste E-Mail (Index 0).',
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
            }
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
            }
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
            }
          },
          required: ['to', 'subject', 'body']
        }
      },
      {
        name: 'outlook_close',
        description: 'Schließt die Hintergrund-Browserinstanz.',
        inputSchema: {
          type: 'object',
          properties: {}
        }
      },
      {
        name: 'outlook_list_folders',
        description: 'Listet die Ordner im Navigationsbereich (Inbox, Sent Items, Archive, Deleted Items, Drafts, Newsletter) mit ihren Namen und Item-Counts.',
        inputSchema: { type: 'object', properties: {} }
      },
      {
        name: 'outlook_navigate_folder',
        description: 'Navigiert in einen Mail-Ordner (inbox|sentitems|archive|deleteditems|drafts|newsletter).',
        inputSchema: {
          type: 'object',
          properties: { folder: { type: 'string', description: 'Ordnerpfad (inbox, sentitems, archive, ...)' } },
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
            maxScrolls: { type: 'number', description: 'Maximale Scroll-Iterationen (Standard 140)' }
          },
          required: ['folder']
        }
      },
      {
        name: 'outlook_read_email',
        description: 'Liest eine E-Mail vollständig im Lesebereich (Markdown) anhand ihrer conversation-ID (data-convid aus bulk_list). Scrollt ggf. automatisch durch den Ordner, bis die conversation gefunden ist.',
        inputSchema: {
          type: 'object',
          properties: {
            id: { type: 'string', description: 'conversation-ID (data-convid)' },
            folder: { type: 'string', description: 'Optional: Ordner, aus dem gelesen werden soll (inbox|sentitems|archive|...)' }
          },
          required: ['id']
        }
      }
    ]
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      case 'outlook_status': {
        const status = await owaClient.checkStatus();
        return {
          content: [{ type: 'text', text: JSON.stringify(status, null, 2) }]
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

      case 'outlook_search': {
        const results = await owaClient.searchEmails(args.query, args.limit || 10);
        return {
          content: [{ type: 'text', text: JSON.stringify(results, null, 2) }]
        };
      }

      case 'outlook_get_email': {
        const indexOrId = args?.id || (typeof args?.index === 'number' ? args.index : 0);
        const email = await owaClient.getEmailDetails(indexOrId);
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
        const res = await owaClient.readEmailByConvId(args.id, args.folder);
        return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }] };
      }

      default:
        throw new Error(`Unbekanntes Tool: ${name}`);
    }
  } catch (error) {
    return {
      isError: true,
      content: [{ type: 'text', text: `Fehler in ${name}: ${error.message}` }]
    };
  }
});

async function run() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

run().catch((err) => {
  console.error('Fatal error in outlook-mcp:', err);
  process.exit(1);
});

import { test, describe, after } from 'node:test';
import assert from 'node:assert/strict';
import { browserManager } from '../src/browser.js';
import {
  AVAILABILITY_STATUS_MAP,
  ATTENDEE_TOTAL_BUDGET_MS,
  GET_SCHEDULE_QUERY,
  normalizeAttendeeList,
  parseJwtExp,
  isTokenUsable,
  timeZoneOffsetMs,
  zonedTimeToUtcIso,
  buildGetScheduleRequest,
  decodeAvailabilityView,
  parseGetScheduleResponse,
  parseJwtIdentity,
  intersectFreeSegments,
  formatUtcIsoInZone,
  OwaClient
} from '../src/owa.js';

// ===========================================================================
// Regressionstests: outlook_get_freebusy MIT attendees
// (getAttendeeAvailability über das OWA-GraphQL-Gateway, Live-Befund 27.10.2026).
//
// Reine Funktionen werden direkt getestet; die Methode wird gegen einen
// Fake-Page-Kontext gerufen (Token aus gecachetem Zustand, POST über
// context.request). So ist die ECHTE Produktionsmethode abgedeckt — nicht
// nur ein Duplikat der Logik.
// ===========================================================================

// --- normalizeAttendeeList ---------------------------------------------------

describe('normalizeAttendeeList', () => {
  test('extrahiert E-Mails aus freien Textformen und normalisiert lowercase', () => {
    const r = normalizeAttendeeList(['Silas Beckmann <Silas.Beckmann@D-velop.de>', 'mathis.kuenzel@d-velop.de']);
    assert.deepEqual(r.valid, ['silas.beckmann@d-velop.de', 'mathis.kuenzel@d-velop.de']);
    assert.deepEqual(r.invalid, []);
  });

  test('dedupliziert case-insensitive', () => {
    const r = normalizeAttendeeList(['a@x.de', 'A@X.DE', '  a@x.de  ']);
    assert.deepEqual(r.valid, ['a@x.de']);
  });

  test('sammelt Ungültiges in invalid, wirft NICHT', () => {
    const r = normalizeAttendeeList(['keine-mail', 42, null, '', 'x@y.de']);
    assert.deepEqual(r.valid, ['x@y.de']);
    assert.ok(r.invalid.includes('keine-mail'));
    assert.ok(r.invalid.includes('42'));
  });

  test('nicht-Array-Eingabe -> leeres Ergebnis', () => {
    assert.deepEqual(normalizeAttendeeList(undefined), { valid: [], invalid: [] });
    assert.deepEqual(normalizeAttendeeList('a@x.de'), { valid: [], invalid: [] });
  });

  test('Cap bei max Teilnehmern', () => {
    const many = Array.from({ length: 30 }, (_, i) => `p${i}@x.de`);
    const r = normalizeAttendeeList(many, 20);
    assert.equal(r.valid.length, 20);
  });
});

// --- parseJwtExp / isTokenUsable ----------------------------------------------

function makeFakeJwt(payload) {
  const b64url = (obj) => Buffer.from(JSON.stringify(obj)).toString('base64url');
  return `${b64url({ alg: 'RS256', typ: 'JWT' })}.${b64url(payload)}.SIG`;
}

describe('parseJwtExp / isTokenUsable', () => {
  test('liest exp als ms', () => {
    const exp = Math.floor(Date.now() / 1000) + 3600;
    const t = makeFakeJwt({ exp });
    const expMs = parseJwtExp(t);
    assert.ok(Math.abs(expMs - exp * 1000) < 1000);
  });

  test('ungültiger Token -> null', () => {
    assert.equal(parseJwtExp(null), null);
    assert.equal(parseJwtExp('a.b.c-nicht-base64'), null);
    assert.equal(parseJwtExp('nur-ein-string'), null);
  });

  test('isTokenUsable: gültig mit Marge, ablaufend -> false', () => {
    const fresh = makeFakeJwt({ exp: Math.floor((Date.now() + 3600 * 1000) / 1000) });
    const expiring = makeFakeJwt({ exp: Math.floor((Date.now() + 30 * 1000) / 1000) }); // < 2min Marge
    const expired = makeFakeJwt({ exp: Math.floor((Date.now() - 60 * 1000) / 1000) });
    assert.equal(isTokenUsable(fresh), true);
    assert.equal(isTokenUsable(expiring), false);
    assert.equal(isTokenUsable(expired), false);
    assert.equal(isTokenUsable(null), false);
  });

  test('parseJwtIdentity: upn/unique_name nur wenn E-Mail-artig', () => {
    const t = makeFakeJwt({ upn: 'ext-dhal@d-velop.de', unique_name: 'ext-dhal@d-velop.de', name: 'Hallmann, David' });
    const id = parseJwtIdentity(t);
    assert.equal(id.upn, 'ext-dhal@d-velop.de');
    assert.equal(id.name, 'Hallmann, David');
    const noMail = parseJwtIdentity(makeFakeJwt({ upn: 'keine-mail', name: 'X' }));
    assert.equal(noMail.upn, null);
    assert.deepEqual(parseJwtIdentity(null), {});
  });
});

// --- timeZoneOffsetMs / zonedTimeToUtcIso --------------------------------------

describe('zonedTimeToUtcIso', () => {
  test('Europe/Berlin Winterzeit (CET, UTC+1): 09:30 -> 08:30Z', () => {
    const utc = zonedTimeToUtcIso('2026-01-15T09:30:00', 'Europe/Berlin');
    assert.equal(utc, '2026-01-15T08:30:00.000Z');
  });

  test('Europe/Berlin Sommerzeit (CEST, UTC+2): 09:30 -> 07:30Z', () => {
    // 15.09.2026 liegt IN der Sommerzeit (Umstellung erst 25.10.)
    const utc = zonedTimeToUtcIso('2026-09-15T09:30:00', 'Europe/Berlin');
    assert.equal(utc, '2026-09-15T07:30:00.000Z');
  });

  test('Europe/Berlin Winterzeit (CET, UTC+1): 28.10. 09:30 -> 08:30Z', () => {
    // 28.10.2026 liegt NACH der Umstellung am 25.10.
    const utc = zonedTimeToUtcIso('2026-10-28T09:30:00', 'Europe/Berlin');
    assert.equal(utc, '2026-10-28T08:30:00.000Z');
  });

  test('Wert mit Z bleibt UTC-unverändert', () => {
    assert.equal(zonedTimeToUtcIso('2026-10-28T07:30:00Z'), '2026-10-28T07:30:00.000Z');
  });

  test('Offset-Form (+02:00) wird respektiert', () => {
    assert.equal(zonedTimeToUtcIso('2026-10-28T09:30:00+02:00'), '2026-10-28T07:30:00.000Z');
  });

  test('DST-Grenze: Sommerzeit-Tag liefert UTC+2, Winterzeit-Tag UTC+1', () => {
    // 25.10.2026: Umstellung CEST(UTC+2) -> CET(UTC+1) um 03:00
    assert.equal(timeZoneOffsetMs(Date.parse('2026-10-25T00:30:00Z'), 'Europe/Berlin'), 2 * 3600000);
    assert.equal(timeZoneOffsetMs(Date.parse('2026-10-25T01:30:00Z'), 'Europe/Berlin'), 1 * 3600000);
  });

  test('ungültig -> null', () => {
    assert.equal(zonedTimeToUtcIso(''), null);
    assert.equal(zonedTimeToUtcIso('garbage-date'), null);
  });
});

// --- buildGetScheduleRequest ----------------------------------------------------

describe('buildGetScheduleRequest', () => {
  test('baut OWA-identischen Array-Wrapper', () => {
    const body = buildGetScheduleRequest({
      userId: 'david.hallmann@d-velop.de',
      attendeeEmails: ['a@x.de', 'b@x.de'],
      startUtcIso: '2026-10-28T00:00:00.000Z',
      endUtcIso: '2026-10-28T23:59:59.999Z',
      intervalMinutes: 15
    });
    assert.ok(Array.isArray(body));
    assert.equal(body[0].operationName, 'GetSchedule');
    const input = body[0].variables.input;
    assert.equal(input.userId, 'david.hallmann@d-velop.de');
    assert.equal(input.availabilityViewInterval, 15);
    assert.deepEqual(input.schedules, ['a@x.de', 'b@x.de']);
    // dateTime OHNE Z (wie OWA selbst), timeZone UTC
    assert.equal(input.startTime.dateTime, '2026-10-28T00:00:00.000');
    assert.equal(input.endTime.dateTime, '2026-10-28T23:59:59.999');
    assert.equal(input.startTime.timeZone.name, 'UTC');
    assert.ok(body[0].query.includes('getSchedule(request: $input)'));
  });

  test('ohne userId fällt auf ersten Teilnehmer zurück', () => {
    const body = buildGetScheduleRequest({
      attendeeEmails: ['a@x.de'],
      startUtcIso: '2026-10-28T00:00:00Z',
      endUtcIso: '2026-10-28T01:00:00Z'
    });
    assert.equal(body[0].variables.input.userId, 'a@x.de');
  });

  test('GET_SCHEDULE_QUERY enthält Felder, die der Decode braucht', () => {
    for (const f of ['availabilityView', 'scheduleId', 'scheduleItems', 'workingHours', 'error']) {
      assert.ok(GET_SCHEDULE_QUERY.includes(f), `Query fehlt Feld ${f}`);
    }
  });
});

// --- decodeAvailabilityView -------------------------------------------------------

describe('decodeAvailabilityView', () => {
  const W = { startUtcIso: '2026-10-28T08:00:00.000Z', endUtcIso: '2026-10-28T09:00:00.000Z', intervalMinutes: 15 };

  test('4 Slots frei -> ein zusammengeführtes free-Segment', () => {
    const r = decodeAvailabilityView('0000', W);
    assert.equal(r.slotCount, 4);
    assert.equal(r.segments.length, 1);
    assert.equal(r.segments[0].status, 'free');
    assert.equal(r.busySegments.length, 0);
    assert.equal(r.segments[0].startUtcIso, '2026-10-28T08:00:00.000Z');
    assert.equal(r.segments[0].endUtcIso, '2026-10-28T09:00:00.000Z');
  });

  test('free->busy->tentative wird in 3 Segmente gemergt, Status-Map korrekt', () => {
    const r = decodeAvailabilityView('0021', W);
    assert.deepEqual(r.segments.map(s => s.status), ['free', 'busy', 'tentative']);
    assert.equal(r.segments[0].minutes, 30);
    assert.equal(r.segments[1].startUtcIso, '2026-10-28T08:30:00.000Z');
    assert.equal(r.segments[2].startUtcIso, '2026-10-28T08:45:00.000Z');
    assert.equal(r.busySegments.length, 2);
  });

  test('oof (3) zählt zu busySegments', () => {
    const r = decodeAvailabilityView('03', { ...W, intervalMinutes: 30 });
    assert.equal(r.busySegments.length, 1);
    assert.equal(r.busySegments[0].status, 'oof');
  });

  test('View kürzer als Fenster: Segmente enden am View-Ende (Rest fehlt ehrlich)', () => {
    const r = decodeAvailabilityView('02', W); // nur 30 min View für 60 min Fenster
    assert.equal(r.slotCount, 2);
    // Kein Clipping ans Fenster-Ende — die API lieferte nur 2 Slots:
    assert.equal(r.segments[r.segments.length - 1].endUtcIso, '2026-10-28T08:30:00.000Z');
    assert.equal(r.segments[0].minutes, 15);
  });

  test('View länger als Fenster: wird auf Fenster geclippt', () => {
    const r = decodeAvailabilityView('0000000000', W);
    assert.equal(r.segments[0].endUtcIso, W.endUtcIso);
  });

  test('ungültige Eingaben -> leere Segmente (kein Throw)', () => {
    assert.equal(decodeAvailabilityView('', W).segments.length, 0);
    assert.equal(decodeAvailabilityView(null, W).segments.length, 0);
    assert.equal(decodeAvailabilityView('0000', {}).segments.length, 0);
    assert.equal(decodeAvailabilityView('0000', { startUtcIso: 'x', endUtcIso: 'y' }).segments.length, 0);
  });

  test('AVAILABILITY_STATUS_MAP deckt 0-4 ab', () => {
    assert.deepEqual(AVAILABILITY_STATUS_MAP, { 0: 'free', 1: 'tentative', 2: 'busy', 3: 'oof', 4: 'workingElsewhere' });
  });
});

// --- parseGetScheduleResponse -----------------------------------------------------

describe('parseGetScheduleResponse', () => {
  const SCHEDULES = [{ scheduleId: 'a@x.de', availabilityView: '00', error: null, scheduleItems: [], workingHours: null }];

  test('Array-Wrapper (OWA-Format)', () => {
    const text = JSON.stringify([{ data: { getSchedule: { schedules: SCHEDULES } } }]);
    const r = parseGetScheduleResponse(text);
    assert.equal(r.ok, true);
    assert.equal(r.schedules.length, 1);
    assert.equal(r.schedules[0].scheduleId, 'a@x.de');
  });

  test('nacktes Objekt (Fallback)', () => {
    const r = parseGetScheduleResponse(JSON.stringify({ data: { getSchedule: { schedules: SCHEDULES } } }));
    assert.equal(r.ok, true);
  });

  test('GraphQL-Errors -> ok:false mit errorInfo', () => {
    const r = parseGetScheduleResponse(JSON.stringify([{ errors: [{ message: 'Access denied' }] }]));
    assert.equal(r.ok, false);
    assert.equal(r.errorInfo.kind, 'graphql-errors');
    assert.equal(r.errorInfo.errors[0].message, 'Access denied');
  });

  test('Nicht-JSON -> ok:false, kind non-json', () => {
    const r = parseGetScheduleResponse('<html>401</html>');
    assert.equal(r.ok, false);
    assert.equal(r.errorInfo.kind, 'non-json');
  });

  test('fehlende schedules -> ok:false, kind missing-schedules', () => {
    const r = parseGetScheduleResponse(JSON.stringify([{ data: {} }]));
    assert.equal(r.ok, false);
    assert.equal(r.errorInfo.kind, 'missing-schedules');
  });

  test('leere Eingabe -> ok:false (kein Throw)', () => {
    assert.equal(parseGetScheduleResponse('').ok, false);
    assert.equal(parseGetScheduleResponse(null).ok, false);
  });
});

// --- intersectFreeSegments ---------------------------------------------------------

describe('intersectFreeSegments', () => {
  const W = { startUtcIso: '2026-10-28T08:00:00.000Z', endUtcIso: '2026-10-28T10:00:00.000Z' };
  // Fenster 08:00-10:00, 4 Viertelstunden-Paare:
  // A busy 08:30-09:00 | B busy 09:00-09:30  -> gemeinsam frei 09:30-10:00 (+ 08:00-08:30)
  const busyA = [{ startUtcIso: '2026-10-28T08:30:00.000Z', endUtcIso: '2026-10-28T09:00:00.000Z' }];
  const busyB = [{ startUtcIso: '2026-10-28T09:00:00.000Z', endUtcIso: '2026-10-28T09:30:00.000Z' }];

  test('Schnittmenge zweier Teilnehmer', () => {
    const r = intersectFreeSegments([busyA, busyB], { ...W, minDurationMinutes: 15 });
    assert.deepEqual(r, [
      { start: '2026-10-28T08:00:00.000Z', end: '2026-10-28T08:30:00.000Z', durationMinutes: 30 },
      { start: '2026-10-28T09:30:00.000Z', end: '2026-10-28T10:00:00.000Z', durationMinutes: 30 }
    ]);
  });

  test('keine gemeinsame Lücke -> leere Liste', () => {
    const busyC = [
      { startUtcIso: '2026-10-28T08:00:00.000Z', endUtcIso: '2026-10-28T09:00:00.000Z' }
    ];
    const r = intersectFreeSegments([busyA, busyB, busyC], { ...W, minDurationMinutes: 15 });
    // busyC lässt 09:00-10:00 frei; B ist 09:00-09:30 busy -> Schnitt: 09:30-10:00
    assert.deepEqual(r, [{ start: '2026-10-28T09:30:00.000Z', end: '2026-10-28T10:00:00.000Z', durationMinutes: 30 }]);
  });

  test('minDuration filtert kurze Reste', () => {
    const busyD = [{ startUtcIso: '2026-10-28T08:00:00.000Z', endUtcIso: '2026-10-28T08:15:00.000Z' }];
    const r = intersectFreeSegments([busyA, busyD], { ...W, minDurationMinutes: 20 });
    // frei: 08:15-08:30 (15min, gefiltert), 09:00-10:00
    assert.deepEqual(r, [{ start: '2026-10-28T09:00:00.000Z', end: '2026-10-28T10:00:00.000Z', durationMinutes: 60 }]);
  });

  test('überlappende Busy-Intervalle desselben Teilnehmers werden gemergt', () => {
    const busyOverlap = [
      { startUtcIso: '2026-10-28T08:00:00.000Z', endUtcIso: '2026-10-28T08:45:00.000Z' },
      { startUtcIso: '2026-10-28T08:30:00.000Z', endUtcIso: '2026-10-28T09:30:00.000Z' }
    ];
    const r = intersectFreeSegments([busyOverlap], { ...W, minDurationMinutes: 15 });
    assert.deepEqual(r, [{ start: '2026-10-28T09:30:00.000Z', end: '2026-10-28T10:00:00.000Z', durationMinutes: 30 }]);
  });

  test('ohne Busy-Listen: ganzes Fenster frei', () => {
    const r = intersectFreeSegments([[], []], { ...W, minDurationMinutes: 15 });
    assert.deepEqual(r, [{ start: W.startUtcIso, end: W.endUtcIso, durationMinutes: 120 }]);
  });
});

// --- formatUtcIsoInZone ------------------------------------------------------------

describe('formatUtcIsoInZone', () => {
  test('formatiert de-DE Lokalzeit', () => {
    // 28.10.2026 = CET (UTC+1) nach der Umstellung am 25.10.
    assert.equal(formatUtcIsoInZone('2026-10-28T07:30:00.000Z', 'Europe/Berlin'), '28.10.2026, 08:30');
  });
  test('ungültig -> null', () => {
    assert.equal(formatUtcIsoInZone('nope'), null);
  });
});

// --- getAttendeeAvailability (echte Methode, Fake-Kontext) ---------------------------

describe('OwaClient.getAttendeeAvailability (Fake-Kontext)', () => {
  const TOKEN = makeFakeJwt({ exp: Math.floor((Date.now() + 3600 * 1000) / 1000) });
  const GRAPHQL_URL = 'https://outlook.cloud.microsoft/outlookgatewayb2/graphql';
  const GOOD_RESPONSE = JSON.stringify([{
    data: {
      getSchedule: {
        schedules: [
          {
            scheduleId: 'silas.beckmann@d-velop.de',
            availabilityView: '0022000000',
            error: null,
            scheduleItems: [{ status: 'Busy', startTime: { dateTime: '2026-10-28T08:30:00Z' }, endTime: { dateTime: '2026-10-28T09:00:00Z' }, subject: 'Daily', isPrivate: false }],
            workingHours: { daysOfWeek: ['Monday'], startTime: '08:00:00', endTime: '17:00:00', timeZone: { name: 'W. Europe Standard Time', bias: 0 } }
          },
          {
            scheduleId: 'extern@other.com',
            availabilityView: '',
            error: { message: 'Recipient not found', responseCode: 'ErrorMailRecipientNotFound' },
            scheduleItems: [],
            workingHours: null
          }
        ]
      }
    }
  }]);

  // browserManager.ensureContext patchen (wie in getfreebusy-result.test.mjs):
  // die echte Methode holt darüber die Seite; hier kommt der Fake zurück.
  const origEnsure = browserManager.ensureContext;
  after(() => { browserManager.ensureContext = origEnsure; });

  function fakePageContext({ posts = [], statusFactory = null } = {}) {
    const request = {
      post: async (url, opts) => {
        posts.push({ url, opts });
        if (statusFactory) return statusFactory(posts.length);
        return { status: () => 200, text: async () => GOOD_RESPONSE };
      }
    };
    const page = { context: () => ({ request }) };
    return { page, posts };
  }

  function clientWithCache() {
    const client = new OwaClient('d-velop.de');
    client._scheduleToken = TOKEN;
    client._scheduleTokenExpMs = parseJwtExp(TOKEN);
    client._scheduleGraphqlUrl = GRAPHQL_URL;
    client._scheduleUserId = 'david.hallmann@d-velop.de';
    return client;
  }

  test('Erfolgsfall: Decodierung, Teilnehmer-Fehler isoliert, POST-Body korrekt', async () => {
    const posts = [];
    const { page } = fakePageContext({ posts });
    browserManager.ensureContext = async () => ({ page });
    const client = clientWithCache();
    const res = await client.getAttendeeAvailability({
      attendees: ['silas.beckmann@d-velop.de', 'extern@other.com'],
      start: '2026-10-28T08:00:00',
      end: '2026-10-28T08:50:00',
      intervalMinutes: 15
    });
    assert.equal(res.status, 'ok');
    assert.equal(res.attendees.length, 2);

    const silas = res.attendees[0];
    assert.equal(silas.email, 'silas.beckmann@d-velop.de');
    assert.equal(silas.error, null);
    // Fenster: 08:00-08:50 Berlin = 07:00-07:50Z. View '0022000000' (15-min-Slots):
    // 07:00-07:30 frei, 07:30-08:00 busy (2 Slots gemergt) -> auf 07:50Z geclippt.
    assert.equal(silas.busySlots.length, 1);
    assert.equal(silas.busySlots[0].status, 'busy');
    assert.equal(silas.busySlots[0].start, '2026-10-28T07:30:00.000Z');
    assert.equal(silas.busySlots[0].end, '2026-10-28T07:50:00.000Z');
    assert.ok(silas.busySlots[0].startLocal.includes('08:30'));
    assert.equal(silas.scheduleItems.length, 1);
    assert.ok(res.window.startUtc.endsWith('Z'));

    const ext = res.attendees[1];
    assert.ok(ext.error);
    assert.equal(ext.error.responseCode, 'ErrorMailRecipientNotFound');
    assert.equal(ext.busySlots.length, 0);

    // POST-Verifikation: exakt 1 Call, korrekter Body
    assert.equal(posts.length, 1);
    assert.equal(posts[0].url, GRAPHQL_URL);
    const sent = posts[0].opts.data[0];
    assert.equal(sent.variables.input.userId, 'david.hallmann@d-velop.de');
    assert.deepEqual(sent.variables.input.schedules, ['silas.beckmann@d-velop.de', 'extern@other.com']);
    assert.ok(posts[0].opts.headers['Authorization'].startsWith('Bearer '));
    assert.equal(res.diagnostics.httpStatus, 200);
    assert.equal(res.note.includes('WARNUNG'), true, 'note muss Teilnehmer-Fehler erwähnen');
  });

  test('401 -> Token-Refresh (force capture) -> Retry erfolgreich', async () => {
    const posts = [];
    let call = 0;
    const { page } = fakePageContext({
      posts,
      statusFactory: () => {
        call++;
        if (call === 1) return { status: () => 401, text: async () => '' };
        return { status: () => 200, text: async () => GOOD_RESPONSE };
      }
    });
    browserManager.ensureContext = async () => ({ page });

    const client = clientWithCache();
    // Cached Token "abgelaufen": der Boot-Capture ist Live-verifiziert
    // (Probe 5/6) und im Unit-Test nicht abbildbar — hier wird der Capture
    // deterministisch simuliert (Erst-Call: boot, nach 401: force-refresh).
    // Patch-Signatur: (page, options) — force kommt aus dem ZWEITEN Argument.
    client.captureScheduleToken = async (pageArg, { force } = {}) => {
      client._scheduleToken = TOKEN;
      client._scheduleTokenExpMs = parseJwtExp(TOKEN);
      client._scheduleGraphqlUrl = GRAPHQL_URL;
      return { token: TOKEN, userId: 'david.hallmann@d-velop.de', url: GRAPHQL_URL, source: force ? 'refresh' : 'boot' };
    };
    // Erster Token-Cache unbrauchbar machen, damit überhaupt gecaptet wird:
    client._scheduleToken = makeFakeJwt({ exp: Math.floor((Date.now() - 5000) / 1000) });
    client._scheduleTokenExpMs = parseJwtExp(client._scheduleToken);

    const res = await client.getAttendeeAvailability({
      attendees: ['silas.beckmann@d-velop.de'],
      start: '2026-10-28T08:00:00',
      end: '2026-10-28T09:00:00'
    });
    assert.equal(res.status, 'ok', 'Retry nach Token-Refresh muss Erfolg liefern');
    assert.equal(posts.length, 2, 'Erst-Call 401 + Retry');
    assert.equal(res.diagnostics.httpStatus, 200);
    assert.equal(res.diagnostics.tokenSource, 'refresh');
  });

  test('Gateway 500 -> strukturiertes errorResult (kein Throw)', async () => {
    const { page } = fakePageContext({
      statusFactory: () => ({ status: () => 500, text: async () => 'boom' })
    });
    browserManager.ensureContext = async () => ({ page });
    const client = clientWithCache();
    const res = await client.getAttendeeAvailability({
      attendees: ['a@x.de'], start: '2026-10-28T08:00:00', end: '2026-10-28T09:00:00'
    });
    assert.equal(res.status, 'error');
    assert.ok(res.error.includes('500'));
    assert.ok(res.diagnostics.usedMs < ATTENDEE_TOTAL_BUDGET_MS + 1000);
  });

  test('GraphQL-Errors in 200-Antwort -> errorResult mit errorInfo', async () => {
    const { page } = fakePageContext({
      statusFactory: () => ({ status: () => 200, text: async () => JSON.stringify([{ errors: [{ message: 'denied' }] }]) })
    });
    browserManager.ensureContext = async () => ({ page });
    const client = clientWithCache();
    const res = await client.getAttendeeAvailability({
      attendees: ['a@x.de'], start: '2026-10-28T08:00:00', end: '2026-10-28T09:00:00'
    });
    assert.equal(res.status, 'error');
    assert.equal(res.diagnostics.errorInfo.kind, 'graphql-errors');
  });

  test('POST-Exception -> errorResult (kein Throw)', async () => {
    const request = { post: async () => { throw new Error('network down'); } };
    const page = { context: () => ({ request }) };
    browserManager.ensureContext = async () => ({ page });
    const client = clientWithCache();
    const res = await client.getAttendeeAvailability({
      attendees: ['a@x.de'], start: '2026-10-28T08:00:00', end: '2026-10-28T09:00:00'
    });
    assert.equal(res.status, 'error');
    assert.ok(res.error.includes('network down'));
  });

  test('Parametervalidierung wirft (wie getFreeBusy)', async () => {
    const client = clientWithCache();
    await assert.rejects(() => client.getAttendeeAvailability({ attendees: [], start: '2026-10-28', end: '2026-10-28' }), /attendees/);
    await assert.rejects(() => client.getAttendeeAvailability({ attendees: ['keine-mail'], start: '2026-10-28', end: '2026-10-28' }), /gültige E-Mail/);
    await assert.rejects(() => client.getAttendeeAvailability({ attendees: ['a@x.de'], start: '', end: '2026-10-28' }), /start/);
    await assert.rejects(() => client.getAttendeeAvailability({ attendees: ['a@x.de'], start: '2026-10-28T10:00:00', end: '2026-10-28T09:00:00' }), /nach "start"/);
  });
});

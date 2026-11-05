// R3 Regressionsprobe (READ-ONLY): KW46, KW47 (2x Stabilitaet), aktuelle Woche,
// get_email-Determinismus. Kein Versand, keine Mutation.
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';

const API_KEY = process.env.OUTLOOK_MCP_API_KEY || '';
const transport = new StreamableHTTPClientTransport(new URL('http://127.0.0.1:3220/mcp'), {
  requestInit: { headers: { 'X-API-Key': API_KEY } }
});
const client = new Client({ name: 'r3-regression-probe', version: '1.0.0' }, { capabilities: {} });
const parse = (res) => JSON.parse(res.content[0].text);

const week = async (start, end) => {
  const data = parse(await client.callTool({
    name: 'outlook_list_calendar_events',
    arguments: { start, end }
  }));
  return {
    count: data.count,
    verified: data.weekVerified,
    dataQuality: data.dataQuality,
    settlement: data.eventSettlement,
    titles: (data.events || []).map((e) => `${e.date} ${e.start} ${e.title}`),
    note: (data.note || '').includes('WARNUNG') ? 'NOTE-WARNUNG: ' + (data.note || '') : null
  };
};

try {
  await client.connect(transport);

  const kw46a = await week('2026-11-09', '2026-11-15');
  const kw46b = await week('2026-11-09', '2026-11-15');
  console.log('=== KW46 (09.-15.11.) Lauf A ===');
  console.log(JSON.stringify({ count: kw46a.count, verified: kw46a.verified, settlement: kw46a.settlement, warn: kw46a.note }, null, 2));
  console.log('KW46 A titles:', JSON.stringify(kw46a.titles, null, 1));
  console.log('KW46 deterministisch A==B:', JSON.stringify(kw46a.titles) === JSON.stringify(kw46b.titles), `(B count=${kw46b.count})`);

  const kw47a = await week('2026-11-16', '2026-11-22');
  const kw47b = await week('2026-11-16', '2026-11-22');
  console.log('=== KW47 (16.-22.11.) A/B ===');
  console.log(JSON.stringify({ countA: kw47a.count, countB: kw47b.count, settlementA: kw47a.settlement, settlementB: kw47b.settlement, warnA: kw47a.note, warnB: kw47b.note }, null, 2));
  console.log('KW47 A titles:', JSON.stringify(kw47a.titles, null, 1));

  const cur = await week('2026-11-02', '2026-11-08');
  console.log('=== Aktuelle Woche (02.-08.11.) ===');
  console.log(JSON.stringify({ count: cur.count, verified: cur.verified, settlement: cur.settlement, warn: cur.note }, null, 2));
  console.log('Aktuelle Woche titles:', JSON.stringify(cur.titles, null, 1));

  // get_email-Determinismus: Mail-ID holen und 2x rendern (nur Metadaten vergleichen)
  const recent = parse(await client.callTool({ name: 'outlook_list_recent', arguments: { limit: 3 } }));
  const first = (recent.emails || recent.messages || recent.items || [])[0];
  if (first?.id) {
    const m1 = parse(await client.callTool({ name: 'outlook_read_email', arguments: { id: first.id } }));
    const m2 = parse(await client.callTool({ name: 'outlook_read_email', arguments: { id: first.id } }));
    const e1 = m1.email || m1;
    const e2 = m2.email || m2;
    console.log('=== get_email-Determinismus ===');
    console.log(JSON.stringify({
      id: first.id,
      deterministic: (e1.subject === e2.subject) && (e1.from === e2.from || e1.from?.address === e2.from?.address),
      subject: e1.subject
    }, null, 2));
  } else {
    console.log('get_email-Determinismus: keine Mail gefunden (skipped)');
  }

  await client.close();
} catch (e) {
  console.error('Fehler:', e.message || e);
  process.exit(1);
}

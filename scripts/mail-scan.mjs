/**
 * Partner email activity, read from the mailbox and filed against tracker rows.
 *
 *   node scripts/mail-scan.mjs [--days 30]
 *   -> data/mail-activity.json     per partner: last inbound, last outbound, waiting?
 *   -> data/mail-unmapped.json     domains seen that nobody has classified yet
 *
 * What leaves the mailbox: a partner name, two dates and two counts. No subject, no
 * body, no address, no person's name. data/mail-activity.json is gitignored anyway,
 * because this repo is public and "we have been chasing Agoda for 12 days" is not.
 *
 * Partners are resolved from data/partner-domains.json only. Letting the matcher guess
 * from the name filed read.ai under Zentrumhub and trip.com under Tripjack, and a wrong
 * attribution is worse than a gap - it tells someone a partner replied when they did not.
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const REPORT = path.join(root, '..', '..', 'REPORT');
const require = createRequire(path.join(REPORT, 'package.json'));
const { PublicClientApplication, LogLevel } = require('@azure/msal-node');
const axios = require('axios');

const CFG = JSON.parse(fs.readFileSync(path.join(REPORT, 'scripts/auto/graph-config.json'), 'utf8'));
const CACHE = path.join(REPORT, 'auth/msal-cache.json');

const argv = process.argv.slice(2);
const DAYS = Number(argv[argv.indexOf('--days') + 1]) || 30;

/* ------------------------------------------------------------------ auth */
const cachePlugin = {
  beforeCacheAccess: async (c) => { if (fs.existsSync(CACHE)) c.tokenCache.deserialize(fs.readFileSync(CACHE, 'utf8')); },
  afterCacheAccess: async (c) => { if (c.cacheHasChanged) fs.writeFileSync(CACHE, c.tokenCache.serialize()); },
};
const pca = new PublicClientApplication({
  auth: { clientId: CFG.clientId, authority: `https://login.microsoftonline.com/${CFG.tenantId}` },
  cache: { cachePlugin },
  system: { loggerOptions: { loggerCallback: () => {}, logLevel: LogLevel.Error } },
});

async function token() {
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (!accounts.length) throw new Error('No cached account. Run REPORT/scripts/auto/save-graph-token.js first.');
  const { accessToken } = await pca.acquireTokenSilent({ account: accounts[0], scopes: ['Mail.Read', 'User.Read'] });
  return accessToken;
}

/* ------------------------------------------------------------------ inputs */
const trackerPath = path.join(root, 'data', 'tracker.json');
if (!fs.existsSync(trackerPath)) {
  console.error('\nNo data/tracker.json. Run "npm run build" first.\n');
  process.exit(1);
}
const tracker = JSON.parse(fs.readFileSync(trackerPath, 'utf8'));
const known = new Set(tracker.rows.map((r) => r.project));

const mapPath = path.join(root, 'data', 'partner-domains.json');
const cfg = JSON.parse(fs.readFileSync(mapPath, 'utf8'));
const DOMAINS = new Map(Object.entries(cfg.map || {}).map(([d, p]) => [d.toLowerCase(), p]));
const ADDRESSES = new Map(Object.entries(cfg.addresses || {}).map(([a, p]) => [a.toLowerCase(), p]));
const IGNORED = new Set((cfg.ignore || []).map((d) => d.toLowerCase()));
const REVIEW = new Set(Object.keys(cfg.review || {}).map((d) => d.toLowerCase()));

/** Partner names in the map that no longer exist in the tracker, so the map can be tidied. */
const stale = [...new Set([...DOMAINS.values(), ...ADDRESSES.values()])].filter((p) => !known.has(p));

const domainOf = (addr) => String(addr || '').toLowerCase().split('@')[1] || '';
const isIgnored = (d) => [...IGNORED].some((i) => d === i || d.endsWith(`.${i}`));

function partnerFor(address) {
  const a = String(address || '').toLowerCase();
  if (ADDRESSES.has(a)) return ADDRESSES.get(a);
  const d = domainOf(a);
  if (!d) return null;
  if (DOMAINS.has(d)) return DOMAINS.get(d);
  // mail.partner.com and apacmail.partner.com are the same company as partner.com.
  for (const [dom, p] of DOMAINS) if (d.endsWith(`.${dom}`)) return p;
  return null;
}

/* ------------------------------------------------------------------ scan */
async function pageThrough(url, H, onItem, cap = 20) {
  let pages = 0;
  while (url && pages < cap) {
    const { data } = await axios.get(url, H);
    for (const m of data.value || []) onItem(m);
    url = data['@odata.nextLink'];
    pages += 1;
  }
}

const activity = new Map();
const unmapped = new Map();

/**
 * Who we were writing to, or who wrote. The Chase Log asks for this and it is a fact
 * the mailbox holds, so it is carried through - unlike the stage or the gist of the
 * thread, which are judgements. Kept against the latest message in each direction.
 */
function note(partner, dir, iso, who) {
  const v = activity.get(partner)
    || { lastIn: null, lastOut: null, in: 0, out: 0, lastInWho: null, lastOutWho: null, days: new Map() };
  const day = iso.slice(0, 10);
  if (dir === 'in') {
    v.in += 1;
    if (!v.lastIn || day >= v.lastIn) { v.lastIn = day; if (who) v.lastInWho = who; }
  } else {
    v.out += 1;
    if (!v.lastOut || day >= v.lastOut) { v.lastOut = day; if (who) v.lastOutWho = who; }
  }
  // Every day with traffic, not just the latest. Keeping only the last outbound meant
  // the Chase Log got one row per partner however many times they were chased, and a
  // day the job did not run was lost for good.
  const d = v.days.get(day) || { in: 0, out: 0, who: null };
  d[dir] += 1;
  if (dir === 'out' && who && !d.who) d.who = who;
  if (dir === 'in' && who) d.who = who;
  v.days.set(day, d);
  activity.set(partner, v);
}

/** "Qi Li <qi.li@klook.com>" -> "Qi Li". Falls back to the address when unnamed. */
const personOf = (entry) => {
  const name = String(entry?.name || '').trim();
  const addr = String(entry?.address || '').trim();
  if (name && name.toLowerCase() !== addr.toLowerCase()) return name;
  return addr.split('@')[0] || addr;
};

const run = async () => {
  const accessToken = await token();
  const H = { headers: { Authorization: `Bearer ${accessToken}` } };
  const since = new Date(Date.now() - DAYS * 86400000).toISOString();

  const folders = [
    ['inbox', 'in', 'receivedDateTime'],
    ['sentitems', 'out', 'sentDateTime'],
  ];

  for (const [folder, dir, stampField] of folders) {
    const url = `https://graph.microsoft.com/v1.0/me/mailFolders/${folder}/messages`
      + `?$select=receivedDateTime,sentDateTime,from,toRecipients,ccRecipients`
      + `&$filter=${stampField} ge ${since}&$top=100`;

    await pageThrough(url, H, (m) => {
      const when = m[stampField] || m.receivedDateTime || m.sentDateTime;
      if (!when) return;
      const people = dir === 'in'
        ? [m.from?.emailAddress]
        : [...(m.toRecipients || []), ...(m.ccRecipients || [])].map((t) => t.emailAddress);

      // One partner can appear on several addresses in a thread; the first external
      // one that resolves names the counterpart for that message.
      const hit = new Map();
      for (const entry of people.filter((e) => e?.address)) {
        const d = domainOf(entry.address);
        if (!d || isIgnored(d)) continue;
        const partner = partnerFor(entry.address);
        if (partner) { if (!hit.has(partner)) hit.set(partner, personOf(entry)); continue; }
        const v = unmapped.get(d) || { seen: 0, review: REVIEW.has(d) };
        v.seen += 1;
        unmapped.set(d, v);
      }
      for (const [p, who] of hit) note(p, dir, when, who);
    });
  }

  /* ---------------------------------------------------------- shape output */
  const today = new Date().toISOString().slice(0, 10);
  const daysBetween = (from) => Math.round((Date.parse(`${today}T00:00:00Z`) - Date.parse(`${from}T00:00:00Z`)) / 86400000);

  const out = {};
  for (const [partner, v] of [...activity.entries()].sort(([a], [b]) => a.localeCompare(b))) {
    // Waiting means our last message is newer than theirs, or they have never written.
    const waiting = Boolean(v.lastOut) && (!v.lastIn || v.lastOut > v.lastIn);
    out[partner] = {
      lastIn: v.lastIn,
      lastOut: v.lastOut,
      inCount: v.in,
      outCount: v.out,
      waitingReply: waiting,
      waitingDays: waiting ? daysBetween(v.lastOut) : null,
      // Who we wrote to, and who last wrote back. Local file only - this never reaches
      // the published board.
      lastOutWho: v.lastOutWho,
      lastInWho: v.lastInWho,
      // One entry per day with traffic, so the Chase Log can carry a row per chase
      // rather than a row per partner.
      days: [...v.days.entries()]
        .sort(([a], [b]) => a.localeCompare(b))
        .map(([date, c]) => ({ date, in: c.in, out: c.out, who: c.who })),
    };
  }

  fs.writeFileSync(
    path.join(root, 'data', 'mail-activity.json'),
    `${JSON.stringify({ generatedAt: new Date().toISOString(), windowDays: DAYS, partners: out }, null, 1)}\n`,
  );

  const pending = [...unmapped.entries()]
    .filter(([, v]) => v.seen >= 2)
    .sort((a, b) => b[1].seen - a[1].seen)
    .map(([domain, v]) => ({ domain, seen: v.seen, flaggedForReview: v.review }));
  fs.writeFileSync(
    path.join(root, 'data', 'mail-unmapped.json'),
    `${JSON.stringify({ generatedAt: new Date().toISOString(), domains: pending }, null, 1)}\n`,
  );

  /* ---------------------------------------------------------- report */
  const matched = Object.keys(out).length;
  const waiting = Object.values(out).filter((v) => v.waitingReply);
  console.log(`\n  ${DAYS}일 · 매칭된 파트너 ${matched}곳 · 회신 대기 ${waiting.length}곳 · 미분류 도메인 ${pending.length}개\n`);

  for (const [p, v] of Object.entries(out).sort((a, b) => (b[1].waitingDays ?? -1) - (a[1].waitingDays ?? -1)).slice(0, 12)) {
    const w = v.waitingReply ? `회신 대기 ${v.waitingDays}일` : '회신 받음';
    console.log(`  ${p.slice(0, 22).padEnd(24)} 받음 ${String(v.inCount).padStart(2)} 보냄 ${String(v.outCount).padStart(2)}   ${w}`);
  }

  if (pending.length) {
    console.log('\n  미분류 (2회 이상 등장) — data/partner-domains.json 의 map 또는 ignore 로 옮기세요:');
    for (const d of pending.slice(0, 12)) {
      console.log(`    ${d.domain.padEnd(30)} ${String(d.seen).padStart(3)}회${d.flaggedForReview ? '   (review 목록에 있음)' : ''}`);
    }
  }
  if (stale.length) console.log(`\n  지도에 있으나 트래커에 없는 파트너: ${stale.join(', ')}`);
  console.log('');
};

run().catch((e) => {
  const msg = e.response?.data?.error?.message || e.message;
  console.error(`\n  메일 스캔 실패: ${msg}\n`);
  process.exit(1);
});

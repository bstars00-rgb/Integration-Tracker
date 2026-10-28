/**
 * Append email activity to the Chase Log sheet in the SharePoint workbook.
 *
 *   node scripts/chase-write.mjs            보낼 행만 출력, 쓰지 않음
 *   node scripts/chase-write.mjs --write    실제로 기록
 *
 * Reads data/mail-activity.json (written by mail-scan.mjs) and adds one row per
 * outbound chase that is not already logged. Existing rows are read first and matched
 * on project + date, so running twice in a day adds nothing the second time.
 *
 * Only facts the mailbox can prove are written: who, when, by email, and whether a
 * reply came back. Stage, counterpart and the content column are left for a person -
 * deciding that "let us start testing next week" moves a milestone is a judgement, and
 * a wrong one ends up in the weekly report as a go-live that never happened.
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
const SHEET = 'Chase Log';
const WRITE = process.argv.includes('--write');

/** Chase Log columns, A..K. Blank means "a person fills this in". */
const COLS = ['No.', 'Project', 'Category', 'Stage', 'Chase date', 'Method', 'Counterpart', 'Content', 'Responded?', 'Next chase', 'Owner'];
const DAYS_TO_NEXT_CHASE = 7;

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

const encodeShare = (u) => `u!${Buffer.from(u).toString('base64').replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-')}`;
const addDays = (iso, n) => new Date(Date.parse(`${iso}T00:00:00Z`) + n * 86400000).toISOString().slice(0, 10);

// The sheet stores dates as Excel serials (46300 = 5 Oct 2026). Writing ISO strings
// would leave a column half numbers and half text, and the dedupe would never match.
const EPOCH = Date.UTC(1899, 11, 30);
const toSerial = (iso) => (iso ? Math.round((Date.parse(`${iso}T00:00:00Z`) - EPOCH) / 86400000) : '');
const fromSerial = (v) => {
  const n = Number(v);
  if (!Number.isFinite(n) || n < 20000) return String(v ?? '').trim().slice(0, 10);
  return new Date(EPOCH + n * 86400000).toISOString().slice(0, 10);
};

/* ------------------------------------------------------------------ inputs */
const mailPath = path.join(root, 'data', 'mail-activity.json');
if (!fs.existsSync(mailPath)) {
  console.error('\n  data/mail-activity.json 없음. 먼저 npm run mail 을 실행하세요.\n');
  process.exit(1);
}
const mail = JSON.parse(fs.readFileSync(mailPath, 'utf8'));
const tracker = JSON.parse(fs.readFileSync(path.join(root, 'data', 'tracker.json'), 'utf8'));

/** Category and owner come from the tracker so the log row matches the sheet. */
const meta = new Map();
for (const r of tracker.rows) {
  if (meta.has(r.project)) continue;
  // The log files rows under the original v2 category, not the v3 sheet name.
  meta.set(r.project, { category: r.categoryOrig || r.category, owner: r.pic || '', stage: r.currentStage || '' });
}

const run = async () => {
  const accounts = await pca.getTokenCache().getAllAccounts();
  const { accessToken } = await pca.acquireTokenSilent({
    account: accounts[0],
    scopes: ['Files.ReadWrite', 'Sites.Read.All', 'User.Read'],
  });
  const H = { headers: { Authorization: `Bearer ${accessToken}` } };

  const { data: item } = await axios.get(`https://graph.microsoft.com/v1.0/shares/${encodeShare(CFG.shareUrl)}/driveItem`, H);
  const wb = `https://graph.microsoft.com/v1.0/drives/${item.parentReference.driveId}/items/${item.id}/workbook`;

  const { data: used } = await axios.get(`${wb}/worksheets('${SHEET}')/usedRange(valuesOnly=true)?$select=address,values,rowCount`, H);
  const rows = used.values || [];

  // Header is row 3 in this sheet; data starts at 4. Found rather than assumed, so a
  // sheet that grows a banner row does not shift every write by one.
  const headerIdx = rows.findIndex((r) => String(r[0] ?? '').trim() === 'No.');
  const firstDataRow = (headerIdx >= 0 ? headerIdx : 2) + 1;

  // Already logged, so a second run today is a no-op.
  const logged = new Set();
  let lastNo = 0;
  for (let i = firstDataRow; i < rows.length; i += 1) {
    const r = rows[i] || [];
    const project = String(r[1] ?? '').trim();
    if (!project) continue;
    const date = fromSerial(r[4]);
    logged.add(`${project}::${date}`);
    const n = Number(r[0]);
    if (Number.isFinite(n)) lastNo = Math.max(lastNo, n);
  }

  const pending = [];
  for (const [project, m] of Object.entries(mail.partners || {})) {
    if (!m.lastOut) continue;
    if (logged.has(`${project}::${m.lastOut}`)) continue;
    const info = meta.get(project) || { category: '', owner: '', stage: '' };
    pending.push({
      project,
      category: info.category,
      chaseDate: m.lastOut,
      responded: m.waitingReply ? 'N' : 'Y',
      // The person on the other end of the chase, from the mailbox. A reply names the
      // sender; otherwise it is whoever we wrote to.
      counterpart: m.lastInWho || m.lastOutWho || '',
      nextChase: m.waitingReply ? addDays(m.lastOut, DAYS_TO_NEXT_CHASE) : '',
      owner: info.owner,
      stage: info.stage,
    });
  }
  pending.sort((a, b) => a.chaseDate.localeCompare(b.chaseDate) || a.project.localeCompare(b.project));

  console.log(`\n  Chase Log ${rows.length}행 · 기록된 조합 ${logged.size}건 · 추가 대상 ${pending.length}건\n`);
  for (const p of pending) {
    console.log(`  ${p.chaseDate}  ${p.project.slice(0, 26).padEnd(28)} 응답 ${p.responded}  ${p.nextChase ? `다음 ${p.nextChase}` : ''}`);
  }

  if (!pending.length) { console.log('\n  추가할 행이 없습니다.\n'); return; }
  if (!WRITE) { console.log('\n  --write 를 붙이면 실제로 기록합니다. 지금은 쓰지 않았습니다.\n'); return; }

  const startRow = rows.length + 1;
  const values = pending.map((p, i) => [
    lastNo + i + 1, p.project, p.category, p.stage, toSerial(p.chaseDate), 'Email',
    p.counterpart, '메일 활동 자동 기록 (auto from mailbox)', p.responded, toSerial(p.nextChase), p.owner,
  ]);
  const address = `A${startRow}:K${startRow + values.length - 1}`;

  await axios.patch(`${wb}/worksheets('${SHEET}')/range(address='${address}')`, { values }, H);

  // Newly written cells inherit General, so the serials read as 46321 next to dates
  // above them. Chase date and next chase are formatted to match.
  const fmt = Array.from({ length: values.length }, () => ['yyyy-mm-dd']);
  for (const col of ['E', 'J']) {
    const range = `${col}${startRow}:${col}${startRow + values.length - 1}`;
    await axios.patch(`${wb}/worksheets('${SHEET}')/range(address='${range}')`, { numberFormat: fmt }, H);
  }

  console.log(`\n  ${values.length}행 기록 완료 · ${SHEET}!${address}\n`);
};

run().catch((e) => {
  const msg = e.response?.data?.error?.message || e.message;
  // 423 Locked is what Graph returns while someone has the workbook open for editing.
  if (e.response?.status === 423) {
    console.error('\n  시트가 편집 중으로 잠겨 있습니다. 오늘은 건너뜁니다.\n');
    process.exit(0);
  }
  console.error(`\n  Chase Log 기록 실패: ${msg}\n`);
  process.exit(1);
});

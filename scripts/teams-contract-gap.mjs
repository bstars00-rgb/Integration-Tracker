/**
 * One-off Teams message: which integrations have no contract on file.
 *
 *   node scripts/teams-contract-gap.mjs [YYYY-MM-DD]
 *   -> output/teams-contract-gap-<date>.html     Global Sales Team (Korean)
 *
 * Built from ../contracts/register.json, which is the SharePoint 채널계약서 folder
 * cross-referenced against data/tracker.json. Read it before sending:
 *   node scripts/post-teams.js msg "Global Sales Team" <html>          (dry run)
 *   node scripts/post-teams.js msg "Global Sales Team" <html> --send   (send)
 * from the AI CRM_Claude project, where post-teams.js lives.
 *
 * Partners still at first contact are left out on purpose. Having no contract at
 * that stage is the normal state, and listing 47 of them would bury the two that
 * are live and selling without one.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(root, 'output');
const BOARD_URL = 'https://bstars00-rgb.github.io/Integration-Tracker/';
const FOLDER_URL =
  'https://ohmylab.sharepoint.com/sites/msteams_b81d8f/Shared%20Documents/%EC%B1%84%EB%84%90%EA%B3%84%EC%95%BD%EC%84%9C';

const FONT = "font-family:'Segoe UI',system-ui,-apple-system,sans-serif";
const esc = (v) =>
  String(v ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

const wrap = (body) => `<div style="${FONT};font-size:14px;color:#242424">${body}</div>`;
const title = (text) =>
  `<div style="${FONT};font-size:15px;font-weight:700;margin:0 0 2px 0">${text}</div>` +
  `<div style="${FONT};font-size:12px;margin:0 0 12px 0"><a href="${BOARD_URL}">${BOARD_URL}</a></div>`;
const line = `style="${FONT};font-size:14px;line-height:1.55;margin:2px 0"`;
const rule = (text) =>
  `<div style="${FONT};font-size:13px;font-weight:700;color:#5b4bd6;margin:14px 0 4px 0">${text}</div>`;
const foot = (text) => `<div style="${FONT};color:#888;font-size:12px;margin:14px 0 0 0">${text}</div>`;
const red = (t) => `<span style="color:#c0392b;font-weight:700">${t}</span>`;
const dim = (t) => `<span style="color:#666">${t}</span>`;

/* ------------------------------------------------------------------- data */
const regPath = path.join(root, '..', 'contracts', 'register.json');
const trkPath = path.join(root, 'data', 'tracker.json');
for (const p of [regPath, trkPath]) {
  if (!fs.existsSync(p)) {
    console.error(`\n없음: ${p}\n`);
    process.exit(1);
  }
}
const reg = JSON.parse(fs.readFileSync(regPath, 'utf8'));
const trk = JSON.parse(fs.readFileSync(trkPath, 'utf8'));

const argv = process.argv.slice(2);
const stamp = argv.find((a) => /^\d{4}-\d{2}-\d{2}$/.test(a)) || new Date().toISOString().slice(0, 10);
const shortDate = (iso) => {
  const [, m, d] = iso.split('-');
  return `${Number(m)}/${Number(d)}`;
};

const rowOf = new Map(trk.rows.map((r) => [r.project, r]));

/** "GSM: Aiden → Jane (Dec~)" reads as "Aiden → Jane"; the rest is tracker bookkeeping. */
const picName = (pic) =>
  String(pic ?? '-')
    .replace(/^GSM:\s*/i, '')
    .replace(/\s*\(.*?\)\s*/g, '')
    .trim() || '-';

const gaps = reg.trackerNoContract
  .filter((r) => !r.parked)
  .map((r) => ({ ...r, row: rowOf.get(r.project) }))
  // A partner at first contact has no contract because there is nothing to sign yet.
  .filter((r) => (r.progress || 0) > 0 || r.status !== 'Contact');

const liveGaps = gaps.filter((r) => (r.progress || 0) >= 100 || r.status === 'Live');
const movingGaps = gaps
  .filter((r) => !liveGaps.includes(r))
  .sort((a, b) => (b.progress || 0) - (a.progress || 0));

const parkedN = reg.trackerNoContract.filter((r) => r.parked).length;
const contactN = reg.trackerNoContract.filter(
  (r) => !r.parked && (r.progress || 0) === 0 && r.status === 'Contact',
).length;

// The other direction: a signed contract with nobody tracking the integration.
// Some folder names are ours, not a counterparty's - OMH's own standard templates
// and catch-all folders are not companies missing from the tracker.
const NOT_A_COMPANY = /^(후불 계약서|닷비즈 선불계약서|기타문서|Loop|Ohmyhotel|OMH)/i;
const fileCount = new Map();
for (const f of reg.files) fileCount.set(f.company, (fileCount.get(f.company) || 0) + 1);
const orphanAll = reg.contractNoTracker
  .filter((c) => !NOT_A_COMPANY.test(c))
  .map((c) => ({ company: c, n: fileCount.get(c) || 0 }))
  .sort((a, b) => b.n - a.n);
// Two or more files reads as a real relationship; a lone NDA often went nowhere.
const orphanCos = orphanAll.filter((c) => c.n >= 2);

/* ------------------------------------------------------------------ build */
let b = title(`Integration Tracker &mdash; 계약서 누락 점검 (${shortDate(stamp)})`);

b += `<div ${line}>SharePoint 채널계약서 폴더 <b>${reg.files.length}건</b>(업체 ` +
  `${new Set(reg.files.map((f) => f.company)).size}곳)을 트래커 <b>${trk.rows.length}건</b>과 대조했습니다.</div>`;

if (liveGaps.length) {
  b += rule(`라이브인데 계약서가 없는 건 ${liveGaps.length}건`);
  for (const g of liveGaps) {
    const live = g.row?.liveDate ? ` ${dim(`${g.row.liveDate} 라이브`)}` : '';
    b += `<div ${line}>${red('·')} <b>${esc(g.project)}</b> ${dim(esc(g.country))} · ` +
      `담당 <b>${esc(picName(g.row?.pic))}</b>${live}</div>`;
  }
  b += `<div ${line}>판매는 돌아가는데 폴더에 계약서가 없습니다. 가지고 계시면 올려주시고, ` +
    `원본이 없으면 알려주세요.</div>`;
}

if (movingGaps.length) {
  b += rule(`진행 중인데 계약서가 없는 건 ${movingGaps.length}건`);
  for (const g of movingGaps) {
    b += `<div ${line}>· <b>${esc(g.project)}</b> ${esc(g.status)} ${g.progress || 0}% ` +
      `${dim(esc(g.country))} · 담당 <b>${esc(picName(g.row?.pic))}</b></div>`;
  }
}

if (orphanCos.length) {
  b += rule(`반대 경우 — 계약서는 있는데 트래커에 없는 업체 ${orphanAll.length}곳`);
  b += `<div ${line}>${orphanCos
    .slice(0, 12)
    .map((c) => `${esc(c.company)} <span style="color:#666">${c.n}</span>`)
    .join(' · ')}</div>`;
  b += `<div ${line}>트래커에서 빠진 건지, 종료된 계약인지 확인이 필요합니다.</div>`;
}

b += rule('올리실 곳');
b += `<div ${line}><a href="${FOLDER_URL}">SharePoint 채널계약서 폴더</a> ` +
  `${dim('— 업체 폴더가 없으면 지역 폴더 아래 새로 만들어 주세요')}</div>`;

b += foot(
  `신규 접촉 단계 ${contactN}건과 보류·드롭 ${parkedN}건은 계약 전이라 제외 · ` +
    `채널계약서 폴더 ${reg.files.length}건 기준 · ${stamp} 대조`,
);

fs.mkdirSync(OUT, { recursive: true });
const dest = path.join(OUT, `teams-contract-gap-${stamp}.html`);
fs.writeFileSync(dest, wrap(b));
console.log(`\n작성: ${dest}`);
console.log(`  라이브 누락 ${liveGaps.length} · 진행중 누락 ${movingGaps.length} · 역방향 ${orphanAll.length}`);
console.log('\n발송:');
console.log(`  node scripts/post-teams.js msg "Global Sales Team" "${dest}" --send`);
console.log('  (cwd: C:\\Users\\LENOVO\\Desktop\\AI CRM_Claude)\n');

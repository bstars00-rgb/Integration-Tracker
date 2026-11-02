/**
 * Daily tracker refresh. Scheduled for 15:00.
 *
 *   node scripts/daily.mjs [--send]
 *
 *   1. pull the current workbook from SharePoint
 *   2. rebuild and publish the board          (no email data - the repo is public)
 *   3. scan the mailbox for partner activity
 *   4. rebuild locally with email merged, and write a digest
 *   5. post the digest, only with --send
 *
 * Step 2 publishes and step 4 does not: the public board carries what the sheet says,
 * the digest carries who has not written back. Each step is independent, so a mailbox
 * that will not authenticate still leaves the board updated.
 */
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const REPORT = path.join(root, '..', '..', 'REPORT');
const send = process.argv.includes('--send');
const today = new Date().toISOString().slice(0, 10);

/**
 * When the job last finished, so a Monday run reports Saturday and Sunday too.
 * Nothing is lost without it - the mailbox scan always looks back 30 days - but the
 * digest would otherwise say "today" and quietly skip a weekend's replies.
 */
const STATE = path.join(root, 'data', 'daily-state.json');
const readState = () => {
  try { return JSON.parse(fs.readFileSync(STATE, 'utf8')); } catch { return {}; }
};
const state = readState();
const lastRun = state.lastRun || null;

/** Days this run is accountable for: everything after the last run, today included. */
const coveredDays = (() => {
  if (!lastRun || lastRun >= today) return [today];
  const out = [];
  for (let d = new Date(`${lastRun}T00:00:00Z`); ; ) {
    d = new Date(d.getTime() + 86400000);
    const iso = d.toISOString().slice(0, 10);
    out.push(iso);
    if (iso >= today) break;
  }
  return out;
})();

/**
 * The scheduled run writes nowhere, so a failed step went unnoticed for three days
 * while the digest quietly reused stale data. Everything is teed to a dated log and
 * the process exits non-zero, which is what Task Scheduler records.
 */
const LOG_DIR = path.join(root, 'output', 'daily-logs');
fs.mkdirSync(LOG_DIR, { recursive: true });
const logFile = path.join(LOG_DIR, `${today}.log`);
const logStream = fs.createWriteStream(logFile, { flags: 'a' });
const tee = (chunk) => { try { logStream.write(chunk); } catch { /* logging must not break the run */ } };
for (const key of ['log', 'error', 'warn']) {
  const original = console[key].bind(console);
  console[key] = (...args) => { tee(`${args.join(' ')}\n`); original(...args); };
}

console.log(`# ${new Date().toISOString()}  covering ${coveredDays.join(', ')}`);

const step = (title) => console.log(`\n${'='.repeat(58)}\n▶ ${title}\n${'='.repeat(58)}`);
const run = (cmd, args, cwd) => {
  // Child output is captured so it reaches the log as well as the console.
  const out = execFileSync(cmd, args, { cwd, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
  if (out) console.log(out.trimEnd());
  return out;
};

const failures = [];
const attempt = (title, fn) => {
  step(title);
  try { fn(); } catch (e) {
    failures.push(title);
    const detail = [e.message, e.stdout?.toString?.(), e.stderr?.toString?.()].filter(Boolean).join('\n');
    console.error(`\n  실패: ${detail.trim()}\n`);
  }
};
const failed = () => failures.length;

attempt('1. SharePoint 최신본 받기', () => {
  run('node', ['scripts/auto/fetch-crm-from-share.js'], REPORT);
});

attempt('2. 보드 재빌드 · 배포', () => {
  // The script, not "npm run weekly": spawning npm without a shell is ENOENT on Windows,
  // where npm is a .cmd rather than an executable.
  run('node', ['scripts/weekly-local.mjs'], root);
});

attempt('3. 메일 스캔', () => {
  run('node', ['scripts/mail-scan.mjs'], root);
});

attempt('4. 메일 반영 빌드 (로컬 전용)', () => {
  run('node', ['scripts/build.mjs', '--with-mail'], root);
  writeDigest();
});

attempt('5. Chase Log 기록', () => {
  // Rows already in the sheet are matched on project and date, so a second run on the
  // same day writes nothing. A workbook someone has open comes back 423 and is skipped
  // until tomorrow rather than retried.
  run('node', ['scripts/chase-write.mjs', '--write'], root);
});

/* ------------------------------------------------------------------ digest */
function writeDigest() {
  const tracker = JSON.parse(fs.readFileSync(path.join(root, 'data', 'tracker.json'), 'utf8'));
  const FONT = "font-family:'Segoe UI',system-ui,-apple-system,sans-serif";
  const esc = (v) => String(v ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
  const line = `style="${FONT};font-size:14px;line-height:1.55;margin:2px 0"`;
  const dim = (t) => `<span style="color:#666">${t}</span>`;
  const red = (t) => `<span style="color:#c0392b;font-weight:700">${t}</span>`;

  // One row per partner: v3 splits two-way partners by direction, and the mailbox
  // cannot tell which direction an email was about.
  const seen = new Set();
  const waiting = tracker.rows
    .filter((r) => r.mail?.waitingReply)
    .filter((r) => (seen.has(r.project) ? false : seen.add(r.project)))
    .sort((a, b) => b.mail.waitingDays - a.mail.waitingDays);

  // Everything since the last run, so Monday reports Saturday and Sunday as well.
  const inWindow = (d) => Boolean(d) && coveredDays.includes(d);
  const replied = tracker.rows.filter((r) => r.mail && !r.mail.waitingReply && inWindow(r.mail.lastIn));
  const chased = tracker.rows.filter((r) => r.mail && inWindow(r.mail.lastOut));
  const span = coveredDays.length > 1
    ? `${coveredDays[0].slice(5)}~${coveredDays[coveredDays.length - 1].slice(5)}`
    : today.slice(5);

  let b = `<div style="${FONT};font-size:15px;font-weight:700;margin:0 0 2px 0">&#9993; 파트너 메일 현황 (${today})</div>`;
  b += `<div style="${FONT};font-size:12px;margin:0 0 12px 0">${
    coveredDays.length > 1 ? `${span} 분 (${coveredDays.length}일)` : '오늘'
  } · 회신 대기는 최근 30일 기준 · 매일 15시 자동</div>`;

  if (chased.length) {
    b += `<div ${line}>보낸 곳 — ${chased.map((r) => `<b>${esc(r.project)}</b>`).join(' · ')}</div>`;
  }

  if (replied.length) {
    b += `<div ${line}>회신 온 곳 — ${replied.map((r) => `<b>${esc(r.project)}</b>`).join(' · ')}</div>`;
  }

  if (!waiting.length) {
    b += `<div ${line}>회신 대기 중인 파트너가 없습니다.</div>`;
  } else {
    b += `<div ${line}>보내고 답을 못 받은 곳 <b>${waiting.length}곳</b></div><div style="margin:8px 0 0 0">`;
    for (const r of waiting) {
      const d = r.mail.waitingDays;
      const age = d >= 14 ? red(`${d}일`) : `${d}일`;
      b += `<div ${line}>&nbsp;&nbsp;${age} &nbsp; <b>${esc(r.project)}</b> &nbsp;${dim(`${r.progress}% · ${esc(r.status)}`)}</div>`;
    }
    b += '</div>';
  }

  const unmappedPath = path.join(root, 'data', 'mail-unmapped.json');
  if (fs.existsSync(unmappedPath)) {
    const u = JSON.parse(fs.readFileSync(unmappedPath, 'utf8')).domains || [];
    if (u.length) {
      b += `<div style="${FONT};color:#888;font-size:12px;margin:14px 0 0 0">`
        + `분류 안 된 도메인 ${u.length}개 (상위: ${u.slice(0, 3).map((x) => esc(x.domain)).join(', ')}) · `
        + `partner-domains.json 에 넣으면 집계에 포함됩니다</div>`;
    }
  }

  const out = path.join(root, 'output', `mail-digest-${today}.html`);
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, `<div style="${FONT};font-size:14px;color:#242424">${b}</div>`, 'utf8');
  console.log(`\n  회신 대기 ${waiting.length}곳 · 오늘 회신 ${replied.length}곳`);
  console.log(`  ${path.relative(root, out)}`);
}

if (send) {
  attempt('5. 다이제스트 발송', () => {
    const POSTER = path.join(root, '..', '..', 'AI CRM_Claude');
    const msg = path.join(root, 'output', `mail-digest-${today}.html`);
    run('node', ['scripts/post-teams.js', 'msg', 'Global Sales Team', msg, '--send'], POSTER);
  });
} else {
  console.log('\n  발송하지 않았습니다. 보내려면 --send 를 붙이세요.');
}

// The board is left as the public build so a later manual publish cannot leak email data.
try { run('node', ['scripts/build.mjs'], root); } catch { /* best effort */ }

console.log(`\n${failed ? `완료 (실패 ${failed}단계)` : '완료'}\n`);

/**
 * Teams messages from the tracker.
 *
 *   node scripts/teams-msg.mjs [YYYY-MM-DD]
 *   -> output/teams-tracker-dev-<date>.html      IT x GSM + Contents  (English)
 *   -> output/teams-tracker-sales-<date>.html    Global Sales Team    (Korean)
 *   -> output/teams-tracker-leaders-<date>.html  Global SCM Director  (Korean)
 *
 * Three audiences, three different questions:
 *   developers - what do I owe this week
 *   sales      - which of my partners has stopped moving
 *   leaders    - how big is the problem and what needs deciding
 *
 * Only in-flight integrations appear. Live ones are finished and the not-started ones
 * are a pipeline question, not a weekly one - including either buried the ten or so
 * rows that actually moved.
 *
 * A go-live is announced exactly once, the week it happens, and then never again.
 * data/live-announced.json remembers which partners have already been named, so a
 * partner serving bookings at 80% while launch monitoring runs stops reappearing in the
 * developer list week after week. Delete an entry there to have it announced again.
 *
 * Posting is a separate step, so this can be read before anything is sent:
 *   node scripts/post-teams.js msg "<chat>" <html>          (dry run)
 *   node scripts/post-teams.js msg "<chat>" <html> --send   (send)
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(root, 'output');
const BOARD_URL = 'https://bstars00-rgb.github.io/Integration-Tracker/';

/** Days without a milestone before a partner counts as stalled. */
const STALE_DAYS = 45;

const FONT = "font-family:'Segoe UI',system-ui,-apple-system,sans-serif";
const esc = (v) =>
  String(v ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

const CH = { dev: 'IT x GSM + Contents', sales: 'Global Sales Team', leaders: 'Global SCM Director' };

/* ------------------------------------------------------------------ data */
const dataPath = path.join(root, 'data', 'tracker.json');
if (!fs.existsSync(dataPath)) {
  console.error('\nNo data/tracker.json. Run "npm run build" first.\n');
  process.exit(1);
}
const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));

const argv = process.argv.slice(2);
/**
 * Two dates, on purpose.
 *
 * stamp names the output files, and the weekly pipeline passes the report week here so
 * it can find them afterwards. asOf is what the message says, and it is the moment the
 * tracker data was actually built. They differ: the weekly for 5-11 September runs on
 * the 13th, and a message headed "9/5" carrying days-since figures counted to the 13th
 * had the reader doing arithmetic to work out which day was true.
 */
const stamp = argv.find((a) => /^\d{4}-\d{2}-\d{2}$/.test(a)) || new Date().toISOString().slice(0, 10);
const asOf = (data.generatedAt || new Date().toISOString()).slice(0, 10);
/**
 * Marks the message as a correction. Posting a second, near-identical message into the
 * same channel on the same day reads as a duplicate unless it says why it is there.
 */
const correction = argv.includes('--correction');
const shortDate = (iso) => {
  const [y, m, d] = iso.split('-');
  return `${Number(m)}/${Number(d)}`;
};

/* ------------------------------------------------------------------ live, once */
/**
 * A partner that has gone live is news exactly once.
 *
 * "Live" is the sheet's Status, not the 100% milestone: Klook and bizplay both serve
 * bookings at 80% while launch monitoring runs, and until now they came back in the
 * developer list every week at the same percentage. Announced once, then out of the
 * recurring sections for good.
 *
 * The ledger is committed alongside the data, so the answer does not depend on which
 * machine ran the weekly.
 */
const LEDGER = path.join(root, 'data', 'live-announced.json');

const isLive = (r) => r.status === 'Live' || r.progress >= 100;

function readLedger() {
  if (!fs.existsSync(LEDGER)) return {};
  try {
    return JSON.parse(fs.readFileSync(LEDGER, 'utf8'));
  } catch {
    console.warn('  live-announced.json unreadable; treating every live partner as new.');
    return {};
  }
}

const ledger = readLedger();
const liveRows = data.rows.filter(isLive);
const newlyLive = liveRows.filter((r) => !ledger[r.project]);

/** Written only when the messages are actually produced, never on a dry inspection. */
function recordAnnounced() {
  if (!newlyLive.length) return;
  for (const r of newlyLive) ledger[r.project] = asOf;
  const sorted = Object.fromEntries(Object.entries(ledger).sort(([a], [b]) => a.localeCompare(b)));
  fs.writeFileSync(LEDGER, `${JSON.stringify(sorted, null, 1)}\n`, 'utf8');
}

/**
 * The weekly working set. A live partner drops out even at 80%, because the team asked
 * not to see it again after the go-live is announced.
 */
const inFlight = data.rows.filter((r) => r.progress > 0 && r.progress < 100 && !isLive(r));

// route and watch are computed in build.mjs and shipped on the row, so these messages
// and the board cannot disagree about who owns what.
const stalled = inFlight
  .filter((r) => r.days !== null && r.days >= STALE_DAYS)
  .sort((a, b) => b.days - a.days);
const moving = inFlight.filter((r) => r.days === null || r.days < STALE_DAYS);

const OMH_WORK = ['omhbuild', 'omhsupport', 'switchreview'];
const devWork = inFlight
  .filter((r) => OMH_WORK.includes(r.watch))
  .sort((a, b) => {
    const rank = { omhbuild: 0, switchreview: 1, omhsupport: 2 };
    return rank[a.watch] - rank[b.watch] || b.progress - a.progress;
  });
const partnerWait = inFlight.filter((r) => r.watch === 'partnerbuild');

const stageOf = (row) => {
  const hit = row.stages.filter((s) => s.date).pop();
  return hit ? hit.label : '-';
};

const byPic = (rows) => {
  const m = new Map();
  for (const r of rows) {
    const key = r.pic || '-';
    if (!m.has(key)) m.set(key, []);
    m.get(key).push(r);
  }
  return [...m.entries()].sort((a, b) => b[1].length - a[1].length || b[1][0].days - a[1][0].days);
};

/* ------------------------------------------------------------------ shared bits */
const wrap = (body) =>
  `<div style="${FONT};font-size:14px;color:#242424">${correction ? CORRECTION_NOTE : ''}${body}</div>`;

const CORRECTION_NOTE =
  `<div style="${FONT};font-size:13px;color:#c0392b;font-weight:600;` +
  `border-left:3px solid #c0392b;padding:2px 0 2px 10px;margin:0 0 12px 0">` +
  `정정 — 앞서 보낸 메시지는 일주일 전 데이터로 작성됐습니다. 아래가 최신 기준입니다.</div>`;
const title = (text) =>
  `<div style="${FONT};font-size:15px;font-weight:700;margin:0 0 2px 0">${text}</div>` +
  `<div style="${FONT};font-size:12px;margin:0 0 12px 0"><a href="${BOARD_URL}">${BOARD_URL}</a></div>`;
const line = `style="${FONT};font-size:14px;line-height:1.55;margin:2px 0"`;
const rule = (text) =>
  `<div style="${FONT};font-size:13px;font-weight:700;color:#5b4bd6;margin:14px 0 4px 0">${text}</div>`;
const foot = (text) => `<div style="${FONT};color:#888;font-size:12px;margin:14px 0 0 0">${text}</div>`;
const red = (t) => `<span style="color:#c0392b;font-weight:700">${t}</span>`;
const dim = (t) => `<span style="color:#666">${t}</span>`;

/* ------------------------------------------------------------------ 1. developers */
// Only the rows where OMH engineering owes a deliverable. Everything else is noise to
// someone deciding what to pick up on Monday.
function devMessage() {
  const WHAT = {
    omhbuild: 'OMH writes the code',
    omhsupport: 'partner builds, OMH owes a deliverable',
    switchreview: 'switching platform builds, OMH reviews',
  };
  const MARK = { omhbuild: '&#128308;', omhsupport: '&#128992;', switchreview: '&#128993;' };

  let b = title(`&#128225; Integration Tracker &mdash; ${shortDate(asOf)}`);

  if (!devWork.length) {
    b += `<div ${line}>Nothing is waiting on OMH engineering this week.</div>`;
  } else {
    b += `<div ${line}>OMH engineering owes something on <b>${devWork.length}</b> integration${
      devWork.length === 1 ? '' : 's'
    }:</div>`;
    b += '<div style="margin:8px 0 0 0">';
    for (const r of devWork) {
      const age = r.days === null ? 'no record' : `${r.days}d`;
      const flag = r.days !== null && r.days >= STALE_DAYS ? ` ${red('&#9888;')}` : '';
      b += `<div ${line}>${MARK[r.watch]} <b>${esc(r.project)}</b> &nbsp;${r.progress}% &nbsp;${dim(
        esc(stageOf(r)),
      )} &nbsp;&middot;&nbsp; ${esc(WHAT[r.watch])} &nbsp;${dim(age)}${flag}</div>`;
    }
    b += '</div>';
  }

  if (newlyLive.length) {
    b += `<div ${line} style="margin-top:6px">&#127881; Live this week: ${newlyLive
      .map((r) => `<b>${esc(r.project)}</b>`)
      .join(' · ')}</div>`;
  }

  if (partnerWait.length) {
    const fresh = partnerWait.filter((r) => r.days !== null && r.days < 30).length;
    b += `<div ${line}>&#9203; Waiting on partners: <b>${partnerWait.length}</b> (${fresh} moved &lt;30d).</div>`;
  }

  b += foot(`In-flight ${inFlight.length} · stalled = no milestone ${STALE_DAYS}d+`);
  return wrap(b);
}

/* ------------------------------------------------------------------ 2. sales team */
// An alert, so it is grouped by the person who can act on it rather than by stage.
function salesMessage() {
  let b = title(`&#128680; Integration Tracker &mdash; 정체 알럿 (${shortDate(asOf)})`);

  if (!stalled.length) {
    b += `<div ${line}>진행중 ${inFlight.length}건 모두 최근 ${STALE_DAYS}일 안에 움직였습니다.</div>`;
    return wrap(b);
  }

  if (newlyLive.length) {
    b += `<div ${line}>&#127881; 이번 주 라이브 — ${newlyLive
      .map((r) => `<b>${esc(r.project)}</b>`)
      .join(' · ')}</div>`;
  }

  b += `<div ${line}>진행중 <b>${inFlight.length}건</b> 중 <b>${STALE_DAYS}일+</b> 정체 ` +
    `${red(`<b>${stalled.length}건</b>`)}.</div>`;

  // PIC별 건수 한 줄
  const picSummary = byPic(stalled).map(([pic, rows]) => `${esc(pic)} ${rows.length}`).join(' · ');
  b += `<div ${line}>PIC별: <b>${picSummary}</b></div>`;

  // 최장 정체 Top 5
  const top5 = stalled.slice(0, 5)
    .map((r) => `<b>${esc(r.project)}</b> ${red(`${r.days}일`)}${r.impact === 'High' ? ' <b>High</b>' : ''}`)
    .join(' · ');
  b += rule('최장 정체 Top 5');
  b += `<div ${line}>${top5}</div>`;

  // Never got past the NDA. These are the ones most likely to be dead rather than slow.
  const nda = stalled.filter((r) => r.progress <= 20);
  if (nda.length) {
    b += rule(`확인 필요 — NDA만 찍고 멈춘 ${nda.length}건`);
    b += `<div ${line}>${nda.map((r) => esc(r.project)).join(' · ')} &rarr; 살아있는지 확인, 아니면 <b>Hold/Drop</b></div>`;
  }

  b += foot(`진행중 ${inFlight.length}건 집계 · 매주 자동 생성`);
  return wrap(b);
}

/* ------------------------------------------------------------------ 3. leaders */
/**
 * The directors' room: 대표 · CSO · CTO and the line directors.
 *
 * They are not chasing partners, so a list of them is the wrong output. What this has
 * to answer is which of the five business lines is actually running, which is stalled,
 * and what the shape implies about where the next quarter's engineering goes.
 *
 * Live rows count here, unlike the other two messages. "Twenty-nine of our thirty live
 * integrations are on one line" is the finding - it is invisible if you only look at
 * what moved this week.
 */

/** Sheet category -> business line. Anything unmapped still shows, under its own name. */
const LINES = [
  { key: 'Channel API', label: '고객사 연동' },
  { key: '3rd Party Hotel', label: '공급사 연동' },
  { key: 'Switching System', label: '스위칭 연동' },
  { key: '3rd Party Activity', label: '액티비티 공급사' },
  { key: 'CRS', label: 'CRS' },
];
/** Lines that bring inventory in, as opposed to selling it. */
const SUPPLY = ['3rd Party Hotel', '3rd Party Activity', 'CRS'];

function lineStats() {
  const seen = new Set(LINES.map((l) => l.key));
  const extra = [...new Set(data.rows.map((r) => r.category))]
    .filter((c) => !seen.has(c))
    .map((c) => ({ key: c, label: c }));

  return [...LINES, ...extra]
    .map((l) => {
      const rows = data.rows.filter((r) => r.category === l.key);
      // Same live rule as everywhere else: the sheet's Status, not the 100% milestone.
      // Counting by milestone here left Klook and bizplay in the in-flight column while
      // the header above had already retired them.
      const live = rows.filter(isLive);
      const inf = rows.filter((r) => r.progress > 0 && r.progress < 100 && !isLive(r));
      return {
        ...l,
        rows,
        n: rows.length,
        live: live.length,
        inflight: inf.length,
        stalled: inf.filter((r) => r.days !== null && r.days >= STALE_DAYS).length,
        idle: rows.length - live.length - inf.length,
        omh: rows.filter((r) => ['direct', 'shared'].includes(r.route)).length,
        conv: rows.length ? Math.round((live.length / rows.length) * 100) : 0,
      };
    })
    .filter((l) => l.n > 0)
    .sort((a, b) => b.n - a.n);
}

/** Partners that appear on more than one line - selling to us and buying from us. */
function twoWay() {
  const norm = (s) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
  const m = new Map();
  for (const r of data.rows) {
    const k = norm(r.project);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  }
  return [...m.values()]
    .filter((v) => v.length > 1 && new Set(v.map((r) => r.category)).size > 1)
    .sort((a, b) => Math.max(...b.map((r) => r.progress)) - Math.max(...a.map((r) => r.progress)));
}

function leadersMessage() {
  const lines = lineStats();
  const totalLive = liveRows.length;
  const idle = data.rows.filter((r) => r.progress === 0).length;

  let b = title(`&#128225; Integration Tracker &mdash; 라인별 현황 (${shortDate(asOf)})`);
  b += `<div ${line}>전체 <b>${data.rows.length}건</b> · 라이브 <b>${totalLive}</b> · ` +
    `진행중 <b>${inFlight.length}</b> (${STALE_DAYS}일+ 정체 ${red(`${stalled.length}`)}) · 미착수 <b>${idle}</b></div>`;

  if (newlyLive.length) {
    b += `<div ${line}>&#127881; 이번 주 라이브 — ${newlyLive
      .map((r) => `<b>${esc(r.project)}</b> ${dim(esc(r.category))}`)
      .join(' · ')}</div>`;
  }

  /* ---- the five lines, side by side ---- */
  const th = `style="${FONT};font-size:12px;color:#666;font-weight:600;text-align:right;padding:4px 8px;border-bottom:1px solid #ddd"`;
  const td = `style="${FONT};font-size:14px;text-align:right;padding:5px 8px;border-bottom:1px solid #eee"`;
  const tdL = `style="${FONT};font-size:14px;text-align:left;padding:5px 8px;border-bottom:1px solid #eee"`;

  b += rule('라인별');
  b += `<table style="border-collapse:collapse;margin:2px 0 0 0">
    <tr>
      <th ${th} style="text-align:left">라인</th><th ${th}>총</th><th ${th}>라이브</th>
      <th ${th}>진행중</th><th ${th}>미착수</th><th ${th}>전환율</th><th ${th}>OMH 구현</th>
    </tr>`;
  for (const l of lines) {
    const stall = l.stalled ? ` ${red(`(정체 ${l.stalled})`)}` : '';
    const conv = l.conv === 0 ? red('0%') : `${l.conv}%`;
    b += `<tr>
      <td ${tdL}><b>${esc(l.label)}</b></td>
      <td ${td}>${l.n}</td><td ${td}><b>${l.live}</b></td>
      <td ${td}>${l.inflight}${stall}</td><td ${td}>${l.idle}</td>
      <td ${td}>${conv}</td><td ${td}>${l.omh}</td>
    </tr>`;
  }
  b += '</table>';

  /* ---- 핵심 (한 줄씩) ---- */
  b += rule('핵심');
  const supplyLive = lines.filter((l) => SUPPLY.includes(l.key)).reduce((s, l) => s + l.live, 0);
  const top = lines.slice().sort((a, b2) => b2.live - a.live)[0];
  if (top && totalLive > 0 && top.live / totalLive >= 0.8) {
    b += `<div ${line}>&middot; <b>매출 집중</b> — 라이브 ${totalLive}건 중 <b>${top.live}건이 ${esc(top.label)}</b>, 소스(공급) 라이브 ${supplyLive}건.</div>`;
  }
  const sw = lines.find((l) => l.key === 'Switching System');
  if (sw && sw.live === 0) {
    const behind = data.rows.filter((r) => r.route === 'switching' && r.progress < 100).length;
    b += `<div ${line}>&middot; <b>스위칭 라이브 ${red('0')}</b> — 미실현 레버리지 (경유 대기 ${behind}건).</div>`;
  }
  const supply = lines.filter((l) => SUPPLY.includes(l.key));
  const sN = supply.reduce((s, l) => s + l.n, 0);
  const sOmh = supply.reduce((s, l) => s + l.omh, 0);
  if (sN && sOmh / sN >= 0.7) {
    b += `<div ${line}>&middot; <b>공급 라인 = OMH 직접 개발</b> ${sOmh}/${sN}건 (현재 가동 ${devWork.length}).</div>`;
  }
  const both = twoWay();
  if (both.length) {
    b += `<div ${line}>&middot; <b>양방향 상대</b> ${both.length}곳 — ${both.slice(0, 5).map((v) => esc(v[0].project)).join(' · ')}${both.length > 5 ? ' 외' : ''}.</div>`;
  }

  /* ---- 판단 필요 ---- */
  const advanced = stalled.filter((r) => r.progress >= 50).sort((a, b2) => b2.progress - a.progress || b2.days - a.days);
  if (advanced.length) {
    b += rule(`판단 필요 — 50%+ 진행 중 정체 ${advanced.length}건`);
    b += `<div ${line}>${advanced.slice(0, 6).map((r) => `<b>${esc(r.project)}</b> ${red(`${r.days}일`)}`).join(' · ')}` +
      `${advanced.length > 6 ? ` 외 ${advanced.length - 6}건` : ''}</div>`;
  }

  const noImpact = data.rows.length - (data.counts?.withImpact ?? 0);
  b += foot(`정체 기준 ${STALE_DAYS}일+ · 라인=시트 Category · 매주 자동${noImpact ? ` · Biz Impact 미입력 ${noImpact}/${data.rows.length}` : ''}`);
  return wrap(b);
}

/* ------------------------------------------------------------------ write */
fs.mkdirSync(OUT, { recursive: true });

const files = [
  ['dev', devMessage()],
  ['sales', salesMessage()],
  ['leaders', leadersMessage()],
].map(([kind, html]) => {
  const file = path.join(OUT, `teams-tracker-${kind}-${stamp}.html`);
  fs.writeFileSync(file, html, 'utf8');
  return { kind, file, size: html.length };
});

console.log(`\n  In-flight ${inFlight.length}  ·  stalled ${STALE_DAYS}d+ ${stalled.length}  ·  OMH work ${devWork.length}  ·  partner wait ${partnerWait.length}\n`);
for (const f of files) {
  console.log(`  ${CH[f.kind].padEnd(22)} ${path.relative(root, f.file)}  (${f.size} chars)`);
}
  // Recorded only after the files exist, so a failed run does not silently mark a
  // go-live as already announced and swallow it.
  recordAnnounced();
  if (newlyLive.length) {
    console.log(`  live announced (once only): ${newlyLive.map((r) => r.project).join(', ')}`);
  }

console.log('\n  Nothing sent. To post one:');
console.log(`    node scripts/post-teams.js msg "${CH.dev}" "<file>" --send\n`);

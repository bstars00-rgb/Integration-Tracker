/**
 * Build the Contracts tab's dataset.
 *
 *   node scripts/contracts-build.mjs
 *   -> data/contracts.json + data/contracts.js
 *
 * Three inputs, in order of authority:
 *   1. 계약 Contracts sheet in data/Dev_Schedule.xlsx  - what a person curated
 *   2. ../contracts/terms.json                         - what OCR read, with quotes
 *   3. ../contracts/register.json                      - what files exist in SharePoint
 *
 * A curated value always wins. An OCR value is only used when it is backed by a
 * quote that a person can check, and it is labelled as such in the board, because
 * a commission rate or an expiry date read off a scan is a claim, not a fact.
 *
 * Both outputs are gitignored. The repository is public, and contract dates, rates
 * and notes have no business on a public GitHub Pages site.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as XLSX from 'xlsx';

// The ESM build will not touch the filesystem until it is handed fs.
XLSX.set_fs(fs);

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const REG = path.join(root, '..', 'contracts', 'register.json');
const TERMS = path.join(root, '..', 'contracts', 'terms.json');
const BOOK = path.join(root, 'data', 'Dev_Schedule.xlsx');
const SIDECAR = path.join(root, '..', 'contracts', '계약관리.xlsx');
const TRK = path.join(root, 'data', 'tracker.json');
const SHEET = '계약 Contracts';

const today = new Date().toISOString().slice(0, 10);
const EXPIRING_DAYS = 90;

const tidy = (v) => String(v ?? '').replace(/\s+/g, ' ').trim();
const pad2 = (n) => String(n).padStart(2, '0');

/** Excel dates arrive as Date objects; local getters avoid the UTC+9 off-by-one. */
function isoOf(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date) {
    return `${value.getFullYear()}-${pad2(value.getMonth() + 1)}-${pad2(value.getDate())}`;
  }
  const s = tidy(value);
  let m = /^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/.exec(s);
  if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;
  m = /^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})/.exec(s);
  if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;
  return '';
}

const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

/** Every date inside one quote, in the order they appear. */
function datesIn(text) {
  const out = [];
  const s = String(text ?? '');
  const num = /(20\d{2})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})/g;
  for (let m = num.exec(s); m; m = num.exec(s)) {
    const [, y, mo, d] = m;
    if (+mo >= 1 && +mo <= 12 && +d >= 1 && +d <= 31) out.push({ at: m.index, iso: `${y}-${pad2(mo)}-${pad2(d)}` });
  }
  const dmy = /(\d{1,2})\s+([A-Za-z]{3})\w*,?\s+(20\d{2})/g;
  for (let m = dmy.exec(s); m; m = dmy.exec(s)) {
    const mo = MONTHS[m[2].toLowerCase()];
    if (mo) out.push({ at: m.index, iso: `${m[3]}-${pad2(mo)}-${pad2(m[1])}` });
  }
  const mdy = /([A-Za-z]{3})\w*\s+(\d{1,2}),?\s+(20\d{2})/g;
  for (let m = mdy.exec(s); m; m = mdy.exec(s)) {
    const mo = MONTHS[m[1].toLowerCase()];
    if (mo) out.push({ at: m.index, iso: `${m[3]}-${pad2(mo)}-${pad2(m[2])}` });
  }
  const seen = new Set();
  return out
    .sort((a, b) => a.at - b.at)
    .filter((x) => (seen.has(x.iso) ? false : seen.add(x.iso)))
    .map((x) => x.iso);
}

/**
 * The same dates, each with the few characters either side, so a caller can tell
 * "…부터 2023년 7월 5일" from a date that merely happens to be nearby.
 */
function datesWithContext(text) {
  const s = String(text ?? '');
  const all = [];
  const patterns = [
    /(20\d{2})\s*[.\-/년年]\s*(\d{1,2})\s*[.\-/월月]\s*(\d{1,2})/g,
    /(\d{1,2})\s+([A-Za-z]{3})\w*,?\s+(20\d{2})/g,
    /([A-Za-z]{3})\w*\s+(\d{1,2}),?\s+(20\d{2})/g,
  ];
  for (const [i, re] of patterns.entries()) {
    for (let m = re.exec(s); m; m = re.exec(s)) {
      let iso = '';
      if (i === 0) {
        const [, y, mo, d] = m;
        if (+mo < 1 || +mo > 12 || +d < 1 || +d > 31) continue;
        iso = `${y}-${pad2(mo)}-${pad2(d)}`;
      } else {
        const mo = MONTHS[(i === 1 ? m[2] : m[1]).toLowerCase()];
        if (!mo) continue;
        iso = i === 1 ? `${m[3]}-${pad2(mo)}-${pad2(m[1])}` : `${m[3]}-${pad2(mo)}-${pad2(m[2])}`;
      }
      all.push({
        iso,
        at: m.index,
        before: s.slice(Math.max(0, m.index - 14), m.index),
        after: s.slice(m.index + m[0].length, m.index + m[0].length + 14),
      });
    }
  }
  const seen = new Set();
  return all
    .sort((a, b) => a.at - b.at)
    .filter((x) => (seen.has(x.at) ? false : seen.add(x.at)));
}

/** A quote only yields a start/end pair when it actually talks about a term. */
const PERIOD_WORDS =
  /계약\s*기간|유효\s*기간|존속기간|term of (?:this|the) agreement|valid (?:for|from|until|through)|effective (?:date|day|period)|有效期|合同期限|合約期限|期限/i;

// A date immediately after one of these opens the term; after the other, closes it.
const FROM_MARK = /(?:부터|로부터|from|commencing on|starting|自|起)\s*$/i;
const TO_MARK = /(?:까지|to|until|till|through|through to|至|迄)\s*$/i;

/**
 * Only an explicit range counts.
 *
 * Taking the earliest and latest date in the quote looked fine until Myrealtrip:
 * its clause reads "본 계약 체결일로부터 2025년 2월 28일 ... 까지", so the start is a
 * defined term, not a date, and the picker invented 2025-01-01 from an unrelated
 * date nearby. An invented start date on a contract row is worse than a blank, so
 * both ends now have to be marked by 부터/까지, from/to or 自/至.
 */
function periodFrom(hits) {
  for (const h of hits || []) {
    if (!PERIOD_WORDS.test(h.quote)) continue;
    const found = datesWithContext(h.quote);
    let start = found.find((d) => FROM_MARK.test(d.before));
    let end = found.find((d) => TO_MARK.test(d.after) || TO_MARK.test(d.before));

    // Failing an explicit range, a pair that lands on an anniversary is almost
    // always a term: 2023-09-08 to 2024-09-07 is a year to the day, which no
    // unrelated pair of dates in a contract does by accident.
    if (!start || !end) {
      const anniversary = [];
      for (let i = 0; i < found.length; i += 1) {
        for (let j = i + 1; j < found.length; j += 1) {
          const [a, b] = found[i].iso < found[j].iso ? [found[i], found[j]] : [found[j], found[i]];
          const days = (Date.parse(`${b.iso}T00:00:00Z`) - Date.parse(`${a.iso}T00:00:00Z`)) / 86400000;
          for (let years = 1; years <= 5; years += 1) {
            if (Math.abs(days - years * 365.25) <= 3) anniversary.push([a, b]);
          }
        }
      }
      if (anniversary.length !== 1) continue;
      [start, end] = anniversary[0];
    }

    if (!start || !end || start.iso >= end.iso) continue;
    const span = (Date.parse(`${end.iso}T00:00:00Z`) - Date.parse(`${start.iso}T00:00:00Z`)) / 86400000;
    // A real term runs months to years.
    if (span < 30 || span > 7300) continue;
    return { start: start.iso, end: end.iso, page: h.page, engine: h.engine, quote: h.quote };
  }
  return null;
}

/**
 * Commission is reported as "there is a clause", never as a number.
 *
 * Reading the rate off a scan was measured and it was wrong nearly every time:
 * the "1%" that came back for twenty-five partners is a late-payment penalty
 * ("지연된 날짜에 대하여 청구된 금액의 1%를 추가로 지불"), Agoda's "45" is a row count in a
 * tier table, and RateHawk's "0.1%" is a one-off compensation. A rate that wrong
 * would travel straight into a negotiation, so the tab points at the clause and
 * lets a person read it.
 */
function commissionFrom(hits) {
  const list = hits || [];
  if (!list.length) return null;
  const best = list.find((h) => h.values && h.values.length) || list[0];
  return { value: '', hits: list.length, page: best.page, engine: best.engine, quote: best.quote };
}

/* ------------------------------------------------------------------ inputs */
if (!fs.existsSync(REG)) {
  console.error(`\n없음: ${REG}\n  먼저 계약 대장을 만들어야 합니다.\n`);
  process.exit(1);
}
const reg = JSON.parse(fs.readFileSync(REG, 'utf8'));
const trk = fs.existsSync(TRK) ? JSON.parse(fs.readFileSync(TRK, 'utf8')) : { rows: [] };
const terms = fs.existsSync(TERMS) ? JSON.parse(fs.readFileSync(TERMS, 'utf8')) : [];
const termsBy = new Map(terms.map((t) => [`${t.name}|${t.folder}`, t]));

/** "있음" / "Y" / "auto" all mean the same thing to whoever typed it. */
function renewalOf(v) {
  const s = tidy(v).toLowerCase();
  if (!s) return '';
  if (/^(있음|자동|자동갱신|y|yes|auto|o|예)$/.test(s)) return 'auto';
  if (/^(없음|n|no|none|x|아니오|아니요)$/.test(s)) return 'none';
  return '';
}

/**
 * Curated rows, keyed by partner. Both files are optional - with neither, the tab
 * still works and says every value is an OCR estimate.
 *
 *   ../contracts/계약관리.xlsx   the standalone sheet people fill in
 *   data/Dev_Schedule.xlsx       the same sheet, if it is added to the master
 *
 * The master wins where both carry a partner, because that is the shared copy.
 */
function readSheet() {
  const map = new Map();
  for (const book of [SIDECAR, BOOK]) {
    if (!fs.existsSync(book)) continue;
    let wb;
    try {
      wb = XLSX.readFile(book, { cellDates: true });
    } catch (err) {
      console.warn(`  열지 못함 (건너뜀): ${path.basename(book)} — ${err.message}`);
      continue;
    }
    if (!wb.SheetNames.includes(SHEET)) continue;
    for (const r of XLSX.utils.sheet_to_json(wb.Sheets[SHEET], { defval: '' })) {
      const key = tidy(r['업체'] || r['Company'] || r['파트너'] || '');
      if (!key) continue;
      const entry = {
        company: key,
        project: tidy(r['연동건'] || r['트래커 연동건'] || r['Project'] || ''),
        start: isoOf(r['계약시점'] || r['Start'] || ''),
        end: isoOf(r['계약종료시점'] || r['End'] || ''),
        renewal: renewalOf(r['자동갱신'] || r['Renewal'] || ''),
        commission: tidy(r['수수료'] || r['Commission'] || '').replace(/%$/, ''),
        note: tidy(r['특이사항'] || r['Note'] || ''),
        owner: tidy(r['담당'] || r['PIC'] || ''),
      };
      // A row nobody has filled in yet must not count as confirmed, or the board
      // would report "confirmed, term unknown" for all 163 seeded rows.
      if (!entry.start && !entry.end && !entry.renewal && !entry.commission && !entry.note) continue;
      map.set(key.toLowerCase(), entry);
    }
  }
  return map;
}
const curated = readSheet();

/* ------------------------------------------------------------------- merge */
const NOT_A_COMPANY = /^(후불 계약서|닷비즈 선불계약서|기타문서|Loop|Ohmyhotel|OMH)/i;
const DEAL_TYPES = new Set(['Agreement', 'Addendum', 'SLA', 'NDA']);

const picOf = new Map(trk.rows.map((r) => [r.project, r.pic || '']));
const statusOf = new Map(trk.rows.map((r) => [r.project, r.status || '']));
const progressOf = new Map(trk.rows.map((r) => [r.project, r.progress ?? null]));

/** "GSM: Aiden → Jane (Dec~)" reads as "Aiden → Jane"; the rest is bookkeeping. */
const picName = (pic) =>
  String(pic ?? '')
    .replace(/^GSM:\s*/i, '')
    .replace(/\s*\(.*?\)\s*/g, '')
    .trim();

// One row per company, carrying its documents. Keyed by the tracker partner where
// there is one, so the same company filed under two folders - "Cola" under Taiwan
// and "COLA TOUR" under .biz - lands on one row instead of two half-empty ones.
const byCompany = new Map();
for (const f of reg.files) {
  if (NOT_A_COMPANY.test(f.company)) continue;
  const key = f.trackerProject || f.company;
  if (!byCompany.has(key)) {
    byCompany.set(key, { company: f.company, names: new Set(), project: f.trackerProject || '', docs: [] });
  }
  const c = byCompany.get(key);
  c.names.add(f.company);
  // With a tracker project in its own column, the shortest folder name is the
  // cleanest label ("Tourmind", not "Tourmind_제오픈"); without one, the fuller
  // name carries more of what the folder actually is.
  const better = f.trackerProject
    ? f.company.length < c.company.length
    : f.company.length > c.company.length;
  if (better) c.company = f.company;
  if (!c.project && f.trackerProject) c.project = f.trackerProject;
  const t = termsBy.get(`${f.fileName}|${f.folder}`);
  c.docs.push({
    docType: f.docType,
    docDate: f.docDate,
    fileName: f.fileName,
    webUrl: f.webUrl,
    bucket: f.bucket,
    region: f.region,
    pages: t?.npages ?? null,
    method: t?.method ?? '',
    unreadable: Boolean(t) && t.chars === 0,
    period: t ? periodFrom([...(t.period || []), ...(t.renewal || [])]) : null,
    commission: t ? commissionFrom(t.commission) : null,
    autoRenew: t ? (t.noRenewal?.length ? 'none' : t.renewal?.length ? 'auto' : '') : '',
    renewQuote: t?.renewal?.[0]?.quote || t?.noRenewal?.[0]?.quote || '',
    disagree: (t?.pctDisagree || []).length > 0,
  });
}

/** The newest document that could carry a term, so the tab shows the live one. */
function leadDoc(docs) {
  const deals = docs.filter((d) => DEAL_TYPES.has(d.docType));
  const pool = deals.length ? deals : docs;
  return [...pool].sort((a, b) => String(b.docDate).localeCompare(String(a.docDate)))[0] || null;
}

const rows = [];
for (const c of byCompany.values()) {
  const cur = curated.get(c.company.toLowerCase());
  const lead = leadDoc(c.docs);
  const per = c.docs.map((d) => d.period).find(Boolean) || null;
  const com = c.docs.map((d) => d.commission).find(Boolean) || null;
  const ren = c.docs.find((d) => d.autoRenew) || null;

  const start = cur?.start || per?.start || '';
  const end = cur?.end || per?.end || '';
  const renewal = cur?.renewal || ren?.autoRenew || '';
  rows.push({
    company: c.company,
    project: cur?.project || c.project || '',
    pic: picName(cur?.owner || picOf.get(c.project) || ''),
    folders: [...c.names].sort(),
    trackerStatus: statusOf.get(c.project) || '',
    progress: progressOf.get(c.project) ?? null,
    docCount: c.docs.length,
    docTypes: [...new Set(c.docs.map((d) => d.docType))],
    latestDoc: lead?.fileName || '',
    latestDocDate: lead?.docDate || '',
    latestDocUrl: lead?.webUrl || '',
    bucket: lead?.bucket || '',
    region: lead?.region || '',
    start,
    end,
    renewal,
    commission: cur?.commission || com?.value || '',
    note: cur?.note || '',
    curated: Boolean(cur),
    // Where each inferred value came from, so the board can show it as unconfirmed.
    evidence: {
      period: cur?.start || cur?.end ? null : per,
      commission: cur?.commission ? null : com,
      renewal: cur?.renewal ? null : (ren ? { page: null, engine: '', quote: ren.renewQuote } : null),
    },
    unreadable: c.docs.some((d) => d.unreadable),
    disagree: c.docs.some((d) => d.disagree),
    hasContract: true,
    docs: c.docs.map(({ period, commission, renewQuote, ...rest }) => rest),
  });
}

// Tracker rows with no contract folder at all. They belong in this tab more than
// anywhere else: the question "where is the contract" is a contract question.
//
// The list comes from register.json rather than being matched again here. The
// Teams message reads the same list, and the same question answered two ways in
// two places is worse than either answer.
const trkRowOf = new Map(trk.rows.map((r) => [r.project, r]));
for (const miss of reg.trackerNoContract) {
  const r = { ...miss, ...(trkRowOf.get(miss.project) || {}) };
  rows.push({
    company: r.project,
    project: r.project,
    pic: picName(r.pic || ''),
    trackerStatus: r.status || '',
    progress: r.progress ?? null,
    docCount: 0,
    docTypes: [],
    latestDoc: '', latestDocDate: '', latestDocUrl: '',
    bucket: r.category || '', region: r.country || '',
    start: '', end: '', renewal: '', commission: '', note: '',
    curated: false,
    evidence: { period: null, commission: null, renewal: null },
    unreadable: false, disagree: false,
    hasContract: false,
    parked: Boolean(r.parked),
    docs: [],
  });
}

/** What the tab sorts and colours by. */
const dayDiff = (iso) => Math.round((Date.parse(`${iso}T00:00:00Z`) - Date.parse(`${today}T00:00:00Z`)) / 86_400_000);
for (const r of rows) {
  r.daysToEnd = r.end ? dayDiff(r.end) : null;
  if (!r.hasContract) {
    // "No contract" only means something where a contract should already exist.
    // A partner at first contact has nothing to sign yet, and 47 of those at the
    // top of the table buried the handful that are live without one.
    r.state = r.parked ? 'parked'
      : (r.progress || 0) > 0 || r.trackerStatus !== 'Contact' ? 'missing'
      : 'lead';
  } else if (r.daysToEnd == null) r.state = 'unknown';
  else if (r.daysToEnd < 0) r.state = r.renewal === 'auto' ? 'renewed' : 'expired';
  else if (r.daysToEnd <= EXPIRING_DAYS) r.state = 'expiring';
  else r.state = 'active';
}

rows.sort((a, b) => {
  const rank = { expired: 0, expiring: 1, missing: 2, renewed: 3, unknown: 4, active: 5, lead: 6, parked: 7 };
  if (rank[a.state] !== rank[b.state]) return rank[a.state] - rank[b.state];
  if (a.daysToEnd != null && b.daysToEnd != null) return a.daysToEnd - b.daysToEnd;
  return a.company.localeCompare(b.company);
});

const counts = rows.reduce((m, r) => ({ ...m, [r.state]: (m[r.state] || 0) + 1 }), {});
const payload = {
  generatedAt: new Date().toISOString(),
  today,
  expiringDays: EXPIRING_DAYS,
  curatedSheet: curated.size ? SHEET : '',
  curatedCount: curated.size,
  sourceFolder: '채널계약서 (SharePoint)',
  fileCount: reg.files.length,
  counts,
  rows,
};

fs.writeFileSync(path.join(root, 'data', 'contracts.json'), JSON.stringify(payload, null, 1));
fs.writeFileSync(
  path.join(root, 'data', 'contracts.js'),
  `/* Generated by scripts/contracts-build.mjs — do not edit. */\nwindow.CONTRACTS = ${JSON.stringify(payload)};\n`,
);

console.log(`\n  계약 ${rows.length}행  ->  data/contracts.json`);
console.log(`  ${Object.entries(counts).map(([k, v]) => `${k} ${v}`).join(' · ')}`);
console.log(`  큐레이션 시트: ${curated.size ? `${SHEET} ${curated.size}행` : '없음 (OCR 추정값 표시)'}\n`);

/**
 * One-off announcements, posted alongside the weekly and never repeated.
 *
 *   node scripts/announce.mjs [--send]
 *
 * Drop an HTML file into announcements/ named <channel-key>--<slug>.html and the next
 * weekly posts it once. announcements/sent.json records what has gone out, so a rerun,
 * a restored working tree or a second machine cannot post the same notice twice - the
 * same guarantee the go-live ledger gives, for the same reason.
 *
 * Without --send this lists what would be posted and sends nothing.
 */
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const DIR = path.join(root, 'announcements');
const SENT = path.join(DIR, 'sent.json');
const POSTER = path.join(root, '..', '..', 'AI CRM_Claude');

/** Channel key in the filename -> the Teams chat to post into. */
const CHANNELS = {
  dev: 'IT × GSM + Contents',
  sales: 'Global Sales Team',
  leaders: 'Global SCM Director',
};

const send = process.argv.includes('--send');

if (!fs.existsSync(DIR)) {
  console.log('\n  No announcements/ directory. Nothing to post.\n');
  process.exit(0);
}

const sent = fs.existsSync(SENT) ? JSON.parse(fs.readFileSync(SENT, 'utf8')) : {};

const pending = fs
  .readdirSync(DIR)
  .filter((f) => f.endsWith('.html') && !sent[f])
  .sort();

if (!pending.length) {
  console.log('\n  No pending announcements.\n');
  process.exit(0);
}

console.log('');
let posted = 0;
for (const file of pending) {
  const key = file.split('--')[0];
  const chat = CHANNELS[key];
  if (!chat) {
    console.warn(`  ${file}: filename must start with one of ${Object.keys(CHANNELS).join(' / ')} — skipped.`);
    continue;
  }

  const full = path.join(DIR, file);
  if (!send) {
    console.log(`  [dry] ${chat.padEnd(22)} ${file}`);
    continue;
  }

  try {
    execFileSync('node', ['scripts/post-teams.js', 'msg', chat, full, '--send'], {
      cwd: POSTER,
      stdio: 'inherit',
    });
    // Recorded per file the moment it lands, so a failure halfway through does not
    // re-post the ones that already went.
    sent[file] = new Date().toISOString().slice(0, 10);
    fs.writeFileSync(SENT, `${JSON.stringify(sent, null, 1)}\n`, 'utf8');
    posted += 1;
  } catch (e) {
    console.error(`  ✗ ${chat}: ${e.message}`);
  }
}

console.log(send ? `\n  Announcements posted: ${posted}/${pending.length}\n` : '\n  Dry run — nothing sent. Add --send to post.\n');

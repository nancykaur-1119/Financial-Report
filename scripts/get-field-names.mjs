import 'dotenv/config';
import path from 'path';
import { fileURLToPath } from 'url';
import { readFileSync } from 'fs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
try {
  const env = readFileSync(path.join(__dirname, '..', '.env'), 'utf8');
  for (const line of env.split('\n')) {
    const [k, ...v] = line.split('=');
    if (k && !process.env[k.trim()]) process.env[k.trim()] = v.join('=').trim();
  }
} catch {}

const JIRA_BASE  = (process.env.JIRA_BASE_URL || '').replace(/\/$/, '');
const authHeader = 'Basic ' + Buffer.from(`${process.env.JIRA_EMAIL}:${process.env.JIRA_API_TOKEN}`).toString('base64');

const res = await fetch(`${JIRA_BASE}/rest/api/3/field`, { headers: { Authorization: authHeader, Accept: 'application/json' } });
const fields = await res.json();

// All custom fields seen on FIN-5558 / FIN-5593 that have numeric values
const targetIds = [
  'customfield_10225', 'customfield_10383', 'customfield_10428', 'customfield_10476',
  'customfield_10478', 'customfield_10479', 'customfield_11364', 'customfield_11366',
  'customfield_11369', 'customfield_11373', 'customfield_11377', 'customfield_11378',
  'customfield_11382', 'customfield_11384', 'customfield_11400', 'customfield_11401',
  'customfield_11404', 'customfield_11405', 'customfield_11406', 'customfield_11409',
  'customfield_11411', 'customfield_11412', 'customfield_11414', 'customfield_11415',
  'customfield_11416', 'customfield_11417', 'customfield_11418', 'customfield_11420',
  'customfield_11421', 'customfield_11423', 'customfield_11433', 'customfield_11434',
  'customfield_11440', 'customfield_11442', 'customfield_11992', 'customfield_11993',
  'customfield_11994', 'customfield_11995', 'customfield_12193', 'customfield_12223',
  'customfield_12618', 'customfield_12915', 'customfield_12916', 'customfield_12917',
  'customfield_12918', 'customfield_12981', 'customfield_15167',
];

for (const f of fields) {
  if (targetIds.includes(f.id)) {
    console.log(`${f.id.padEnd(22)} ${f.name}`);
  }
}

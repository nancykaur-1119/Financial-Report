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

async function jiraPost(p, body) {
  const res = await fetch(`${JIRA_BASE}${p}`, {
    method: 'POST',
    headers: { Authorization: authHeader, Accept: 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`${res.status}`);
  return res.json();
}

// Fetch 50 AR - Reselling FIN tickets
const data = await jiraPost('/rest/api/3/search/jql', {
  jql: `project = FIN AND issuetype = "AR - Reselling" ORDER BY created DESC`,
  maxResults: 50,
  fields: ['summary', 'customfield_11529', 'customfield_11418', 'customfield_11364'],
});

let has11529 = 0, null11529 = 0;
console.log(`\nChecking ${data.issues.length} AR - Reselling tickets:\n`);
console.log(`${'Key'.padEnd(12)} ${'11529 (Cost Price)'.padEnd(22)} ${'11418 (OEM Quote)'.padEnd(22)} ${'11364 (OEM Purchase)'.padEnd(22)}`);
console.log('-'.repeat(80));

for (const issue of data.issues) {
  const v11529 = issue.fields.customfield_11529;
  const v11418 = issue.fields.customfield_11418;
  const v11364 = issue.fields.customfield_11364;
  if (v11529 !== null && v11529 !== undefined) has11529++;
  else null11529++;
  const show = v11529 !== null || v11418 !== null;
  if (show) {
    console.log(`${issue.key.padEnd(12)} ${String(v11529 ?? 'NULL').padEnd(22)} ${String(v11418 ?? 'NULL').padEnd(22)} ${String(v11364 ?? 'NULL').padEnd(22)}`);
  }
}

console.log(`\nSummary (${data.issues.length} tickets):`);
console.log(`  customfield_11529 (Cost Price) populated: ${has11529}`);
console.log(`  customfield_11529 (Cost Price) NULL:      ${null11529}`);

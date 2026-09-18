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

const res = await fetch(`${JIRA_BASE}/rest/api/3/search/jql`, {
  method: 'POST',
  headers: { Authorization: authHeader, Accept: 'application/json', 'Content-Type': 'application/json' },
  body: JSON.stringify({
    jql: `issue = ISM-81435`,
    maxResults: 1,
    fields: ['summary', 'issuelinks', 'customfield_11077'],
  }),
});
const data = await res.json();

for (const issue of (data.issues || [])) {
  const rf = issue.fields.customfield_11077;
  const links = issue.fields.issuelinks || [];
  const finLinks = links.filter(l => (l.outwardIssue?.key || l.inwardIssue?.key || '').startsWith('FIN-'));
  console.log(`\n${issue.key} | cf[11077]=${JSON.stringify(rf?.value || rf)} | issuelinks=${links.length} | FIN links=${finLinks.length}`);
  for (const l of finLinks) {
    const k = l.outwardIssue?.key || l.inwardIssue?.key;
    console.log(`  -> ${k} (${l.type?.name})`);
  }
}

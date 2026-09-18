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
  const res = await fetch(`${JIRA_BASE}${p}`, { method: 'POST', headers: { Authorization: authHeader, Accept: 'application/json', 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`${res.status} ${await res.text()}`);
  return res.json();
}

// Search for AP - Reselling FIN tickets
const data = await jiraPost('/rest/api/3/search/jql', {
  jql: `project = FIN AND issuetype = "AP - Reselling" ORDER BY created DESC`,
  maxResults: 5,
  fields: ['summary', 'issuelinks', 'issuetype', 'customfield_11364', 'customfield_10428', 'customfield_11995'],
});

console.log(`Total AP - Reselling: ${data.total}`);
for (const issue of (data.issues || [])) {
  console.log(`\n${issue.key} | ${issue.fields.summary}`);
  console.log(`  OEM Purchase ($): ${issue.fields.customfield_11364}`);
  console.log(`  License Amount ($): ${issue.fields.customfield_10428}`);
  console.log(`  Invoice USD: ${issue.fields.customfield_11995}`);
  const links = issue.fields.issuelinks || [];
  for (const l of links) {
    const rel = l.outwardIssue || l.inwardIssue;
    console.log(`  [${l.outwardIssue ? 'out' : 'in'}] type="${l.type?.name}" key=${rel?.key}`);
  }
}

// Also check FIN tickets linked inside AR FIN tickets from our ISM list
console.log('\n\n--- Checking links inside AR FIN tickets (ISM-81435 → FIN-4351) ---');
const arFin = await jiraPost('/rest/api/3/search/jql', {
  jql: `issue in linkedIssues("ISM-81435") AND project = FIN`,
  maxResults: 10,
  fields: ['summary', 'issuetype'],
});
for (const i of (arFin.issues || [])) {
  console.log(`  ${i.key} | ${i.fields.issuetype?.name} | ${i.fields.summary}`);
}

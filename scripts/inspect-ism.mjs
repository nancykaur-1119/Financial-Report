import 'dotenv/config';
import path from 'path';
import { fileURLToPath } from 'url';
import { readFileSync } from 'fs';

// Load .env from project root
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const envPath = path.join(__dirname, '..', '.env');
try {
  const env = readFileSync(envPath, 'utf8');
  for (const line of env.split('\n')) {
    const [k, ...v] = line.split('=');
    if (k && !process.env[k.trim()]) process.env[k.trim()] = v.join('=').trim();
  }
} catch {}

const JIRA_BASE  = (process.env.JIRA_BASE_URL || '').replace(/\/$/, '');
const JIRA_EMAIL = process.env.JIRA_EMAIL;
const JIRA_TOKEN = process.env.JIRA_API_TOKEN;
const authHeader = 'Basic ' + Buffer.from(`${JIRA_EMAIL}:${JIRA_TOKEN}`).toString('base64');

const ISSUE_KEY = process.argv[2] || 'ISM-81435';

async function jiraGet(path) {
  const res = await fetch(`${JIRA_BASE}${path}`, {
    headers: { Authorization: authHeader, Accept: 'application/json' },
  });
  if (!res.ok) throw new Error(`${path}: ${res.status} ${await res.text()}`);
  return res.json();
}

const ism = await jiraGet(`/rest/api/3/issue/${ISSUE_KEY}`);

console.log('\n========== ISM ISSUE ==========');
console.log(`Key:        ${ism.key}`);
console.log(`Summary:    ${ism.fields.summary}`);
console.log(`Status:     ${ism.fields.status?.name}`);
console.log(`Resolution: ${ism.fields.resolution?.name}`);
console.log(`cf[11077]:  ${JSON.stringify(ism.fields.customfield_11077)}`);

console.log('\n========== LINKED WORK ITEMS ==========');
const links = ism.fields.issuelinks || [];
if (links.length === 0) {
  console.log('No issuelinks found.');
} else {
  for (const l of links) {
    const related = l.outwardIssue || l.inwardIssue;
    const direction = l.outwardIssue ? 'outward' : 'inward';
    console.log(`  [${direction}] type="${l.type?.name}" key=${related?.key} summary="${related?.fields?.summary?.slice(0, 50)}"`);
  }
}

console.log('\n========== FIN TICKETS ==========');
const finLinks = links.filter(l => {
  const key = (l.outwardIssue || l.inwardIssue)?.key || '';
  return key.startsWith('FIN-');
});
if (finLinks.length === 0) {
  console.log('No FIN tickets found in issuelinks.');
} else {
  for (const l of finLinks) {
    const fin = l.outwardIssue || l.inwardIssue;
    console.log(`  ${fin.key} — ${fin.fields?.summary?.slice(0, 60)}`);
  }
}

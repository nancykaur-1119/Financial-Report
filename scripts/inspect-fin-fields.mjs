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

async function jiraGet(p) {
  const res = await fetch(`${JIRA_BASE}${p}`, { headers: { Authorization: authHeader, Accept: 'application/json' } });
  if (!res.ok) throw new Error(`${p}: ${res.status}`);
  return res.json();
}

// AR FIN ticket linked directly to ISM-81435
const AR_FIN = process.argv[2] || 'FIN-4351';

console.log(`\n========== AR FIN TICKET: ${AR_FIN} ==========`);
const arFin = await jiraGet(`/rest/api/3/issue/${AR_FIN}`);

console.log('\n--- Non-null custom fields ---');
for (const [key, val] of Object.entries(arFin.fields)) {
  if (val !== null && val !== undefined && key.startsWith('customfield_')) {
    const display = typeof val === 'object' ? JSON.stringify(val).slice(0, 80) : val;
    console.log(`  ${key}: ${display}`);
  }
}

console.log('\n--- Standard fields ---');
console.log(`  summary: ${arFin.fields.summary}`);
console.log(`  status:  ${arFin.fields.status?.name}`);
console.log(`  issuetype: ${arFin.fields.issuetype?.name}`);

console.log('\n--- issuelinks (AP FIN candidates) ---');
const links = arFin.fields.issuelinks || [];
if (links.length === 0) {
  console.log('  No issuelinks');
} else {
  for (const l of links) {
    const rel = l.outwardIssue || l.inwardIssue;
    const dir = l.outwardIssue ? 'outward' : 'inward';
    console.log(`  [${dir}] type="${l.type?.name}" key=${rel?.key} summary="${rel?.fields?.summary?.slice(0, 60)}"`);
  }
}

// Now inspect the first linked FIN inside AR FIN (AP FIN)
const apFin = links.find(l => {
  const k = (l.outwardIssue || l.inwardIssue)?.key || '';
  return k.startsWith('FIN-');
});
if (apFin) {
  const apKey = (apFin.outwardIssue || apFin.inwardIssue).key;
  console.log(`\n========== AP FIN TICKET: ${apKey} ==========`);
  const apData = await jiraGet(`/rest/api/3/issue/${apKey}`);
  console.log('\n--- Non-null custom fields ---');
  for (const [key, val] of Object.entries(apData.fields)) {
    if (val !== null && val !== undefined && key.startsWith('customfield_')) {
      const display = typeof val === 'object' ? JSON.stringify(val).slice(0, 80) : val;
      console.log(`  ${key}: ${display}`);
    }
  }
  console.log('\n--- Standard fields ---');
  console.log(`  summary: ${apData.fields.summary}`);
  console.log(`  status:  ${apData.fields.status?.name}`);
  console.log(`  issuetype: ${apData.fields.issuetype?.name}`);
} else {
  console.log('\nNo FIN link found inside AR FIN ticket.');
}

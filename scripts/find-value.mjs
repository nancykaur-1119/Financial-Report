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

const KEY = process.argv[2] || 'FIN-4388';
const SEARCH = process.argv[3] || '9556';

const res = await fetch(`${JIRA_BASE}/rest/api/3/issue/${KEY}`, { headers: { Authorization: authHeader, Accept: 'application/json' } });
const issue = await res.json();

// Find all fields containing the search value
console.log(`\nSearching for "${SEARCH}" in ${KEY} fields:\n`);
for (const [key, val] of Object.entries(issue.fields)) {
  if (val === null || val === undefined) continue;
  const str = JSON.stringify(val);
  if (str.includes(SEARCH)) {
    console.log(`  ${key}: ${str.slice(0, 120)}`);
  }
}

// Also show customfield_11529 specifically
console.log(`\ncustomfield_11529 = ${JSON.stringify(issue.fields.customfield_11529)}`);

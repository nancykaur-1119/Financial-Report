/**
 * Reselling Finance Report
 * Generates two Excel files:
 *   reselling-report-FY2025-26.xlsx  (1 Apr 2025 – 30 Mar 2026)
 *   reselling-report-FY2026-27.xlsx  (1 Apr 2026 – today)
 *
 * Columns:
 *   ISM Tkt ID | Company Name | License Amount ($) | OEM Purchase Amount ($) |
 *   AR FIN Tkt ID | Invoice Amount in USD (Excluding Tax) | OEM Quote Amount ($) |
 *   Actual Collection Amount in USD | AP FIN Tkt ID
 */

import 'dotenv/config';
import path from 'path';
import { fileURLToPath } from 'url';
import { readFileSync, writeFileSync } from 'fs';
import ExcelJS from 'exceljs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
try {
  const env = readFileSync(path.join(__dirname, '..', '.env'), 'utf8');
  for (const line of env.split('\n')) {
    const [k, ...v] = line.split('=');
    if (k && !process.env[k.trim()]) process.env[k.trim()] = v.join('=').trim();
  }
} catch {}

const JIRA_BASE  = (process.env.JIRA_BASE_URL || 'https://enreap.atlassian.net').replace(/\/$/, '');
const authHeader = 'Basic ' + Buffer.from(`${process.env.JIRA_EMAIL}:${process.env.JIRA_API_TOKEN}`).toString('base64');

// ── Field IDs ─────────────────────────────────────────────────────────────────
const F_COMPANY_NAME       = 'customfield_10225';  // Company Name
const F_LICENSE_AMOUNT     = 'customfield_10428';  // License Amount ($)
const F_INVOICE_USD_EXCL   = 'customfield_11995';  // Invoice Amount in USD (Excl. Tax)
const F_OEM_QUOTE          = 'customfield_11418';  // OEM Quote Amount ($)
const F_ACTUAL_COLLECTION  = 'customfield_12000';  // Actual Collection Amount in USD
const F_OEM_PURCHASE       = 'customfield_11364';  // OEM Purchase Amount ($)

const AR_FIN_FIELDS = [
  'summary', 'issuetype', 'issuelinks',
  F_COMPANY_NAME, F_LICENSE_AMOUNT, F_INVOICE_USD_EXCL,
  F_OEM_QUOTE, F_ACTUAL_COLLECTION, F_OEM_PURCHASE,
];

const num = (v) => (v === null || v === undefined ? '' : Number(v) || 0);

// ── Jira helpers ──────────────────────────────────────────────────────────────
async function jiraGet(p) {
  const res = await fetch(`${JIRA_BASE}${p}`, {
    headers: { Authorization: authHeader, Accept: 'application/json' },
  });
  if (!res.ok) throw new Error(`GET ${p}: ${res.status}`);
  return res.json();
}

async function jiraPost(p, body) {
  const res = await fetch(`${JIRA_BASE}${p}`, {
    method: 'POST',
    headers: { Authorization: authHeader, Accept: 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`POST ${p}: ${res.status}`);
  return res.json();
}

// ── Fetch all ISM issues for a date range ─────────────────────────────────────
async function fetchIsmIssues(fromDate, toDate) {
  const todayClause = toDate ? ` AND resolved <= "${toDate}"` : '';
  const jql = `project = ISM AND resolved >= "${fromDate}"${todayClause} AND status = "Closed" AND resolution IN ("Fixed", "done") AND cf[11077] = "R - Reselling" ORDER BY created DESC`;

  const allIssues = [];
  let nextPageToken;
  while (true) {
    const body = { jql, maxResults: 50, fields: ['summary', 'issuelinks', 'customfield_11077'] };
    if (nextPageToken) body.nextPageToken = nextPageToken;
    const page = await jiraPost('/rest/api/3/search/jql', body);
    allIssues.push(...(page.issues || []));
    if (!page.nextPageToken || !(page.issues || []).length) break;
    nextPageToken = page.nextPageToken;
  }
  return allIssues;
}

// ── Process one ISM → one or more rows ───────────────────────────────────────
async function processIsm(ismIssue) {
  const ismKey = ismIssue.key;

  // Direct FIN links on ISM = AR FIN tickets
  const arFinKeys = (ismIssue.fields.issuelinks || [])
    .map((l) => l.outwardIssue?.key || l.inwardIssue?.key)
    .filter((k) => k && k.startsWith('FIN-'));

  if (arFinKeys.length === 0) {
    return [{ ismKey, companyName: '', licenseAmount: '', arFinKey: '', invoiceUSD: '', oemQuote: '', actualCollection: '', apFinKey: '', oemPurchase: '' }];
  }

  const rows = [];

  for (const arFinKey of arFinKeys) {
    let arFin;
    try {
      arFin = await jiraGet(`/rest/api/3/issue/${arFinKey}?fields=${AR_FIN_FIELDS.join(',')}`);
    } catch (err) {
      console.warn(`  Skipping AR FIN ${arFinKey}: ${err.message}`);
      rows.push({ ismKey, companyName: '', licenseAmount: '', arFinKey, invoiceUSD: '', oemQuote: '', actualCollection: '', apFinKey: '', oemPurchase: '' });
      continue;
    }

    const companyName      = arFin.fields[F_COMPANY_NAME]  || '';
    const licenseAmount    = num(arFin.fields[F_LICENSE_AMOUNT]);
    const invoiceUSD       = num(arFin.fields[F_INVOICE_USD_EXCL]);
    const oemQuote         = num(arFin.fields[F_OEM_QUOTE]);
    const actualCollection = num(arFin.fields[F_ACTUAL_COLLECTION]);
    const oemPurchaseOnAR  = num(arFin.fields[F_OEM_PURCHASE]);

    // AP FIN = outward FIN link inside the AR FIN ticket
    const apFinLinks = (arFin.fields.issuelinks || [])
      .filter((l) => {
        const k = l.outwardIssue?.key || l.inwardIssue?.key || '';
        return k.startsWith('FIN-') && k !== arFinKey;
      });

    if (apFinLinks.length === 0) {
      rows.push({ ismKey, companyName, licenseAmount, arFinKey, invoiceUSD, oemQuote, actualCollection, apFinKey: '', oemPurchase: oemPurchaseOnAR });
    } else {
      for (const apLink of apFinLinks) {
        const apFinKey = apLink.outwardIssue?.key || apLink.inwardIssue?.key;
        let oemPurchase = oemPurchaseOnAR;
        try {
          const apFin = await jiraGet(`/rest/api/3/issue/${apFinKey}?fields=${F_OEM_PURCHASE}`);
          oemPurchase = num(apFin.fields[F_OEM_PURCHASE]);
        } catch (err) {
          console.warn(`  Skipping AP FIN ${apFinKey}: ${err.message}`);
        }
        rows.push({ ismKey, companyName, licenseAmount, arFinKey, invoiceUSD, oemQuote, actualCollection, apFinKey, oemPurchase });
      }
    }
  }

  return rows;
}

// ── Build Excel workbook ──────────────────────────────────────────────────────
async function buildExcel(rows, sheetName) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Finance Report';
  const ws = wb.addWorksheet(sheetName, { views: [{ state: 'frozen', ySplit: 1 }] });

  ws.columns = [
    { header: 'ISM Tkt ID',                              key: 'ismKey',           width: 16 },
    { header: 'Company Name',                             key: 'companyName',      width: 40 },
    { header: 'License Amount ($)',                       key: 'licenseAmount',    width: 20 },
    { header: 'OEM Purchase Amount ($)',                  key: 'oemPurchase',      width: 24 },
    { header: 'AR FIN Tkt ID',                            key: 'arFinKey',         width: 16 },
    { header: 'Invoice Amount in USD (Excluding Tax)',    key: 'invoiceUSD',       width: 32 },
    { header: 'OEM Quote Amount ($)',                     key: 'oemQuote',         width: 22 },
    { header: 'Actual Collection Amount in USD',          key: 'actualCollection', width: 30 },
    { header: 'AP FIN Tkt ID',                            key: 'apFinKey',         width: 16 },
  ];

  // Header styling
  const hdrRow = ws.getRow(1);
  hdrRow.height = 32;
  hdrRow.eachCell((cell) => {
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF172B4D' } };
    cell.font      = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border    = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
  });

  const numCols = new Set([3, 4, 6, 7, 8]); // 1-based col indices for numeric

  for (const row of rows) {
    const r = ws.addRow([
      row.ismKey,
      row.companyName,
      row.licenseAmount,
      row.oemPurchase,
      row.arFinKey,
      row.invoiceUSD,
      row.oemQuote,
      row.actualCollection,
      row.apFinKey,
    ]);
    r.height = 18;
    r.eachCell({ includeEmpty: true }, (cell, colNum) => {
      cell.border    = { top: { style: 'thin', color: { argb: 'FFDFE1E6' } }, bottom: { style: 'thin', color: { argb: 'FFDFE1E6' } }, left: { style: 'thin', color: { argb: 'FFDFE1E6' } }, right: { style: 'thin', color: { argb: 'FFDFE1E6' } } };
      cell.alignment = { vertical: 'middle' };
      if (numCols.has(colNum)) cell.alignment = { ...cell.alignment, horizontal: 'right' };
    });
    r.commit();
  }

  return wb;
}

// ── Main ──────────────────────────────────────────────────────────────────────
const today = new Date().toISOString().slice(0, 10);

const periods = [
  { label: 'FY2025-26', from: '2025-04-01', to: '2026-03-30', file: 'reselling-report-FY2025-26.xlsx' },
  { label: 'FY2026-27', from: '2026-04-01', to: today,         file: 'reselling-report-FY2026-27.xlsx' },
];

for (const period of periods) {
  console.log(`\n=== ${period.label}: ${period.from} → ${period.to} ===`);

  const ismIssues = await fetchIsmIssues(period.from, period.to);
  console.log(`Found ${ismIssues.length} ISM issues`);

  const allRows = [];
  for (let i = 0; i < ismIssues.length; i++) {
    const ism = ismIssues[i];
    process.stdout.write(`\r  Processing ${i + 1}/${ismIssues.length}: ${ism.key}         `);
    try {
      const rows = await processIsm(ism);
      allRows.push(...rows);
    } catch (err) {
      console.warn(`\n  Error processing ${ism.key}: ${err.message}`);
    }
  }
  console.log(`\n  → ${allRows.length} rows`);

  const wb = await buildExcel(allRows, period.label);
  const outPath = path.join(__dirname, '..', period.file);
  await wb.xlsx.writeFile(outPath);
  console.log(`  Saved: ${outPath}`);
}

console.log('\nDone.');

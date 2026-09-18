require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const express = require('express');
const cors    = require('cors');
const ExcelJS = require('exceljs');

const app = express();
app.use(cors());

const JIRA_BASE            = (process.env.JIRA_BASE_URL || '').replace(/\/$/, '');
const JIRA_EMAIL           = process.env.JIRA_EMAIL;
const JIRA_TOKEN           = process.env.JIRA_API_TOKEN;
const SERVICE_AMOUNT_FIELD    = process.env.SERVICE_AMOUNT_FIELD    || 'customfield_10383';
const EARNED_AMOUNT_FIELD     = process.env.EARNED_AMOUNT_FIELD     || 'customfield_12322';
const INVOICE_AMOUNT_FIELD    = process.env.INVOICE_AMOUNT_FIELD    || 'customfield_11366';
const CLAIM_DATE_FIELD        = process.env.CLAIM_DATE_FIELD        || 'customfield_16311';
const PROJECT_START_FIELD     = process.env.PROJECT_START_FIELD     || 'customfield_10800';
const PROJECT_END_FIELD       = process.env.PROJECT_END_FIELD       || 'customfield_10801';
const TOTAL_PLANNED_HRS_FIELD = process.env.TOTAL_PLANNED_HRS_FIELD || 'customfield_12458';
const BILLING_RATE_FIELD      = process.env.BILLING_RATE_FIELD      || 'customfield_10942';
const NON_ACCRUED_HRS_FIELD       = process.env.NON_ACCRUED_HRS_FIELD       || 'customfield_12459';
const INVOICE_USD_FIELD           = process.env.INVOICE_USD_FIELD           || 'customfield_11995';
const TOTAL_NON_ACCRUABLE_FIELD   = process.env.TOTAL_NON_ACCRUABLE_FIELD   || 'customfield_12326';
const SALES_INVOICE_RATE_FIELD    = process.env.SALES_INVOICE_RATE_FIELD    || 'customfield_12563';
const EARNED_MONTHLY_FIELD        = process.env.EARNED_MONTHLY_FIELD        || 'customfield_12321';
const PO_AMOUNT_FIELD             = process.env.PO_AMOUNT_FIELD             || 'customfield_12684';
const SALES_INVOICE_DATE_FIELD    = process.env.SALES_INVOICE_DATE_FIELD    || 'customfield_11377';

const F_COMPANY_NAME      = 'customfield_10225';
const F_LICENSE_AMOUNT    = 'customfield_10428';
const F_OEM_QUOTE         = 'customfield_11418';
const F_ACTUAL_COLLECTION = 'customfield_12000';
const F_OEM_PURCHASE      = 'customfield_11364';

// Extra Info report fields — verified against enreap.atlassian.net.
// NOTE: several of these are shared custom field IDs configured on both the ISM
// and FIN projects, and can hold DIFFERENT values on each side. Verified per-field
// which side is the real source (see extra-info-field-map.txt for the audit).
const F_BILLING_ENTITY     = 'customfield_11055';
const F_BILLING_FREQUENCY  = 'customfield_11415';
const F_PVR_AMOUNT         = 'customfield_15167';
const F_SERVICES_AMOUNT    = 'customfield_10383';
const F_MARKET_TEAM        = 'customfield_11373';
const F_SEGMENT_BU         = 'customfield_11384';
const F_SALES_INVOICE_NUMBER = 'customfield_11416'; // FIN only — blank until the ticket reaches the invoicing stage

// ISM-sourced (confirmed: populated on ISM, null/inconsistent or differing on FIN)
const F_REVENUE_STREAM      = 'customfield_11077';
const F_BILLING_TYPE        = 'customfield_11729';
const F_DF_AMOUNT           = 'customfield_11827';
const F_DR_AMOUNT           = 'customfield_11826';
const F_DR_REBATE_DF        = 'customfield_12028';
const F_FOREX_EXCHANGE_RATE = 'customfield_14721';
const F_EXCHANGE_RATE_TYPE  = 'customfield_14532'; // cascading select (parent + child)
const F_FOREX_MARKUP_INR    = 'customfield_12193';
const F_MARKET_TYPE         = 'customfield_11074';  // cascading select (parent + child)
const F_OEM                 = 'customfield_11092';  // cascading select (parent + child)
const F_GEO                 = 'customfield_11134';
const F_REGION_CODE         = 'customfield_11085';  // cascading select — NOT customfield_10789
                                                     // ("Region and Sub-Region"), which is unused/always null
const F_PROFITABILITY       = 'customfield_10479';  // confirmed to differ from FIN — per business decision, use ISM
// F_LICENSE_AMOUNT (customfield_10428) and INVOICE_AMOUNT_FIELD (customfield_11366) also
// confirmed to differ from FIN — per business decision, use ISM (declared above, reused here)

const EXTRA_INFO_ISM_FIELDS = [
  F_REVENUE_STREAM, F_BILLING_TYPE, F_DF_AMOUNT, F_DR_AMOUNT, F_DR_REBATE_DF,
  F_FOREX_EXCHANGE_RATE, F_EXCHANGE_RATE_TYPE, F_FOREX_MARKUP_INR, F_MARKET_TYPE, F_OEM, F_GEO, F_REGION_CODE,
  F_PROFITABILITY, F_LICENSE_AMOUNT, INVOICE_AMOUNT_FIELD,
];

const EXTRA_INFO_FIN_FIELDS = [
  F_COMPANY_NAME, F_BILLING_ENTITY, F_BILLING_FREQUENCY,
  F_OEM_PURCHASE, F_PVR_AMOUNT, F_SERVICES_AMOUNT,
  F_MARKET_TEAM, F_SEGMENT_BU, F_SALES_INVOICE_NUMBER,
];

// Select-list custom fields come back as {value|name} or a bare string
const pickValue = (raw) => (typeof raw === 'string' ? raw : raw?.value || raw?.name || '') || '';

// Cascading select fields (parent + child) — display value is "Parent - Child"
const pickCascading = (raw) => {
  if (raw == null) return '';
  if (typeof raw === 'string') return raw;
  const parent = raw.value || raw.name || '';
  const child  = raw.child?.value || raw.child?.name || '';
  return child ? `${parent} - ${child}` : parent;
};

if (!JIRA_BASE || !JIRA_EMAIL || !JIRA_TOKEN) {
  console.error('Missing env vars. Copy .env.example to .env and fill in your credentials.');
  process.exit(1);
}

const authHeader = 'Basic ' + Buffer.from(`${JIRA_EMAIL}:${JIRA_TOKEN}`).toString('base64');

async function jiraGet(path) {
  const res = await fetch(`${JIRA_BASE}${path}`, {
    headers: { Authorization: authHeader, Accept: 'application/json' },
  });
  if (!res.ok) throw new Error(`${path}: ${res.status} ${await res.text()}`);
  return res.json();
}

async function jiraPost(path, body) {
  const res = await fetch(`${JIRA_BASE}${path}`, {
    method: 'POST',
    headers: { Authorization: authHeader, Accept: 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`${path}: ${res.status} ${await res.text()}`);
  return res.json();
}

const num = (v) => (typeof v === 'number' ? v : Number(v) || 0);
const fmt = (n) => num(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtDate = (iso) => {
  if (!iso) return '—';
  try { return new Date(iso).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }); }
  catch { return iso; }
};

// ── Fetch ALL ISM issues (fully paginated) ──────────────────────────────────
async function fetchAllIsmIssues(fromDate = '2025-04-01', toDate = null) {
  const fields = ['summary', SERVICE_AMOUNT_FIELD, 'issuelinks', 'issuetype', 'created', 'status', 'resolution', 'resolutiondate', CLAIM_DATE_FIELD, PO_AMOUNT_FIELD, SALES_INVOICE_DATE_FIELD];
  const todayClause = toDate ? ` AND resolved <= "${toDate}"` : '';
  const jql    = `project = ISM AND resolved >= "${fromDate}"${todayClause} AND status = "Closed" AND resolution IN ("Fixed", "done") AND cf[11077] != "R - Reselling" ORDER BY cf[16311] DESC`;

                 
  
//   const jql = `
//   project = ISM
//   AND created >= "2025-09-01"
//   AND created < "2025-10-01"
//   ORDER BY created DESC
// `;
  const allIssues = [];
  let nextPageToken = undefined;

  while (true) {
    const body = { jql, maxResults: 50, fields };
    if (nextPageToken) body.nextPageToken = nextPageToken;

    const page   = await jiraPost('/rest/api/3/search/jql', body);
    const issues = page.issues || [];
    allIssues.push(...issues);
    console.log(`[fetch] ${allIssues.length} ISM issues fetched so far`);

    if (!page.nextPageToken || issues.length === 0) break;
    nextPageToken = page.nextPageToken;
  }

  return allIssues;
}

// ── Process one ISM issue → linked PRJ/AUS → linked FIN ──────────────────
async function processIsmIssue(ismIssue) {
  const serviceAmount = num(ismIssue.fields[SERVICE_AMOUNT_FIELD]);
  const linkedKeys = (ismIssue.fields.issuelinks || [])
    .map((l) => l.outwardIssue?.key || l.inwardIssue?.key)
    .filter((k) => k && (k.startsWith('PRJ-') || k.startsWith('AUS-')));

  let prjEarned = 0, ausEarned = 0, totalInvoiced = 0, totalInvoicedUSD = 0;
  const rows = [];

  for (const key of linkedKeys) {
    try {
      const linked    = await jiraGet(`/rest/api/3/issue/${key}`);
      const earned          = num(linked.fields[EARNED_AMOUNT_FIELD]);
      const issueType       = linked.fields.issuetype?.name ?? '';
      const projectStart    = linked.fields[PROJECT_START_FIELD] || null;
      const projectEnd      = linked.fields[PROJECT_END_FIELD]   || null;
      const totalPlannedHrs = num(linked.fields[TOTAL_PLANNED_HRS_FIELD]);
      const billingRate     = num(linked.fields[BILLING_RATE_FIELD]);
      const nonAccruedHrs       = num(linked.fields[NON_ACCRUED_HRS_FIELD]);
      const totalNonAccruable   = num(linked.fields[TOTAL_NON_ACCRUABLE_FIELD]);
      const salesInvoiceRate    = num(linked.fields[SALES_INVOICE_RATE_FIELD]);
      const earnedMonthly       = num(linked.fields[EARNED_MONTHLY_FIELD]);

      const finKeys = (linked.fields.issuelinks || [])
        .map((l) => l.outwardIssue?.key || l.inwardIssue?.key)
        .filter((k) => k?.startsWith('FIN-'));

      const finIssues = await Promise.all(
        finKeys.map(async (finKey) => {
          const fin = await jiraGet(`/rest/api/3/issue/${finKey}`);
          return { key: finKey, summary: fin.fields.summary, invoiced: num(fin.fields[INVOICE_AMOUNT_FIELD]), invoicedUSD: num(fin.fields[INVOICE_USD_FIELD]) };
        })
      );

      const rowInvoiced    = finIssues.reduce((s, f) => s + f.invoiced, 0);
      const rowInvoicedUSD = finIssues.reduce((s, f) => s + f.invoicedUSD, 0);
      if (key.startsWith('PRJ-')) prjEarned += earned;
      else if (key.startsWith('AUS-')) ausEarned += earned;
      totalInvoiced    += rowInvoiced;
      totalInvoicedUSD += rowInvoicedUSD;
      rows.push({ key, issueType, prjStatus: linked.fields.status?.name || null, summary: linked.fields.summary, earned, finIssues, rowInvoiced, rowInvoicedUSD, projectStart, projectEnd, totalPlannedHrs, billingRate, nonAccruedHrs, totalNonAccruable, salesInvoiceRate, earnedMonthly });
    } catch (err) {
      console.error(`  Skipping linked issue ${key}:`, err.message);
    }
  }

  return {
    ism: {
      key: ismIssue.key,
      summary: ismIssue.fields.summary,
      serviceAmount,
      created: ismIssue.fields.created,
      claimDate: ismIssue.fields[CLAIM_DATE_FIELD] || null,
      poAmount: num(ismIssue.fields[PO_AMOUNT_FIELD]),
      salesInvoiceDate: ismIssue.fields[SALES_INVOICE_DATE_FIELD] || null,
      status: ismIssue.fields.status?.name,
      resolution: ismIssue.fields.resolution?.name,
      resolutionDate: ismIssue.fields.resolutiondate || null,
    },
    rows, prjEarned, ausEarned, totalInvoiced, totalInvoicedUSD,
    pending: serviceAmount - totalInvoiced,
  };
}

// ── Build Excel workbook from reports ─────────────────────────────────────
async function buildExcel(reports) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'ISM Report';
  wb.created = new Date();

  const ws = wb.addWorksheet('ISM Report', { views: [{ state: 'frozen', ySplit: 1 }] });

  // 15 ISM cols + 20 PRJ/AUS cols + 4 FIN cols = 39 total
  ws.columns = [
    // ISM (1-15): ISM Issue, Issue Name, PO Date, PO Amount, Status, Resolved,
    //             Service Amount, Sales Invoice Creation Date, % Billed Till Date,
    //             Cummulative Billing on Project, Cummulative Billing INR,
    //             Billed Till Date, % Billing Done,
    //             Check for Revenue against PO, Check for billing against PO
    { width: 16 }, { width: 42 }, { width: 16 }, { width: 18 },
    { width: 14 }, { width: 16 }, { width: 18 }, { width: 26 }, { width: 20 },
    { width: 26 }, { width: 22 }, { width: 18 }, { width: 16 },
    { width: 26 }, { width: 26 },
    // PRJ/AUS (16-35)
    { width: 16 }, { width: 16 }, { width: 16 }, { width: 22 }, { width: 22 },
    { width: 20 }, { width: 14 }, { width: 20 }, { width: 22 },
    { width: 22 }, { width: 18 }, { width: 24 }, { width: 22 },
    { width: 26 }, { width: 22 }, { width: 18 }, { width: 22 },
    { width: 24 }, { width: 26 }, { width: 22 },
    // FIN (36-39)
    { width: 18 }, { width: 28 }, { width: 20 }, { width: 20 },
  ];

  const ISM_COLS = 15;
  const NUM_COLS = new Set([4,7,9,10,11,12,13,14,15, 21,22,23,24,25,26,27,28,29,30,31,32,33,34,35, 37,38,39]);

  const thinGrey = { style: 'thin',   color: { argb: 'FFDFE1E6' } };
  const medBlue  = { style: 'medium', color: { argb: 'FF0052CC' } };
  const border   = (last) => ({ top: thinGrey, left: thinGrey, right: thinGrey, bottom: last ? medBlue : thinGrey });
  const ismFill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF4F5F7' } };

  // Header row
  const hdr = ws.addRow([
    // ISM
    'ISM Issue', 'Issue Name', 'PO Date', 'PO Amount', ' ISM Status', 'Resolved',
    'Service Amount', 'Sales Invoice Creation Date',
    '% Billed Till Date', 'Cummulative Billing on Project', 'Cummulative Billing INR',
    'Billed Till Date', '% Billing Done', 'Check for Revenue against PO', 'Check for billing against PO',
    // PRJ/AUS
    'PRJ Status', 'Linked Issue', 'Issue Type', 'Planned Project Start Date', 'Planned Project End Date',
    'Efforts Planned Hours', 'Billing Rate', 'Total Earned Hours', 'Hours Consumed Till Date',
    '% Completion Till Date', 'Unbilled Till Date', 'Cummulative Consumed Hours',
    'Updated Remaining Hours', 'Updated Total Estimated Hours', '% Completetion Till Date',
    'Unbilled in INR', 'Revenue this month FX', 'Recovery Rate this month',
    'Revised Recovery Rate till Date', 'Total Earned Amount ($)',
    // FIN
    'Linked FIN Ticket', 'Invoiced Amount USD (Excl. Tax)', 'Billed This Month FX', 'Billed in Month INR',
  ]);
  hdr.height = 32;
  hdr.eachCell((cell) => {
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF172B4D' } };
    cell.font      = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border    = border(false);
  });

  let rowNum = 2;

  for (const report of reports) {
    const { ism, rows } = report;
    const startRow   = rowNum;
    const rowCount   = Math.max(rows.length, 1);
    const svc        = ism.serviceAmount;
    const totalEarned = report.prjEarned + report.ausEarned;

    const dataRows = rows.length === 0 ? [null] : rows;

    for (let i = 0; i < dataRows.length; i++) {
      const row    = dataRows[i];
      const isLast = i === dataRows.length - 1;
      const empty  = row === null;

      const finKey   = !empty && row.finIssues?.length ? row.finIssues.map(f => f.key).join('\n')              : '—';
      const finINR   = !empty && row.finIssues?.length ? row.finIssues.map(f => fmt(f.invoiced)).join('\n')    : '—';
      const finUSD   = !empty && row.finIssues?.length ? row.finIssues.map(f => fmt(f.invoicedUSD)).join('\n') : '—';

      const d  = (v)       => empty ? '—' : (v  ? fmt(v)  : '—');
      const dc = (cond, v) => empty ? '—' : (cond ? fmt(v) : '—');

      const r = ws.addRow([
        // ISM — written only on first sub-row; cells merged below
        i === 0 ? ism.key                                                          : '',
        i === 0 ? ism.summary                                                      : '',
        i === 0 ? fmtDate(ism.claimDate)                                           : '',
        i === 0 ? (ism.poAmount ? fmt(ism.poAmount) : '—')                        : '',
        i === 0 ? (ism.status || '—')                                              : '',
        i === 0 ? fmtDate(ism.resolutionDate)                                      : '',
        i === 0 ? fmt(svc)                                                         : '',
        i === 0 ? fmtDate(ism.salesInvoiceDate)                                    : '',
        i === 0 ? (svc ? fmt(report.totalInvoiced / svc)         : '—')           : '',
        i === 0 ? fmt(report.totalInvoicedUSD)                                     : '',
        i === 0 ? fmt(report.totalInvoiced)                                        : '',
        i === 0 ? fmt(report.totalInvoicedUSD)                                     : '',
        i === 0 ? (svc ? fmt(report.totalInvoicedUSD / svc)      : '—')           : '',
        i === 0 ? (svc ? fmt(totalEarned / svc)                  : '—')           : '',
        i === 0 ? (svc ? fmt(report.totalInvoicedUSD / svc)      : '—')           : '',
        // PRJ/AUS
        empty ? '—' : (row.prjStatus || '—'),
        empty ? '—' : row.key,
        empty ? '—' : (row.issueType || '—'),
        empty ? '—' : fmtDate(row.projectStart),
        empty ? '—' : fmtDate(row.projectEnd),
        d(row?.totalPlannedHrs),
        d(row?.billingRate),
        empty ? '—' : fmt(row.earned),
        dc(row?.billingRate,     row?.earned / row?.billingRate),
        dc(row?.totalPlannedHrs, row?.earned / row?.totalPlannedHrs),
        d(row?.nonAccruedHrs),
        empty ? '—' : fmt(row.earned),
        dc(row?.totalPlannedHrs, row?.totalPlannedHrs - row?.earned),
        dc(row?.totalPlannedHrs, row?.totalPlannedHrs - row?.earned),
        dc(row?.totalPlannedHrs, row?.earned / row?.totalPlannedHrs),
        (empty || !(row.nonAccruedHrs && row.billingRate))                   ? '—' : fmt(row.nonAccruedHrs * row.billingRate),
        d(row?.earnedMonthly),
        (empty || !(row.rowInvoicedUSD + row.totalNonAccruable))             ? '—' : fmt(row.rowInvoicedUSD / (row.rowInvoicedUSD + row.totalNonAccruable)),
        (empty || !(row.rowInvoicedUSD + row.totalNonAccruable))             ? '—' : fmt(row.rowInvoicedUSD / (row.rowInvoicedUSD + row.totalNonAccruable)),
        empty ? '—' : fmt(row.earned),
        // FIN
        finKey,
        finINR,
        finUSD,
        finINR,
      ]);

      r.height = 20;
      r.eachCell({ includeEmpty: true }, (cell, col) => {
        cell.border    = border(isLast);
        cell.alignment = { vertical: 'middle', wrapText: true };
        if (col <= ISM_COLS)      cell.fill = ismFill;
        if (NUM_COLS.has(col))    cell.alignment = { ...cell.alignment, horizontal: 'right' };
      });
      r.commit();
      rowNum++;
    }

    if (rowCount > 1) {
      for (let col = 1; col <= ISM_COLS; col++) {
        ws.mergeCells(startRow, col, rowNum - 1, col);
        const cell     = ws.getCell(startRow, col);
        cell.fill      = ismFill;
        cell.border    = border(true);
        cell.alignment = { vertical: 'middle', horizontal: NUM_COLS.has(col) ? 'right' : 'left', wrapText: true };
      }
    }
  }

  return wb;
}

// ── API: single ISM issue ─────────────────────────────────────────────────
app.get('/api/ism-report/:issueKey', async (req, res) => {
  try {
    const ism    = await jiraGet(`/rest/api/3/issue/${req.params.issueKey}`);
    const report = await processIsmIssue(ism);
    console.log(`[ism-report] ${ism.key} → ${report.rows.length} linked rows`);
    res.json(report);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── API: all ISM issues (JSON) ────────────────────────────────────────────
app.get('/api/ism-projects', async (_req, res) => {
  try {
    const allIssues = await fetchAllIsmIssues();
    console.log(`[ism-projects] processing ${allIssues.length} issues`);

    const reports = [];
    for (const ism of allIssues) {
      console.log(`[ism-projects] processing ${ism.key}`);
      reports.push(await processIsmIssue(ism));
    }

    console.log(`[ism-projects] done — ${reports.length} reports`);
    res.json({ reports, total: reports.length });
  } catch (err) {
    console.error('[ism-projects] ERROR:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── API: stream ISM reports one-by-one via Server-Sent Events ────────────
app.get('/api/ism-projects/stream', async (req, res) => {
  res.setHeader('Content-Type',  'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection',    'keep-alive');
  const fromDate = req.query.from || '2025-04-01';
  const toDate   = req.query.to   || null;
  res.flushHeaders();

  const send = (event, data) =>
    res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);

  try {
    const allIssues = await fetchAllIsmIssues(fromDate, toDate);
    send('total', { total: allIssues.length });
    console.log(`[stream] ${allIssues.length} issues — streaming reports`);

    for (const ism of allIssues) {
      const report = await processIsmIssue(ism);
      console.log(`[stream] sent ${ism.key}`);
      send('report', report);
    }

    send('done', {});
    res.end();
  } catch (err) {
    console.error('[stream] ERROR:', err.message);
    send('error', { error: err.message });
    res.end();
  }
});

// ── API: download Excel ───────────────────────────────────────────────────
app.get('/api/ism-projects/export', async (req, res) => {
  try {
    const fromDate = req.query.from || '2025-04-01';
    const toDate   = req.query.to   || null;
    const allIssues = await fetchAllIsmIssues(fromDate, toDate);
    console.log(`[export] processing ${allIssues.length} issues`);

    const reports = [];
    for (const ism of allIssues) {
      reports.push(await processIsmIssue(ism));
    }

    const wb   = await buildExcel(reports);
    const date = new Date().toISOString().slice(0, 10);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="ism-report-${date}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
    console.log(`[export] done — ${reports.length} rows`);
  } catch (err) {
    console.error('[export] ERROR:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── API: reselling flat rows ─────────────────────────────────────────────
app.get('/api/reselling-projects', async (req, res) => {
  const fromDate = req.query.from || '2025-04-01';
  const toDate   = req.query.to   || null;
  const todayClause = toDate ? ` AND resolved <= "${toDate}"` : '';
  const jql = `project = ISM AND resolved >= "${fromDate}"${todayClause} AND status = "Closed" AND resolution IN ("Fixed", "done") AND cf[11077] = "R - Reselling" ORDER BY cf[16311] DESC`;
  const AR_FIELDS = ['summary', 'issuetype', 'issuelinks', F_COMPANY_NAME, F_LICENSE_AMOUNT, INVOICE_USD_FIELD, F_OEM_QUOTE, F_ACTUAL_COLLECTION, F_OEM_PURCHASE].join(',');

  try {
    const allIsmIssues = [];
    let nextPageToken;
    while (true) {
      const body = { jql, maxResults: 50, fields: ['summary', 'issuelinks'] };
      if (nextPageToken) body.nextPageToken = nextPageToken;
      const page = await jiraPost('/rest/api/3/search/jql', body);
      allIsmIssues.push(...(page.issues || []));
      if (!page.nextPageToken || !(page.issues || []).length) break;
      nextPageToken = page.nextPageToken;
    }
    console.log(`[reselling] ${allIsmIssues.length} ISM issues`);

    const allRows = [];
    for (const ism of allIsmIssues) {
      const ismKey    = ism.key;
      const arFinKeys = (ism.fields.issuelinks || [])
        .map(l => l.outwardIssue?.key || l.inwardIssue?.key)
        .filter(k => k && k.startsWith('FIN-'));

      if (arFinKeys.length === 0) {
        allRows.push({ ismKey, companyName: '', licenseAmount: '', oemPurchase: '', arFinKey: '', invoiceUSD: '', oemQuote: '', actualCollection: '', apFinKey: '' });
        continue;
      }

      for (const arFinKey of arFinKeys) {
        try {
          const arFin            = await jiraGet(`/rest/api/3/issue/${arFinKey}?fields=${AR_FIELDS}`);
          const companyName      = arFin.fields[F_COMPANY_NAME]  || '';
          const licenseAmount    = num(arFin.fields[F_LICENSE_AMOUNT]);
          const invoiceUSD       = num(arFin.fields[INVOICE_USD_FIELD]);
          const oemQuote         = num(arFin.fields[F_OEM_QUOTE]);
          const actualCollection = num(arFin.fields[F_ACTUAL_COLLECTION]);
          const oemPurchaseOnAR  = num(arFin.fields[F_OEM_PURCHASE]);

          const apFinLinks = (arFin.fields.issuelinks || [])
            .filter(l => { const k = l.outwardIssue?.key || l.inwardIssue?.key || ''; return k.startsWith('FIN-') && k !== arFinKey; });

          if (apFinLinks.length === 0) {
            allRows.push({ ismKey, companyName, licenseAmount, arFinKey, invoiceUSD, oemQuote, actualCollection, apFinKey: '', oemPurchase: oemPurchaseOnAR });
          } else {
            for (const apLink of apFinLinks) {
              const apFinKey  = apLink.outwardIssue?.key || apLink.inwardIssue?.key;
              let oemPurchase = oemPurchaseOnAR;
              try {
                const apFin = await jiraGet(`/rest/api/3/issue/${apFinKey}?fields=${F_OEM_PURCHASE}`);
                oemPurchase  = num(apFin.fields[F_OEM_PURCHASE]);
              } catch { /* keep AR value */ }
              allRows.push({ ismKey, companyName, licenseAmount, arFinKey, invoiceUSD, oemQuote, actualCollection, apFinKey, oemPurchase });
            }
          }
        } catch (err) {
          console.warn(`  Skipping AR FIN ${arFinKey}:`, err.message);
          allRows.push({ ismKey, companyName: '', licenseAmount: '', oemPurchase: '', arFinKey, invoiceUSD: '', oemQuote: '', actualCollection: '', apFinKey: '' });
        }
      }
    }

    console.log(`[reselling] ${allRows.length} rows`);
    res.json({ rows: allRows, total: allRows.length });
  } catch (err) {
    console.error('[reselling] ERROR:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── Shared: ISM → linked FIN ticket, full Extra Info field set ───────────
// onRow(row) is called once per output row, as soon as it's built, so callers
// can either stream it immediately (SSE) or collect it into an array (export).
// onTotal(n) is called once, after the ISM ticket list is known, with the
// count of ISM tickets (not output rows — one ISM ticket can yield 0+ rows).
async function fetchExtraInfoRows(fromDate, toDate, onRow, onTotal) {
  const todayClause = toDate ? ` AND cf[16311] <= "${toDate}"` : '';
  // Only R-Reselling ISM tickets (direct FIN link on the ISM) — per confirmed decision.
  // No status/resolution filter — includes matching ISM tickets regardless of status.
  // Filtered by Claim Date (cf[16311]), not Resolved date — per confirmed decision.
  const jql = `project = ISM AND cf[16311] >= "${fromDate}"${todayClause} AND cf[11077] = "R - Reselling" ORDER BY cf[16311] DESC`;

  const allIsmIssues = [];
  let nextPageToken;
  while (true) {
    const body = { jql, maxResults: 50, fields: ['summary', 'status', 'issuelinks', CLAIM_DATE_FIELD, ...EXTRA_INFO_ISM_FIELDS] };
    if (nextPageToken) body.nextPageToken = nextPageToken;
    const page = await jiraPost('/rest/api/3/search/jql', body);
    allIsmIssues.push(...(page.issues || []));
    if (!page.nextPageToken || !(page.issues || []).length) break;
    nextPageToken = page.nextPageToken;
  }
  console.log(`[extra-info] ${allIsmIssues.length} ISM issues`);
  if (onTotal) onTotal(allIsmIssues.length);

  for (const ism of allIsmIssues) {
    const ismF = ism.fields;
    const base = {
      issueKey: ism.key,
      issueId: ism.id,
      summary: ismF.summary || '',
      status: ismF.status?.name || '',
      claimDate: ismF[CLAIM_DATE_FIELD] || null,
      // ISM-sourced fields (see EXTRA_INFO_ISM_FIELDS comment above)
      billingType:        pickValue(ismF[F_BILLING_TYPE]),
      revenueStream:      pickValue(ismF[F_REVENUE_STREAM]),
      dfAmount:           num(ismF[F_DF_AMOUNT]),
      drAmount:           num(ismF[F_DR_AMOUNT]),
      drRebateDfAmount:   num(ismF[F_DR_REBATE_DF]),
      forexExchangeRate:  pickValue(ismF[F_FOREX_EXCHANGE_RATE]),
      exchangeRateType:   pickCascading(ismF[F_EXCHANGE_RATE_TYPE]),
      forexMarkUpInr:     num(ismF[F_FOREX_MARKUP_INR]),
      marketType:         pickCascading(ismF[F_MARKET_TYPE]),
      oem:                pickCascading(ismF[F_OEM]),
      geo:                pickValue(ismF[F_GEO]),
      regionSubRegion:    pickCascading(ismF[F_REGION_CODE]),
      profitability:      num(ismF[F_PROFITABILITY]),
      licenseAmount:      num(ismF[F_LICENSE_AMOUNT]),
      invoiceAmountLocal: num(ismF[INVOICE_AMOUNT_FIELD]),
    };

    // Reselling pattern: FIN linked directly on the ISM ticket.
    let finKeys = (ismF.issuelinks || [])
      .map(l => l.outwardIssue?.key || l.inwardIssue?.key)
      .filter(k => k && k.startsWith('FIN-'));

    // Service pattern: no direct FIN link — FIN is reached via linked PRJ/AUS tickets.
    if (finKeys.length === 0) {
      const prjAusKeys = (ismF.issuelinks || [])
        .map(l => l.outwardIssue?.key || l.inwardIssue?.key)
        .filter(k => k && (k.startsWith('PRJ-') || k.startsWith('AUS-')));

      const finKeySet = new Set();
      for (const prjAusKey of prjAusKeys) {
        try {
          const linked = await jiraGet(`/rest/api/3/issue/${prjAusKey}?fields=issuelinks`);
          for (const l of (linked.fields.issuelinks || [])) {
            const k = l.outwardIssue?.key || l.inwardIssue?.key;
            if (k && k.startsWith('FIN-')) finKeySet.add(k);
          }
        } catch (err) {
          console.warn(`  Skipping linked ${prjAusKey}:`, err.message);
        }
      }
      finKeys = [...finKeySet];
    }

    if (finKeys.length === 0) {
      onRow({ ...base, finKey: '' });
      continue;
    }

    for (const finKey of finKeys) {
      try {
        const fin = await jiraGet(`/rest/api/3/issue/${finKey}?fields=${EXTRA_INFO_FIN_FIELDS.join(',')}`);
        const f = fin.fields;
        onRow({
          ...base,
          finKey,
          companyName:        pickValue(f[F_COMPANY_NAME]),
          billingEntity:      pickValue(f[F_BILLING_ENTITY]),
          billingFrequency:   pickValue(f[F_BILLING_FREQUENCY]),
          oemPurchaseAmount:  num(f[F_OEM_PURCHASE]),
          pvrAmount:          num(f[F_PVR_AMOUNT]),
          servicesAmount:     num(f[F_SERVICES_AMOUNT]),
          marketTeam:         pickValue(f[F_MARKET_TEAM]),
          segmentBu:          pickValue(f[F_SEGMENT_BU]),
          salesInvoiceNumber: pickValue(f[F_SALES_INVOICE_NUMBER]),
          // licenseAmount, profitability, invoiceAmountLocal intentionally NOT set here —
          // they come from `base` (ISM-sourced), per confirmed business decision.
        });
      } catch (err) {
        console.warn(`  Skipping FIN ${finKey}:`, err.message);
        onRow({ ...base, finKey });
      }
    }
  }
}

// ── API: extra info flat rows (ISM → linked FIN ticket, full field set) ──
// Streamed via SSE so rows appear as soon as each FIN ticket is fetched,
// instead of waiting for the entire batch before rendering anything.
app.get('/api/extra-info-projects/stream', async (req, res) => {
  res.setHeader('Content-Type',  'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection',    'keep-alive');
  res.flushHeaders();

  const send = (event, data) =>
    res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);

  const fromDate = req.query.from || '2025-04-01';
  const toDate   = req.query.to   || null;

  try {
    await fetchExtraInfoRows(
      fromDate, toDate,
      (row) => send('row', row),
      (total) => send('total', { total }),
    );
    send('done', {});
    res.end();
  } catch (err) {
    console.error('[extra-info] ERROR:', err.message);
    send('error', { error: err.message });
    res.end();
  }
});

// ── API: extra info Excel export ──────────────────────────────────────────
const EXTRA_INFO_COLUMNS = [
  { header: 'Issue key',                                          key: 'issueKey' },
  { header: 'Issue id',                                           key: 'issueId' },
  { header: 'FIN',                                                key: 'finKey' },
  { header: 'Claim Date',                                         key: 'claimDate',          isDate: true },
  { header: 'Company Name',                                       key: 'companyName' },
  { header: 'Summary',                                            key: 'summary' },
  { header: 'Billing Entity',                                     key: 'billingEntity' },
  { header: 'Billing Frequency',                                  key: 'billingFrequency' },
  { header: 'Billing Type',                                       key: 'billingType' },
  { header: 'Revenue Stream',                                     key: 'revenueStream' },
  { header: 'Segment BU',                                         key: 'segmentBu' },
  { header: 'Market type',                                        key: 'marketType' },
  { header: 'Market Team',                                        key: 'marketTeam' },
  { header: 'Geo',                                                key: 'geo' },
  { header: 'Region and Sub-Region',                              key: 'regionSubRegion' },
  { header: 'Exchange Rate Type',                                 key: 'exchangeRateType' },
  { header: 'Forex Exchange Rate',                                key: 'forexExchangeRate' },
  { header: 'Forex Mark Up (INR)',                                key: 'forexMarkUpInr',     isNum: true },
  { header: 'OEM',                                                key: 'oem' },
  { header: 'License Amount ($)',                                 key: 'licenseAmount',      isNum: true },
  { header: 'OEM Purchase Amount ($)',                            key: 'oemPurchaseAmount',  isNum: true },
  { header: 'Profitability',                                      key: 'profitability',      isNum: true },
  { header: 'DF Amount ($)',                                      key: 'dfAmount',           isNum: true },
  { header: 'DR Amount ($)',                                      key: 'drAmount',           isNum: true },
  { header: 'DR/Rebate/DF Amount ($)',                            key: 'drRebateDfAmount',   isNum: true },
  { header: 'PVR Amount ($)',                                     key: 'pvrAmount',          isNum: true },
  { header: 'Services Amount ($)',                                key: 'servicesAmount',     isNum: true },
  { header: 'Invoice Amount in local currency (Without Tax)',     key: 'invoiceAmountLocal', isNum: true },
  { header: 'Sales Invoice Number',                               key: 'salesInvoiceNumber' },
  { header: 'Status',                                             key: 'status' },
];

app.get('/api/extra-info-projects/export', async (req, res) => {
  try {
    const fromDate = req.query.from || '2025-04-01';
    const toDate   = req.query.to   || null;

    const rows = [];
    await fetchExtraInfoRows(fromDate, toDate, (row) => rows.push(row));
    console.log(`[extra-info-export] ${rows.length} rows`);

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Finance Report';
    wb.created = new Date();
    const ws = wb.addWorksheet('Extra Info', { views: [{ state: 'frozen', ySplit: 1 }] });

    ws.columns = EXTRA_INFO_COLUMNS.map(c => ({ header: c.header, key: c.key, width: Math.max(16, Math.min(38, c.header.length + 4)) }));

    const hdrRow = ws.getRow(1);
    hdrRow.height = 32;
    hdrRow.eachCell(cell => {
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF172B4D' } };
      cell.font      = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border    = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
    });

    for (const row of rows) {
      const values = {};
      for (const col of EXTRA_INFO_COLUMNS) {
        const v = row[col.key];
        if (col.isDate) values[col.key] = fmtDate(v);
        else if (col.isNum) values[col.key] = v === '' || v === null || v === undefined ? '' : num(v);
        else values[col.key] = v ?? '';
      }
      const r = ws.addRow(values);
      r.height = 18;
      r.eachCell({ includeEmpty: true }, (cell, colNum) => {
        cell.border    = { top: { style: 'thin', color: { argb: 'FFDFE1E6' } }, bottom: { style: 'thin', color: { argb: 'FFDFE1E6' } }, left: { style: 'thin', color: { argb: 'FFDFE1E6' } }, right: { style: 'thin', color: { argb: 'FFDFE1E6' } } };
        cell.alignment = { vertical: 'middle', wrapText: true };
        if (EXTRA_INFO_COLUMNS[colNum - 1]?.isNum) cell.alignment = { ...cell.alignment, horizontal: 'right' };
      });
      r.commit();
    }

    const date = new Date().toISOString().slice(0, 10);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="extra-info-report-${date}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
    console.log(`[extra-info-export] done — ${rows.length} rows`);
  } catch (err) {
    console.error('[extra-info-export] ERROR:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── API: debug — PRJ project status values ───────────────────────────────
app.get('/api/debug/prj-statuses', async (_req, res) => {
  try {
    const data = await jiraGet('/rest/api/3/project/PRJ/statuses');
    const statuses = [...new Set(data.flatMap(issueType => issueType.statuses.map(s => s.name)))].sort();
    res.json({ project: 'PRJ', statuses });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── API: check export — R-Reselling issues + Non-zero Service Amount ────────
app.get('/api/ism-projects/export-check', async (_req, res) => {
  try {
    const checkFields = ['summary', 'status', CLAIM_DATE_FIELD, SERVICE_AMOUNT_FIELD, 'customfield_11077'];

    // Sheet 1: Revenue Stream = "R - Reselling" or "R-Reselling"
    const jql1 = `project = ISM AND cf[16311] >= "2025-03-01" AND cf[11077] IN ("R - Reselling", "R-Reselling") ORDER BY cf[16311] DESC`;
    // Sheet 2: Service Amount > 0, show Revenue Stream
    const jql2 = `project = ISM AND cf[16311] >= "2025-03-01" AND cf[10383] > 0 ORDER BY cf[16311] DESC`;

    const fetchSheet = async (jql) => {
      const issues = [];
      let nextPageToken;
      while (true) {
        const body = { jql, maxResults: 50, fields: checkFields };
        if (nextPageToken) body.nextPageToken = nextPageToken;
        const page = await jiraPost('/rest/api/3/search/jql', body);
        issues.push(...(page.issues || []));
        if (!page.nextPageToken || !page.issues?.length) break;
        nextPageToken = page.nextPageToken;
      }
      return issues;
    };

    const [resellingIssues, nonZeroIssues] = await Promise.all([fetchSheet(jql1), fetchSheet(jql2)]);

    const wb = new ExcelJS.Workbook();

    const addSheet = (name, issues) => {
      const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
      ws.columns = [{ width: 16 }, { width: 42 }, { width: 16 }, { width: 20 }, { width: 20 }, { width: 20 }];
      const hdr = ws.addRow(['ISM Issue', 'Issue Name', 'PO Date', 'Service Amount', 'Revenue Stream', 'Status']);
      hdr.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF172B4D' } };
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });
      for (const issue of issues) {
        const f = issue.fields;
        ws.addRow([
          issue.key,
          f.summary || '',
          fmtDate(f[CLAIM_DATE_FIELD]),
          f[SERVICE_AMOUNT_FIELD] ? fmt(f[SERVICE_AMOUNT_FIELD]) : '—',
          f['customfield_11077']?.value || f['customfield_11077'] || '—',
          f.status?.name || '—',
        ]);
      }
    };

    addSheet('R-Reselling Issues', resellingIssues);
    addSheet('Non-Zero Service Amount', nonZeroIssues);

    const date = new Date().toISOString().slice(0, 10);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="ism-check-${date}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
    console.log(`[export-check] R-Reselling: ${resellingIssues.length}, Non-zero SA: ${nonZeroIssues.length}`);
  } catch (err) {
    console.error('[export-check] ERROR:', err.message);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.API_PORT || 3001;
app.listen(PORT, () => {
  console.log(`API server → http://localhost:${PORT}`);
  console.log(`Jira instance: ${JIRA_BASE}`);
});

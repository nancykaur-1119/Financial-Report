const ForgeResolver = require('@forge/resolver');
const Resolver = ForgeResolver.default || ForgeResolver;
const ExcelJS = require('exceljs');

const resolver = new Resolver();

const JIRA_BASE            = process.env.JIRA_BASE_URL    || 'https://enreap.atlassian.net';
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
const REVENUE_STREAM_FIELD        = process.env.REVENUE_STREAM_FIELD        || 'customfield_11077';

// Reselling-specific fields
const F_COMPANY_NAME      = 'customfield_10225';
const F_LICENSE_AMOUNT    = 'customfield_10428';
const F_OEM_QUOTE         = 'customfield_11418';
const F_ACTUAL_COLLECTION = 'customfield_12000';
const F_OEM_PURCHASE      = 'customfield_11364';

// Extra Info report fields — see extra-info-field-map.txt for the full audit.
// Several of these are shared custom field IDs on both ISM and FIN, verified
// per-field which side is authoritative.
const F_BILLING_ENTITY       = 'customfield_11055';
const F_BILLING_FREQUENCY    = 'customfield_11415';
const F_PVR_AMOUNT           = 'customfield_15167';
const F_SERVICES_AMOUNT      = 'customfield_10383';
const F_MARKET_TEAM          = 'customfield_11373';
const F_SEGMENT_BU           = 'customfield_11384';
const F_SALES_INVOICE_NUMBER = 'customfield_11416'; // FIN only — blank until the ticket reaches the invoicing stage

// ISM-sourced (confirmed: populated on ISM, null/inconsistent or differing on FIN)
const F_BILLING_TYPE         = 'customfield_11729';
const F_DF_AMOUNT            = 'customfield_11827';
const F_DR_AMOUNT            = 'customfield_11826';
const F_DR_REBATE_DF         = 'customfield_12028';
const F_FOREX_EXCHANGE_RATE  = 'customfield_14721';
const F_EXCHANGE_RATE_TYPE   = 'customfield_14532'; // cascading select (parent + child)
const F_FOREX_MARKUP_INR     = 'customfield_12193';
const F_MARKET_TYPE          = 'customfield_11074';  // cascading select (parent + child)
const F_OEM                  = 'customfield_11092';  // cascading select (parent + child)
const F_GEO                  = 'customfield_11134';
const F_REGION_CODE          = 'customfield_11085';  // cascading select — NOT customfield_10789
const F_PROFITABILITY        = 'customfield_10479';  // confirmed to differ from FIN — use ISM
// F_LICENSE_AMOUNT and INVOICE_AMOUNT_FIELD also confirmed to differ from FIN — use ISM

const EXTRA_INFO_ISM_FIELDS = [
  REVENUE_STREAM_FIELD, F_BILLING_TYPE, F_DF_AMOUNT, F_DR_AMOUNT, F_DR_REBATE_DF,
  F_FOREX_EXCHANGE_RATE, F_EXCHANGE_RATE_TYPE, F_FOREX_MARKUP_INR, F_MARKET_TYPE, F_OEM, F_GEO, F_REGION_CODE,
  F_PROFITABILITY, F_LICENSE_AMOUNT, INVOICE_AMOUNT_FIELD,
];

const EXTRA_INFO_FIN_FIELDS = [
  F_COMPANY_NAME, F_BILLING_ENTITY, F_BILLING_FREQUENCY,
  F_OEM_PURCHASE, F_PVR_AMOUNT, F_SERVICES_AMOUNT,
  F_MARKET_TEAM, F_SEGMENT_BU, F_SALES_INVOICE_NUMBER,
];

// Cascading select fields (parent + child) — display value is "Parent - Child"
const pickCascading = (raw) => {
  if (raw == null) return '';
  if (typeof raw === 'string') return raw;
  const parent = raw.value || raw.name || '';
  const child  = raw.child?.value || raw.child?.name || '';
  return child ? `${parent} - ${child}` : parent;
};

const authHeader = 'Basic ' + Buffer.from(`${JIRA_EMAIL}:${JIRA_TOKEN}`).toString('base64');
const num = (v) => (typeof v === 'number' ? v : Number(v) || 0);
// Select-list custom fields come back as {value|name} or a bare string
const pickValue = (raw) => (typeof raw === 'string' ? raw : raw?.value || raw?.name || '') || '';


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

async function processIsmIssue(ismIssue) {
  const serviceAmount = num(ismIssue.fields[SERVICE_AMOUNT_FIELD]);
  const revenueStreamRaw = ismIssue.fields[REVENUE_STREAM_FIELD];
  const revenueStream = typeof revenueStreamRaw === 'string' ? revenueStreamRaw : (revenueStreamRaw?.value || revenueStreamRaw?.name || '');

  // Direct FIN links on the ISM itself (R-Reselling pattern)
  const directFinKeys = (ismIssue.fields.issuelinks || [])
    .map((l) => l.outwardIssue?.key || l.inwardIssue?.key)
    .filter((k) => k && k.startsWith('FIN-'));

  // Detect R-Reselling by field value OR by direct FIN links on ISM
  const isReselling = revenueStream === 'R - Reselling' || directFinKeys.length > 0;

  let rows = [], prjEarned = 0, ausEarned = 0, totalInvoiced = 0, totalInvoicedUSD = 0;

  if (isReselling) {
    // R - Reselling: FIN tickets are linked directly on the ISM issue.
    // Use issuelinks from JQL result directly — no re-fetch needed.
    const finIssues = await Promise.all(
      directFinKeys.map(async (finKey) => {
        try {
          const fin = await jiraGet(`/rest/api/3/issue/${finKey}`);
          return { key: finKey, summary: fin.fields.summary, invoiced: num(fin.fields[INVOICE_AMOUNT_FIELD]), invoicedUSD: num(fin.fields[INVOICE_USD_FIELD]) };
        } catch (err) {
          console.error('Skipping FIN issue', finKey, err.message);
          return null;
        }
      })
    );
    const validFin = finIssues.filter(Boolean);
    const rowInvoiced    = validFin.reduce((s, f) => s + f.invoiced, 0);
    const rowInvoicedUSD = validFin.reduce((s, f) => s + f.invoicedUSD, 0);
    totalInvoiced    = rowInvoiced;
    totalInvoicedUSD = rowInvoicedUSD;
    rows = [{ key: null, issueType: 'Reselling', prjStatus: null, summary: null, earned: 0, finIssues: validFin, rowInvoiced, rowInvoicedUSD, projectStart: null, projectEnd: null, totalPlannedHrs: 0, billingRate: 0, nonAccruedHrs: 0, totalNonAccruable: 0, salesInvoiceRate: 0, earnedMonthly: 0 }];
  } else {
    // Standard: FIN tickets are inside PRJ/AUS linked issues
    const linkedKeys = (ismIssue.fields.issuelinks || [])
      .map((l) => l.outwardIssue?.key || l.inwardIssue?.key)
      .filter((k) => k && (k.startsWith('PRJ-') || k.startsWith('AUS-')));

    const rowResults = await Promise.all(linkedKeys.map(async (key) => {
      try {
        const linked          = await jiraGet(`/rest/api/3/issue/${key}`);
        const earned          = num(linked.fields[EARNED_AMOUNT_FIELD]);
        const issueType       = linked.fields.issuetype?.name || '';
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
          .filter((k) => k && k.startsWith('FIN-'));

        const finIssues = await Promise.all(
          finKeys.map(async (finKey) => {
            const fin = await jiraGet(`/rest/api/3/issue/${finKey}`);
            return { key: finKey, summary: fin.fields.summary, invoiced: num(fin.fields[INVOICE_AMOUNT_FIELD]), invoicedUSD: num(fin.fields[INVOICE_USD_FIELD]) };
          })
        );

        const rowInvoiced    = finIssues.reduce((s, f) => s + f.invoiced, 0);
        const rowInvoicedUSD = finIssues.reduce((s, f) => s + f.invoicedUSD, 0);
        return { key, issueType, prjStatus: linked.fields.status?.name || null, summary: linked.fields.summary, earned, finIssues, rowInvoiced, rowInvoicedUSD, projectStart, projectEnd, totalPlannedHrs, billingRate, nonAccruedHrs, totalNonAccruable, salesInvoiceRate, earnedMonthly };
      } catch (err) {
        console.error('Skipping linked issue', key, err.message);
        return null;
      }
    }));

    rows = rowResults.filter(Boolean);
    for (const row of rows) {
      if (row.key.startsWith('PRJ-')) prjEarned += row.earned;
      else if (row.key.startsWith('AUS-')) ausEarned += row.earned;
      totalInvoiced    += row.rowInvoiced;
      totalInvoicedUSD += row.rowInvoicedUSD;
    }
  }

  return {
    ism: {
      key: ismIssue.key,
      summary: ismIssue.fields.summary,
      serviceAmount,
      revenueStream: revenueStream || null,
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

// ── Access control ─────────────────────────────────────────────────────────
const ACCESS_GROUP = 'Finance Reports';

resolver.define('checkUserAccess', async ({ context }) => {
  const accountId = context?.accountId;
  if (!accountId) return { hasAccess: false };
  try {
    const groups = await jiraGet(`/rest/api/3/user/groups?accountId=${encodeURIComponent(accountId)}`);
    const list = Array.isArray(groups) ? groups : (groups.items || []);
    const hasAccess = list.some(g => g.name === ACCESS_GROUP);
    return { hasAccess };
  } catch {
    return { hasAccess: false };
  }
});

// Single ISM issue
resolver.define('getIsmReport', async ({ payload }) => {
  const { issueKey } = payload;
  const ism = await jiraGet(`/rest/api/3/issue/${issueKey}`);
  return processIsmIssue(ism);
});

// All ISM issues
resolver.define('getIsmProjects', async ({ payload }) => {
  const fromDate    = payload?.fromDate    || '2025-04-01';
  const toDate      = payload?.toDate      || null;
  const reportType  = payload?.reportType  || 'Service';
  const todayClause = toDate ? ` AND resolved <= "${toDate}"` : '';
  const typeClause  = reportType === 'Reselling'
    ? `AND cf[11077] = "R - Reselling"`
    : `AND cf[11077] != "R - Reselling"`;
  const fields = ['summary', SERVICE_AMOUNT_FIELD, 'issuelinks', 'issuetype', 'created', 'status', 'resolution', 'resolutiondate', CLAIM_DATE_FIELD, PO_AMOUNT_FIELD, SALES_INVOICE_DATE_FIELD, REVENUE_STREAM_FIELD];
  const jql    = `project = ISM AND resolved >= "${fromDate}"${todayClause} AND status = "Closed" AND resolution IN ("Fixed", "done") ${typeClause} ORDER BY cf[16311] DESC`;


  const allIssues = [];
  let nextPageToken = undefined;

  while (true) {
    const body = { jql, maxResults: 50, fields };
    if (nextPageToken) body.nextPageToken = nextPageToken;

    const page   = await jiraPost('/rest/api/3/search/jql', body);
    const issues = page.issues || [];
    allIssues.push(...issues);
    if (!page.nextPageToken || issues.length === 0) break;
    nextPageToken = page.nextPageToken;
  }

  // Process all ISM issues in parallel
  const reports = await Promise.all(allIssues.map(ism => processIsmIssue(ism)));
  return { reports, total: reports.length };
});

// ── Reselling report ───────────────────────────────────────────────────────
resolver.define('getResellingProjects', async ({ payload }) => {
  const fromDate = payload?.fromDate || '2025-04-01';
  const toDate   = payload?.toDate   || null;
  const todayClause = toDate ? ` AND resolved <= "${toDate}"` : '';
  const jql = `project = ISM AND resolved >= "${fromDate}"${todayClause} AND status = "Closed" AND resolution IN ("Fixed", "done") AND cf[11077] = "R - Reselling" ORDER BY cf[16311] DESC`;

  // 1. Fetch all ISM issues — nextPageToken cursor pagination, 100 per page
  const allIsmIssues = [];
  let nextPageToken;
  do {
    const body = { jql, maxResults: 100, fields: ['summary', 'issuelinks'] };
    if (nextPageToken) body.nextPageToken = nextPageToken;
    const page = await jiraPost('/rest/api/3/search/jql', body);
    allIsmIssues.push(...(page.issues || []));
    nextPageToken = page.nextPageToken;
  } while (nextPageToken);

  // 2. Collect all unique AR FIN keys from ISM issuelinks
  const arFinKeySet = new Set();
  for (const ism of allIsmIssues) {
    for (const link of (ism.fields.issuelinks || [])) {
      const k = link.outwardIssue?.key || link.inwardIssue?.key;
      if (k && k.startsWith('FIN-')) arFinKeySet.add(k);
    }
  }
  const allArFinKeys = [...arFinKeySet];

  // 3. Batch-fetch ALL AR FIN issues in parallel chunks of 50 (replaces N individual GETs)
  const arFinByKey = {};
  const arFinBatches = [];
  for (let i = 0; i < allArFinKeys.length; i += 50) arFinBatches.push(allArFinKeys.slice(i, i + 50));
  const arFinPages = await Promise.all(arFinBatches.map(chunk =>
    jiraPost('/rest/api/3/search/jql', {
      jql: `key IN (${chunk.map(k => `"${k}"`).join(',')})`,
      maxResults: 50,
      fields: [F_COMPANY_NAME, F_LICENSE_AMOUNT, INVOICE_USD_FIELD, F_OEM_QUOTE, F_ACTUAL_COLLECTION, F_OEM_PURCHASE, 'issuelinks'],
    }).catch(() => ({ issues: [] }))
  ));
  for (const page of arFinPages) {
    for (const issue of (page.issues || [])) arFinByKey[issue.key] = issue.fields;
  }

  // 4. Collect all unique AP FIN keys from AR FIN issuelinks
  const apFinKeySet = new Set();
  for (const [arFinKey, fields] of Object.entries(arFinByKey)) {
    for (const link of (fields.issuelinks || [])) {
      const k = link.outwardIssue?.key || link.inwardIssue?.key;
      if (k && k.startsWith('FIN-') && k !== arFinKey) apFinKeySet.add(k);
    }
  }
  const allApFinKeys = [...apFinKeySet];

  // 5. Batch-fetch ALL AP FIN issues in parallel chunks of 50
  const apFinByKey = {};
  if (allApFinKeys.length > 0) {
    const apFinBatches = [];
    for (let i = 0; i < allApFinKeys.length; i += 50) apFinBatches.push(allApFinKeys.slice(i, i + 50));
    const apFinPages = await Promise.all(apFinBatches.map(chunk =>
      jiraPost('/rest/api/3/search/jql', {
        jql: `key IN (${chunk.map(k => `"${k}"`).join(',')})`,
        maxResults: 50,
        fields: [F_OEM_PURCHASE],
      }).catch(() => ({ issues: [] }))
    ));
    for (const page of apFinPages) {
      for (const issue of (page.issues || [])) apFinByKey[issue.key] = issue.fields;
    }
  }

  // 6. Build rows in memory — zero additional API calls
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
      const arF = arFinByKey[arFinKey];
      if (!arF) {
        allRows.push({ ismKey, companyName: '', licenseAmount: '', oemPurchase: '', arFinKey, invoiceUSD: '', oemQuote: '', actualCollection: '', apFinKey: '' });
        continue;
      }

      const raw = arF[F_COMPANY_NAME];
      const companyName      = (typeof raw === 'string' ? raw : raw?.value || raw?.name || '') || '';
      const licenseAmount    = num(arF[F_LICENSE_AMOUNT]);
      const invoiceUSD       = num(arF[INVOICE_USD_FIELD]);
      const oemQuote         = num(arF[F_OEM_QUOTE]);
      const actualCollection = num(arF[F_ACTUAL_COLLECTION]);
      const oemPurchaseOnAR  = num(arF[F_OEM_PURCHASE]);

      const apFinLinks = (arF.issuelinks || [])
        .filter(l => { const k = l.outwardIssue?.key || l.inwardIssue?.key || ''; return k.startsWith('FIN-') && k !== arFinKey; });

      if (apFinLinks.length === 0) {
        allRows.push({ ismKey, companyName, licenseAmount, arFinKey, invoiceUSD, oemQuote, actualCollection, apFinKey: '', oemPurchase: oemPurchaseOnAR });
      } else {
        for (const apLink of apFinLinks) {
          const apFinKey    = apLink.outwardIssue?.key || apLink.inwardIssue?.key;
          const apF         = apFinByKey[apFinKey];
          const oemPurchase = apF ? num(apF[F_OEM_PURCHASE]) : oemPurchaseOnAR;
          allRows.push({ ismKey, companyName, licenseAmount, arFinKey, invoiceUSD, oemQuote, actualCollection, apFinKey, oemPurchase });
        }
      }
    }
  }

  return { rows: allRows, total: allRows.length };
});

// ── Extra Info report — ISM (R-Reselling) → linked FIN ticket ─────────────
// Filtered by Claim Date (not Resolved date) and restricted to R-Reselling,
// per confirmed business decision. See extra-info-field-map.txt for the audit
// of which fields are ISM-sourced vs FIN-sourced.
// Batch-fetch a set of issue keys via `key IN (...)` JQL, in parallel chunks of
// 50 — replaces N individual GETs with ceil(N/50) requests run concurrently.
// Forge functions have a hard 25s execution limit, so sequential per-issue
// GETs (fine for the local Express server) are not viable here.
async function batchFetchIssues(keys, fields) {
  const byKey = {};
  if (keys.length === 0) return byKey;
  const batches = [];
  for (let i = 0; i < keys.length; i += 50) batches.push(keys.slice(i, i + 50));
  const pages = await Promise.all(batches.map(chunk =>
    jiraPost('/rest/api/3/search/jql', {
      jql: `key IN (${chunk.map(k => `"${k}"`).join(',')})`,
      maxResults: 50,
      fields,
    }).catch(() => ({ issues: [] }))
  ));
  for (const page of pages) {
    for (const issue of (page.issues || [])) byKey[issue.key] = issue.fields;
  }
  return byKey;
}

async function fetchExtraInfoRows(fromDate, toDate) {
  const todayClause = toDate ? ` AND cf[16311] <= "${toDate}"` : '';
  const jql = `project = ISM AND cf[16311] >= "${fromDate}"${todayClause} AND cf[11077] = "R - Reselling" ORDER BY cf[16311] DESC`;

  // 1. Fetch all matching ISM issues (paginated).
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

  // 2. Split ISM issues into "direct FIN link" (Reselling pattern) vs
  //    "needs PRJ/AUS lookup" (Service pattern) — no per-issue awaits yet.
  const directFinByIsm = new Map();
  const prjAusKeysByIsm = new Map();
  for (const ism of allIsmIssues) {
    const directFinKeys = (ism.fields.issuelinks || [])
      .map(l => l.outwardIssue?.key || l.inwardIssue?.key)
      .filter(k => k && k.startsWith('FIN-'));
    if (directFinKeys.length > 0) {
      directFinByIsm.set(ism.key, directFinKeys);
    } else {
      const prjAusKeys = (ism.fields.issuelinks || [])
        .map(l => l.outwardIssue?.key || l.inwardIssue?.key)
        .filter(k => k && (k.startsWith('PRJ-') || k.startsWith('AUS-')));
      if (prjAusKeys.length > 0) prjAusKeysByIsm.set(ism.key, prjAusKeys);
    }
  }

  // 3. Batch-fetch all PRJ/AUS issuelinks needed for the Service-pattern
  //    fallback, then resolve their FIN keys.
  const allPrjAusKeys = [...new Set([...prjAusKeysByIsm.values()].flat())];
  const prjAusByKey = await batchFetchIssues(allPrjAusKeys, ['issuelinks']);
  const finByIsmViaPrjAus = new Map();
  for (const [ismKey, prjAusKeys] of prjAusKeysByIsm.entries()) {
    const finKeySet = new Set();
    for (const prjAusKey of prjAusKeys) {
      const linkedFields = prjAusByKey[prjAusKey];
      if (!linkedFields) continue;
      for (const l of (linkedFields.issuelinks || [])) {
        const k = l.outwardIssue?.key || l.inwardIssue?.key;
        if (k && k.startsWith('FIN-')) finKeySet.add(k);
      }
    }
    if (finKeySet.size > 0) finByIsmViaPrjAus.set(ismKey, [...finKeySet]);
  }

  // 4. Batch-fetch every unique FIN ticket needed, across both patterns.
  const allFinKeys = [...new Set([
    ...[...directFinByIsm.values()].flat(),
    ...[...finByIsmViaPrjAus.values()].flat(),
  ])];
  const finByKey = await batchFetchIssues(allFinKeys, EXTRA_INFO_FIN_FIELDS);

  // 5. Build rows in memory — zero additional API calls.
  const rows = [];
  for (const ism of allIsmIssues) {
    const ismF = ism.fields;
    const base = {
      issueKey: ism.key,
      issueId: ism.id,
      summary: ismF.summary || '',
      status: ismF.status?.name || '',
      claimDate: ismF[CLAIM_DATE_FIELD] || null,
      billingType:        pickValue(ismF[F_BILLING_TYPE]),
      revenueStream:      pickValue(ismF[REVENUE_STREAM_FIELD]),
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

    const finKeys = directFinByIsm.get(ism.key) || finByIsmViaPrjAus.get(ism.key) || [];

    if (finKeys.length === 0) {
      rows.push({ ...base, finKey: '' });
      continue;
    }

    for (const finKey of finKeys) {
      const f = finByKey[finKey];
      if (!f) { rows.push({ ...base, finKey }); continue; }
      rows.push({
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
    }
  }

  return rows;
}

resolver.define('getExtraInfoProjects', async ({ payload }) => {
  const fromDate = payload?.fromDate || '2025-04-01';
  const toDate   = payload?.toDate   || null;
  const rows = await fetchExtraInfoRows(fromDate, toDate);
  return { rows, total: rows.length };
});

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

resolver.define('generateExtraInfoExcel', async ({ payload }) => {
  const { rows } = payload;
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

  const buffer = await wb.xlsx.writeBuffer();
  const base64 = Buffer.from(buffer).toString('base64');
  const date   = new Date().toISOString().slice(0, 10);
  return { base64, filename: `extra-info-report-${date}.xlsx` };
});

resolver.define('generateResellingExcel', async ({ payload }) => {
  const { rows } = payload;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Finance Report';
  const ws = wb.addWorksheet('Reselling', { views: [{ state: 'frozen', ySplit: 1 }] });

  ws.columns = [
    { header: 'ISM Tkt ID',                           key: 'ismKey',           width: 16 },
    { header: 'Company Name',                          key: 'companyName',      width: 40 },
    { header: 'License Amount ($)',                    key: 'licenseAmount',    width: 20 },
    { header: 'OEM Purchase Amount ($)',               key: 'oemPurchase',      width: 24 },
    { header: 'AR FIN Tkt ID',                         key: 'arFinKey',         width: 16 },
    { header: 'Invoice Amount in USD (Excluding Tax)', key: 'invoiceUSD',       width: 32 },
    { header: 'OEM Quote Amount ($)',                  key: 'oemQuote',         width: 22 },
    { header: 'Actual Collection Amount in USD',       key: 'actualCollection', width: 30 },
    { header: 'AP FIN Tkt ID',                         key: 'apFinKey',         width: 16 },
  ];

  const hdrRow = ws.getRow(1);
  hdrRow.height = 32;
  hdrRow.eachCell(cell => {
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF172B4D' } };
    cell.font      = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border    = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
  });

  const numCols = new Set([3, 4, 6, 7, 8]);
  for (const row of rows) {
    const r = ws.addRow([row.ismKey, row.companyName, row.licenseAmount, row.oemPurchase, row.arFinKey, row.invoiceUSD, row.oemQuote, row.actualCollection, row.apFinKey]);
    r.height = 18;
    r.eachCell({ includeEmpty: true }, (cell, colNum) => {
      cell.border    = { top: { style: 'thin', color: { argb: 'FFDFE1E6' } }, bottom: { style: 'thin', color: { argb: 'FFDFE1E6' } }, left: { style: 'thin', color: { argb: 'FFDFE1E6' } }, right: { style: 'thin', color: { argb: 'FFDFE1E6' } } };
      cell.alignment = { vertical: 'middle', horizontal: numCols.has(colNum) ? 'right' : 'left' };
    });
    r.commit();
  }

  const buffer = await wb.xlsx.writeBuffer();
  const base64 = Buffer.from(buffer).toString('base64');
  const date   = new Date().toISOString().slice(0, 10);
  return { base64, filename: `reselling-report-${date}.xlsx` };
});

// ── Excel export ───────────────────────────────────────────────────────────
const fmtDate = (iso) => {
  if (!iso) return '—';
  try { return new Date(iso).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }); }
  catch { return iso; }
};
const fmt = (n) => num(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

async function buildExcel(reports) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Finance Report';
  wb.created = new Date();
  const ws = wb.addWorksheet('ISM Report', { views: [{ state: 'frozen', ySplit: 1 }] });

  ws.columns = [
    { width: 16 }, { width: 42 }, { width: 16 }, { width: 18 },
    { width: 14 }, { width: 16 }, { width: 18 }, { width: 26 }, { width: 20 },
    { width: 26 }, { width: 22 }, { width: 18 }, { width: 16 },
    { width: 26 }, { width: 26 },
    { width: 16 }, { width: 16 }, { width: 16 }, { width: 22 }, { width: 22 },
    { width: 20 }, { width: 14 }, { width: 20 }, { width: 22 },
    { width: 22 }, { width: 18 }, { width: 24 }, { width: 22 },
    { width: 26 }, { width: 22 }, { width: 18 }, { width: 22 },
    { width: 24 }, { width: 26 }, { width: 22 },
    { width: 18 }, { width: 28 }, { width: 20 }, { width: 20 },
  ];

  const ISM_COLS = 15;
  const NUM_COLS = new Set([4,7,9,10,11,12,13,14,15, 21,22,23,24,25,26,27,28,29,30,31,32,33,34,35, 37,38,39]);
  const thinGrey = { style: 'thin',   color: { argb: 'FFDFE1E6' } };
  const medBlue  = { style: 'medium', color: { argb: 'FF0052CC' } };
  const border   = (last) => ({ top: thinGrey, left: thinGrey, right: thinGrey, bottom: last ? medBlue : thinGrey });
  const ismFill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF4F5F7' } };

  const hdr = ws.addRow([
    'ISM Issue', 'Issue Name', 'PO Date', 'PO Amount', ' ISM Status', 'Resolved',
    'Service Amount', 'Sales Invoice Creation Date',
    '% Billed Till Date', 'Cummulative Billing on Project', 'Cummulative Billing INR',
    'Billed Till Date', '% Billing Done', 'Check for Revenue against PO', 'Check for billing against PO',
    'PRJ Status', 'Linked Issue', 'Issue Type', 'Planned Project Start Date', 'Planned Project End Date',
    'Efforts Planned Hours', 'Billing Rate', 'Total Earned Hours', 'Hours Consumed Till Date',
    '% Completion Till Date', 'Unbilled Till Date', 'Cummulative Consumed Hours',
    'Updated Remaining Hours', 'Updated Total Estimated Hours', '% Completetion Till Date',
    'Unbilled in INR', 'Revenue this month FX', 'Recovery Rate this month',
    'Revised Recovery Rate till Date', 'Total Earned Amount ($)',
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
    const startRow    = rowNum;
    const rowCount    = Math.max(rows.length, 1);
    const svc         = ism.serviceAmount;
    const totalEarned = report.prjEarned + report.ausEarned;
    const dataRows    = rows.length === 0 ? [null] : rows;

    for (let i = 0; i < dataRows.length; i++) {
      const row    = dataRows[i];
      const isLast = i === dataRows.length - 1;
      const empty  = row === null;
      const finKey = !empty && row.finIssues?.length ? row.finIssues.map(f => f.key).join('\n')              : '—';
      const finINR = !empty && row.finIssues?.length ? row.finIssues.map(f => fmt(f.invoiced)).join('\n')    : '—';
      const finUSD = !empty && row.finIssues?.length ? row.finIssues.map(f => fmt(f.invoicedUSD)).join('\n') : '—';
      const d  = (v)       => empty ? '—' : (v  ? fmt(v)  : '—');
      const dc = (cond, v) => empty ? '—' : (cond ? fmt(v) : '—');

      const r = ws.addRow([
        i === 0 ? ism.key                                                         : '',
        i === 0 ? ism.summary                                                     : '',
        i === 0 ? fmtDate(ism.claimDate)                                          : '',
        i === 0 ? (ism.poAmount ? fmt(ism.poAmount) : '—')                       : '',
        i === 0 ? (ism.status   || '—')                                           : '',
        i === 0 ? fmtDate(ism.resolutionDate)                                     : '',
        i === 0 ? fmt(svc)                                                        : '',
        i === 0 ? fmtDate(ism.salesInvoiceDate)                                   : '',
        i === 0 ? (svc ? fmt(report.totalInvoiced    / svc) : '—')               : '',
        i === 0 ? fmt(report.totalInvoicedUSD)                                    : '',
        i === 0 ? fmt(report.totalInvoiced)                                       : '',
        i === 0 ? fmt(report.totalInvoicedUSD)                                    : '',
        i === 0 ? (svc ? fmt(report.totalInvoicedUSD / svc) : '—')               : '',
        i === 0 ? (svc ? fmt(totalEarned             / svc) : '—')               : '',
        i === 0 ? (svc ? fmt(report.totalInvoicedUSD / svc) : '—')               : '',
        empty ? '—' : (row.prjStatus  || '—'),
        empty ? '—' : row.key,
        empty ? '—' : (row.issueType  || '—'),
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
        (empty || !(row.nonAccruedHrs && row.billingRate))                 ? '—' : fmt(row.nonAccruedHrs * row.billingRate),
        d(row?.earnedMonthly),
        (empty || !(row.rowInvoicedUSD + row.totalNonAccruable))           ? '—' : fmt(row.rowInvoicedUSD / (row.rowInvoicedUSD + row.totalNonAccruable)),
        (empty || !(row.rowInvoicedUSD + row.totalNonAccruable))           ? '—' : fmt(row.rowInvoicedUSD / (row.rowInvoicedUSD + row.totalNonAccruable)),
        empty ? '—' : fmt(row.earned),
        finKey, finINR, finUSD, finINR,
      ]);

      r.height = 20;
      r.eachCell({ includeEmpty: true }, (cell, col) => {
        cell.border    = border(isLast);
        cell.alignment = { vertical: 'middle', wrapText: true };
        if (col <= ISM_COLS)   cell.fill      = ismFill;
        if (NUM_COLS.has(col)) cell.alignment = { ...cell.alignment, horizontal: 'right' };
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

resolver.define('generateExcelReport', async ({ payload }) => {
  const { reports } = payload;
  const wb     = await buildExcel(reports);
  const buffer = await wb.xlsx.writeBuffer();
  const base64 = Buffer.from(buffer).toString('base64');
  const date   = new Date().toISOString().slice(0, 10);
  return { base64, filename: `ism-report-${date}.xlsx` };
});

module.exports.handler = resolver.getDefinitions();

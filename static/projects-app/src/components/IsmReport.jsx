import { useState, useEffect, useRef } from 'react';
import './IsmReport.css';

const num = (n) => {
  const v = Number(n);
  return isNaN(v) ? '0.00' : v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

const fmtDate = (iso) => {
  if (!iso) return '—';
  return new Date(iso).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
};

const numOrDash = (v) => (v !== '' && v !== null && v !== undefined ? num(v) : '—');

export default function IsmReport() {
  const [reports,       setReports]       = useState([]);
  const [accessState,   setAccessState]   = useState('checking'); // 'checking' | 'granted' | 'denied'
  const [resellingRows, setResellingRows] = useState([]);
  const [extraInfoRows, setExtraInfoRows] = useState([]);
  const [expected,      setExpected]      = useState(0);
  const [done,          setDone]          = useState(false);
  const [loading,       setLoading]       = useState(true);
  const [error,         setError]         = useState(null);
  const [search,        setSearch]        = useState('');
  const [exporting,     setExporting]     = useState(false);
  const [reportType,    setReportType]    = useState('Service');
  const [fromDate, setFromDate] = useState(() => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-01`; });
  const [toDate,   setToDate]   = useState(new Date().toISOString().slice(0, 10));
  const sourceRef = useRef(null);

  const closeStream = () => {
    if (sourceRef.current) { sourceRef.current.close(); sourceRef.current = null; }
  };

  const resetState = () => {
    setReports([]);
    setResellingRows([]);
    setExtraInfoRows([]);
    setExpected(0);
    setDone(false);
    setLoading(true);
    setError(null);
  };

  // ── LOCAL DEV ──────────────────────────────────────────────────────────
  const loadStream = (type) => {
    closeStream();
    resetState();

    if (type === 'Extra Info') {
      const source = new EventSource(`/api/extra-info-projects/stream?from=${fromDate}&to=${toDate}`);
      sourceRef.current = source;
      source.addEventListener('total', (e) => setExpected(JSON.parse(e.data).total));
      source.addEventListener('row',  (e) => { setLoading(false); setExtraInfoRows((p) => [...p, JSON.parse(e.data)]); });
      source.addEventListener('done', () => { setDone(true); setLoading(false); closeStream(); });
      source.addEventListener('error', (e) => {
        try { setError(JSON.parse(e.data).error); } catch { setError('Stream error'); }
        setLoading(false); closeStream();
      });
      return;
    }

    if (type === 'Reselling') {
      fetch(`/api/reselling-projects?from=${fromDate}&to=${toDate}`)
        .then(r => r.json())
        .then(json => { setResellingRows(json.rows); setExpected(json.total); setDone(true); setLoading(false); })
        .catch(err => { setError(String(err)); setLoading(false); });
      return;
    }

    const source = new EventSource(`/api/ism-projects/stream?from=${fromDate}&to=${toDate}&type=${type}`);
    sourceRef.current = source;
    source.addEventListener('total', (e) => setExpected(JSON.parse(e.data).total));
    source.addEventListener('report', (e) => { setLoading(false); setReports((p) => [...p, JSON.parse(e.data)]); });
    source.addEventListener('done',  () => { setDone(true); setLoading(false); closeStream(); });
    source.addEventListener('error', (e) => {
      try { setError(JSON.parse(e.data).error); } catch { setError('Stream error'); }
      setLoading(false); closeStream();
    });
  };

  // ── FORGE PRODUCTION ───────────────────────────────────────────────────
  const loadForge = async (type) => {
    closeStream();
    resetState();
    try {
      const { invoke } = await import('@forge/bridge');
      if (type === 'Reselling') {
        const json = await invoke('getResellingProjects', { fromDate, toDate });
        setResellingRows(json.rows);
        setExpected(json.total);
      } else if (type === 'Extra Info') {
        const json = await invoke('getExtraInfoProjects', { fromDate, toDate });
        setExtraInfoRows(json.rows);
        setExpected(json.total);
      } else {
        const json = await invoke('getIsmProjects', { fromDate, toDate, reportType: type });
        setReports(json.reports);
        setExpected(json.total);
      }
      setDone(true);
    } catch (err) {
      setError(String(err));
    } finally {
      setLoading(false);
    }
  };

  const loadAll = (type) => import.meta.env.DEV ? loadStream(type) : loadForge(type);

  // ── Single ISM issue ───────────────────────────────────────────────────
  const loadSingle = async (key) => {
    closeStream();
    resetState();
    setExpected(1);
    try {
      let json;
      if (import.meta.env.DEV) {
        const res = await fetch(`/api/ism-report/${key}`);
        json = await res.json();
        if (!res.ok) throw new Error(json.error || `HTTP ${res.status}`);
      } else {
        const { invoke } = await import('@forge/bridge');
        json = await invoke('getIsmReport', { issueKey: key });
      }
      setReports([json]);
      setDone(true);
    } catch (err) {
      setError(String(err));
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (import.meta.env.DEV) {
      setAccessState('granted');
      loadAll(reportType);
      return closeStream;
    }
    import('@forge/bridge').then(({ invoke }) =>
      invoke('checkUserAccess')
        .then(({ hasAccess }) => {
          if (hasAccess) { setAccessState('granted'); loadAll(reportType); }
          else setAccessState('denied');
        })
        .catch(() => setAccessState('denied'))
    );
    return closeStream;
  }, []);

  const handleSearch = () => {
    let key = search.trim().toUpperCase();
    if (!key) return;
    if (/^\d+$/.test(key)) key = `ISM-${key}`;
    loadSingle(key);
  };

  const rowCount = reportType === 'Reselling' ? resellingRows.length
    : reportType === 'Extra Info' ? extraInfoRows.length
    : reports.length;

  if (accessState === 'checking') {
    return <div className="ism-access-checking">Checking access…</div>;
  }

  if (accessState === 'denied') {
    return (
      <div className="ism-access-denied">
        <div className="ism-access-denied__icon">🔒</div>
        <div className="ism-access-denied__title">Access Denied</div>
        <div className="ism-access-denied__msg">
          You don't have permission to view this report.<br />
          Ask your Jira admin to add you to the <strong>Finance Reports</strong> group.
        </div>
      </div>
    );
  }

  return (
    <div>
      <div className="ism-search-row">
        <label className="ism-date-label">
          From
          <input className="ism-input ism-input--date" type="date" value={fromDate} max={toDate}
            onChange={(e) => setFromDate(e.target.value)} disabled={loading} />
        </label>
        <label className="ism-date-label">
          To
          <input className="ism-input ism-input--date" type="date" value={toDate} min={fromDate}
            onChange={(e) => setToDate(e.target.value)} disabled={loading} />
        </label>
        <button className="ism-btn" style={{ alignSelf: 'flex-end' }} onClick={() => loadAll(reportType)} disabled={loading}>
          Apply
        </button>
        <select
          className="ism-select"
          style={{ marginLeft: 'auto', alignSelf: 'flex-end' }}
          value={reportType}
          disabled={loading}
          onChange={(e) => { const t = e.target.value; setReportType(t); loadAll(t); }}
        >
          <option value="Service">Service</option>
          <option value="Reselling">Reselling</option>
          <option value="Extra Info">Financial report for Reselling</option>
        </select>
        <button
          className="ism-btn ism-btn--download ism-btn--sm"
          style={{ alignSelf: 'flex-end' }}
          disabled={exporting}
          onClick={async () => {
            setExporting(true);
            try {
              let blob;
              let filename = `ism-report-${new Date().toISOString().slice(0, 10)}.xlsx`;
              if (import.meta.env.DEV) {
                // Note: Reselling has no dedicated xlsx export endpoint in local dev yet —
                // falls back to the Service export, same as before this fix.
                const exportPath = reportType === 'Extra Info' ? '/api/extra-info-projects/export' : '/api/ism-projects/export';
                const res = await fetch(`${exportPath}?from=${fromDate}&to=${toDate}`);
                blob = await res.blob();
                filename = res.headers.get('content-disposition')?.match(/filename="?([^"]+)"?/)?.[1] || filename;
              } else {
                const { invoke } = await import('@forge/bridge');
                const resolverFn = reportType === 'Reselling' ? 'generateResellingExcel'
                  : reportType === 'Extra Info' ? 'generateExtraInfoExcel'
                  : 'generateExcelReport';
                const resolverPayload = reportType === 'Reselling' ? { rows: resellingRows }
                  : reportType === 'Extra Info' ? { rows: extraInfoRows }
                  : { reports };
                const { base64, filename: f } = await invoke(resolverFn, resolverPayload);
                filename = f;
                const binary = atob(base64);
                const bytes  = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
                blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
              }
              const url = URL.createObjectURL(blob);
              const a = document.createElement('a');
              a.href = url; a.download = filename; a.click();
              URL.revokeObjectURL(url);
            } finally {
              setExporting(false);
            }
          }}
        >
          {exporting ? 'Exporting…' : '↓ Download Excel'}
        </button>
      </div>

      {loading && rowCount === 0 && <div className="ism-empty">Loading…</div>}
      {!loading && !done && rowCount > 0 && (
        <div className="ism-progress">
          Loaded {rowCount}{expected > 0 ? ` / ${expected}` : ''} — fetching more…
        </div>
      )}
      {done && (
        <div className="ism-progress ism-progress--done">
          ✓ {reportType === 'Reselling'
            ? `${resellingRows.length} row${resellingRows.length !== 1 ? 's' : ''}`
            : reportType === 'Extra Info'
            ? `${extraInfoRows.length} row${extraInfoRows.length !== 1 ? 's' : ''}`
            : `${reports.length} ISM issue${reports.length !== 1 ? 's' : ''}`} loaded
        </div>
      )}
      {done && reportType === 'Extra Info' && extraInfoRows.length === 0 && (
        <div className="ism-empty">No Extra Info rows for this date range.</div>
      )}
      {error && <div className="ism-error"><strong>Error:</strong> {error}</div>}

      {/* ── Reselling flat table ── */}
      {reportType === 'Reselling' && resellingRows.length > 0 && (
        <div className="ism-table-wrap">
          <table className="ism-report-table">
            <thead>
              <tr>
                <th>ISM Tkt ID</th>
                <th>Company Name</th>
                <th className="ism-th--right">License Amount ($)</th>
                <th className="ism-th--right">OEM Purchase Amount ($)</th>
                <th>AR FIN Tkt ID</th>
                <th className="ism-th--right">Invoice Amount in USD (Excluding Tax)</th>
                <th className="ism-th--right">OEM Quote Amount ($)</th>
                <th className="ism-th--right">Actual Collection Amount in USD</th>
                <th>AP FIN Tkt ID</th>
              </tr>
            </thead>
            <tbody>
              {resellingRows.map((row, i) => (
                <tr key={i}>
                  <td className="ism-td"><span className="ism-key-badge">{row.ismKey}</span></td>
                  <td className="ism-td">{row.companyName || '—'}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.licenseAmount)}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.oemPurchase)}</td>
                  <td className="ism-td">{row.arFinKey ? <span className="ism-fin-badge">{row.arFinKey}</span> : '—'}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.invoiceUSD)}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.oemQuote)}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.actualCollection)}</td>
                  <td className="ism-td">{row.apFinKey ? <span className="ism-fin-badge">{row.apFinKey}</span> : '—'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* ── Extra Info flat table ── */}
      {reportType === 'Extra Info' && extraInfoRows.length > 0 && (
        <div className="ism-table-wrap">
          <table className="ism-report-table">
            <thead>
              <tr>
                <th>Issue key</th>
                <th>Issue id</th>
                <th>FIN</th>
                <th>Claim Date</th>
                <th>Company Name</th>
                <th>Summary</th>
                <th>Billing Entity</th>
                <th>Billing Frequency</th>
                <th>Billing Type</th>
                <th>Revenue Stream</th>
                <th>Segment BU</th>
                <th>Market type</th>
                <th>Market Team</th>
                <th>Geo</th>
                <th>Region and Sub-Region</th>
                <th>Exchange Rate Type</th>
                <th>Forex Exchange Rate</th>
                <th className="ism-th--right">Forex Mark Up (INR)</th>
                <th>OEM</th>
                <th className="ism-th--right">License Amount ($)</th>
                <th className="ism-th--right">OEM Purchase Amount ($)</th>
                <th className="ism-th--right">Profitability</th>
                <th className="ism-th--right">DF Amount ($)</th>
                <th className="ism-th--right">DR Amount ($)</th>
                <th className="ism-th--right">DR/Rebate/DF Amount ($)</th>
                <th className="ism-th--right">PVR Amount ($)</th>
                <th className="ism-th--right">Services Amount ($)</th>
                <th className="ism-th--right">Invoice Amount in local currency (Without Tax)</th>
                <th>Sales Invoice Number</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {extraInfoRows.map((row, i) => (
                <tr key={i}>
                  <td className="ism-td"><span className="ism-key-badge">{row.issueKey}</span></td>
                  <td className="ism-td">{row.issueId ?? '—'}</td>
                  <td className="ism-td">{row.finKey ? <span className="ism-fin-badge">{row.finKey}</span> : '—'}</td>
                  <td className="ism-td">{fmtDate(row.claimDate)}</td>
                  <td className="ism-td">{row.companyName || '—'}</td>
                  <td className="ism-td">{row.summary || '—'}</td>
                  <td className="ism-td">{row.billingEntity || '—'}</td>
                  <td className="ism-td">{row.billingFrequency || '—'}</td>
                  <td className="ism-td">{row.billingType || '—'}</td>
                  <td className="ism-td">{row.revenueStream || '—'}</td>
                  <td className="ism-td">{row.segmentBu || '—'}</td>
                  <td className="ism-td">{row.marketType || '—'}</td>
                  <td className="ism-td">{row.marketTeam || '—'}</td>
                  <td className="ism-td">{row.geo || '—'}</td>
                  <td className="ism-td">{row.regionSubRegion || '—'}</td>
                  <td className="ism-td">{row.exchangeRateType || '—'}</td>
                  <td className="ism-td">{row.forexExchangeRate || '—'}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.forexMarkUpInr)}</td>
                  <td className="ism-td">{row.oem || '—'}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.licenseAmount)}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.oemPurchaseAmount)}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.profitability)}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.dfAmount)}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.drAmount)}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.drRebateDfAmount)}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.pvrAmount)}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.servicesAmount)}</td>
                  <td className="ism-td ism-td--num">{numOrDash(row.invoiceAmountLocal)}</td>
                  <td className="ism-td">{row.salesInvoiceNumber || '—'}</td>
                  <td className="ism-td">{row.status || '—'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* ── Service grouped table ── */}
      {reportType === 'Service' && reports.length > 0 && (
        <div className="ism-table-wrap">
          <table className="ism-report-table">
            <thead>
              <tr>
                <th>ISM Issue</th>
                <th>Issue Name</th>
                <th>PO Date</th>
                <th className="ism-th--right">PO Amount</th>
                <th>ISM Status</th>
                <th>Resolved</th>
                <th>PRJ Status</th>
                <th className="ism-th--right">Service Amount</th>
                <th>Sales Invoice Creation Date</th>
                <th className="ism-th--right">% Billed Till Date</th>
                <th className="ism-th--right">Cummulative Billing on Project</th>
                <th className="ism-th--right">Cummulative Billing INR</th>
                <th className="ism-th--right">Billed Till Date</th>
                <th className="ism-th--right">% Billing Done</th>
                <th className="ism-th--right">Check for Revenue against PO</th>
                <th className="ism-th--right">Check for billing against PO</th>
                <th>Linked Issue</th>
                <th>Issue Type</th>
                <th>Planned Project Start Date</th>
                <th>Planned Project End Date</th>
                <th className="ism-th--right">Efforts Planned Hours</th>
                <th className="ism-th--right">Billing Rate</th>
                <th className="ism-th--right">Total Earned Hours</th>
                <th className="ism-th--right">Hours Consumed Till Date</th>
                <th className="ism-th--right">% Completion Till Date</th>
                <th className="ism-th--right">Unbilled Till Date</th>
                <th className="ism-th--right">Cummulative Consumed Hours</th>
                <th className="ism-th--right">Updated Remaining Hours</th>
                <th className="ism-th--right">Updated Total Estimated Hours</th>
                <th className="ism-th--right">% Completetion Till Date</th>
                <th className="ism-th--right">Unbilled in INR</th>
                <th className="ism-th--right">Revenue this month FX</th>
                <th className="ism-th--right">Recovery Rate this month</th>
                <th className="ism-th--right">Revised Recovery Rate till Date</th>
                <th className="ism-th--right">Total Earned Amount ($)</th>
                <th>Linked FIN Ticket</th>
                <th className="ism-th--right">Invoiced Amount USD (Excl. Tax)</th>
                <th className="ism-th--right">Billed This Month FX</th>
                <th className="ism-th--right">Billed in Month INR</th>
              </tr>
            </thead>
            <tbody>
              {reports.map((report) => {
                const rc = Math.max(report.rows.length, 1);

                if (report.rows.length === 0) {
                  return (
                    <tr key={report.ism.key} className="ism-group-last">
                      <td className="ism-td ism-td--ism"><span className="ism-key-badge">{report.ism.key}</span></td>
                      <td className="ism-td ism-td--ism">{report.ism.summary}</td>
                      <td className="ism-td ism-td--ism">{fmtDate(report.ism.claimDate)}</td>
                      <td className="ism-td ism-td--ism ism-td--num">{report.ism.poAmount ? num(report.ism.poAmount) : '—'}</td>
                      <td className="ism-td ism-td--ism">{report.ism.status || '—'}</td>
                      <td className="ism-td ism-td--ism">{fmtDate(report.ism.resolutionDate)}</td>
                      <td className="ism-td ism-td--ism">—</td>
                      <td className="ism-td ism-td--ism ism-td--num">{num(report.ism.serviceAmount)}</td>
                      <td className="ism-td ism-td--ism">{fmtDate(report.ism.salesInvoiceDate)}</td>
                      <td className="ism-td ism-td--ism ism-td--num">{report.ism.serviceAmount ? num(report.totalInvoiced / report.ism.serviceAmount) : '—'}</td>
                      <td className="ism-td ism-td--ism ism-td--num">{num(report.totalInvoicedUSD)}</td>
                      <td className="ism-td ism-td--ism ism-td--num">{num(report.totalInvoiced)}</td>
                      <td className="ism-td ism-td--ism ism-td--num">{num(report.totalInvoicedUSD)}</td>
                      <td className="ism-td ism-td--ism ism-td--num">{report.ism.serviceAmount ? num(report.totalInvoicedUSD / report.ism.serviceAmount) : '—'}</td>
                      <td className="ism-td ism-td--ism ism-td--num">{report.ism.serviceAmount ? num((report.prjEarned + report.ausEarned) / report.ism.serviceAmount) : '—'}</td>
                      <td className="ism-td ism-td--ism ism-td--num ism-td--last-ism">{report.ism.serviceAmount ? num(report.totalInvoicedUSD / report.ism.serviceAmount) : '—'}</td>
                      <td colSpan={23} className="ism-td ism-muted">No linked issues.</td>
                    </tr>
                  );
                }

                return report.rows.map((row, i) => (
                  <tr key={`${report.ism.key}-${row.key}`} className={i === rc - 1 ? 'ism-group-last' : ''}>
                    {i === 0 && (
                      <>
                        <td rowSpan={rc} className="ism-td ism-td--ism"><span className="ism-key-badge">{report.ism.key}</span></td>
                        <td rowSpan={rc} className="ism-td ism-td--ism">{report.ism.summary}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism">{fmtDate(report.ism.claimDate)}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism ism-td--num">{report.ism.poAmount ? num(report.ism.poAmount) : '—'}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism">{report.ism.status || '—'}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism">{fmtDate(report.ism.resolutionDate)}</td>
                        <td className="ism-td">{row.prjStatus || '—'}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism ism-td--num">{num(report.ism.serviceAmount)}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism">{fmtDate(report.ism.salesInvoiceDate)}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism ism-td--num">{report.ism.serviceAmount ? num(report.totalInvoiced / report.ism.serviceAmount) : '—'}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism ism-td--num">{num(report.totalInvoicedUSD)}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism ism-td--num">{num(report.totalInvoiced)}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism ism-td--num">{num(report.totalInvoicedUSD)}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism ism-td--num">{report.ism.serviceAmount ? num(report.totalInvoicedUSD / report.ism.serviceAmount) : '—'}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism ism-td--num">{report.ism.serviceAmount ? num((report.prjEarned + report.ausEarned) / report.ism.serviceAmount) : '—'}</td>
                        <td rowSpan={rc} className="ism-td ism-td--ism ism-td--num ism-td--last-ism">{report.ism.serviceAmount ? num(report.totalInvoicedUSD / report.ism.serviceAmount) : '—'}</td>
                      </>
                    )}
                    {i > 0 && <td className="ism-td">{row.prjStatus || '—'}</td>}
                    <td className="ism-td"><span className="ism-key-badge">{row.key}</span></td>
                    <td className="ism-td"><span className="ism-type-badge">{row.issueType || '—'}</span></td>
                    <td className="ism-td">{fmtDate(row.projectStart)}</td>
                    <td className="ism-td">{fmtDate(row.projectEnd)}</td>
                    <td className="ism-td ism-td--num">{row.totalPlannedHrs ? num(row.totalPlannedHrs) : '—'}</td>
                    <td className="ism-td ism-td--num">{row.billingRate ? num(row.billingRate) : '—'}</td>
                    <td className="ism-td ism-td--num">{num(row.earned)}</td>
                    <td className="ism-td ism-td--num">{row.billingRate ? num(row.earned / row.billingRate) : '—'}</td>
                    <td className="ism-td ism-td--num">{row.totalPlannedHrs ? num(row.earned / row.totalPlannedHrs) : '—'}</td>
                    <td className="ism-td ism-td--num">{row.nonAccruedHrs ? num(row.nonAccruedHrs) : '—'}</td>
                    <td className="ism-td ism-td--num">{num(row.earned)}</td>
                    <td className="ism-td ism-td--num">{row.totalPlannedHrs ? num(row.totalPlannedHrs - row.earned) : '—'}</td>
                    <td className="ism-td ism-td--num">{row.totalPlannedHrs ? num(row.totalPlannedHrs - row.earned) : '—'}</td>
                    <td className="ism-td ism-td--num">{row.totalPlannedHrs ? num(row.earned / row.totalPlannedHrs) : '—'}</td>
                    <td className="ism-td ism-td--num">{row.nonAccruedHrs && row.billingRate ? num(row.nonAccruedHrs * row.billingRate) : '—'}</td>
                    <td className="ism-td ism-td--num">{row.earnedMonthly ? num(row.earnedMonthly) : '—'}</td>
                    <td className="ism-td ism-td--num">{(row.rowInvoicedUSD + row.totalNonAccruable) ? num(row.rowInvoicedUSD / (row.rowInvoicedUSD + row.totalNonAccruable)) : '—'}</td>
                    <td className="ism-td ism-td--num">{(row.rowInvoicedUSD + row.totalNonAccruable) ? num(row.rowInvoicedUSD / (row.rowInvoicedUSD + row.totalNonAccruable)) : '—'}</td>
                    <td className="ism-td ism-td--num">{num(row.earned)}</td>
                    <td className="ism-td">
                      {row.finIssues.length === 0
                        ? <span className="ism-muted">—</span>
                        : row.finIssues.map((f) => <div key={f.key}><span className="ism-fin-badge">{f.key}</span></div>)}
                    </td>
                    <td className="ism-td ism-td--num">
                      {row.finIssues.length === 0 ? <span className="ism-muted">—</span> : row.finIssues.map((f) => <div key={f.key}>{num(f.invoiced)}</div>)}
                    </td>
                    <td className="ism-td ism-td--num">
                      {row.finIssues.length === 0 ? <span className="ism-muted">—</span> : row.finIssues.map((f) => <div key={f.key}>{num(f.invoicedUSD)}</div>)}
                    </td>
                    <td className="ism-td ism-td--num">
                      {row.finIssues.length === 0 ? <span className="ism-muted">—</span> : row.finIssues.map((f) => <div key={f.key}>{num(f.invoiced)}</div>)}
                    </td>
                  </tr>
                ));
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

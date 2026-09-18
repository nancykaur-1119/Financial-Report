import ExcelJS from 'exceljs';

const numVal = (v) => (typeof v === 'number' ? v : Number(v) || 0);
const fmt = (n) => numVal(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtDate = (iso) => {
  if (!iso) return '—';
  try { return new Date(iso).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }); }
  catch { return iso; }
};

export async function buildExcel(reports) {
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
        finKey,
        finINR,
        finUSD,
        finINR,
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

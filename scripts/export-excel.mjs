import ExcelJS from 'exceljs';

const API_URL = process.env.API_URL || 'http://localhost:3001/api/ism-projects';

// ── Colours ────────────────────────────────────────────────────────────────
const C = {
  headerBg   : 'FF172B4D',
  ismBg      : 'FFF4F5F7', 
  white      : 'FFFFFFFF',
  borderGrey : 'FFDFE1E6',
  groupBlue  : 'FF0052CC',
};

const thinGrey  = { style: 'thin',   color: { argb: C.borderGrey } };
const mediumBlue = { style: 'medium', color: { argb: C.groupBlue  } };

function cellBorder(isGroupLast) {
  const b = isGroupLast ? mediumBlue : thinGrey;
  return { top: thinGrey, left: thinGrey, right: thinGrey, bottom: b };
}

// ── Helpers ────────────────────────────────────────────────────────────────
const fmt = (n) =>
  typeof n === 'number'
    ? n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : '—';

// ── Main ───────────────────────────────────────────────────────────────────
async function main() {
  console.log('Fetching:', API_URL);
  const res = await fetch(API_URL);
  if (!res.ok) throw new Error(`API ${res.status}: ${await res.text()}`);
  const { reports, total } = await res.json();
  console.log(`Received ${total} ISM reports`);

  const wb = new ExcelJS.Workbook();
  wb.creator  = 'ISM Report Export';
  wb.created  = new Date();

  const ws = wb.addWorksheet('ISM Report', { views: [{ state: 'frozen', ySplit: 1 }] });

  // Column definitions
  ws.columns = [
    { key: 'ismKey',         width: 16  },
    { key: 'ismSummary',     width: 42  },
    { key: 'serviceAmount',  width: 20  },
    { key: 'linkedKey',      width: 16  },
    { key: 'issueType',      width: 24  },
    { key: 'earned',         width: 22  },
    { key: 'finKey',         width: 20  },
    { key: 'invoiced',       width: 32  },
  ];

  // ── Header row ──
  const HEADERS = [
    'ISM Issue',
    'Issue Name',
    'Service Amount ($)',
    'Linked Issue',
    'Issue Type',
    'Total Earned Amount ($)',
    'Linked FIN Ticket',
    'Invoiced Amount USD (Excl. Tax)',
  ];

  const headerRow = ws.addRow(HEADERS);
  headerRow.height = 32;
  headerRow.eachCell((cell) => {
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.headerBg } };
    cell.font      = { bold: true, size: 11, color: { argb: C.white } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border    = cellBorder(false);
  });

  // ── Data rows ──
  let rowNum = 2;

  for (const report of reports) {
    const { ism, rows } = report;
    const rowCount  = Math.max(rows.length, 1);
    const startRow  = rowNum;

    if (rows.length === 0) {
      // No linked issues — single row
      addDataRow(ws, rowNum, true, true, {
        ismKey: ism.key, ismSummary: ism.summary, serviceAmount: fmt(ism.serviceAmount),
        linkedKey: '—', issueType: '—', earned: '—', finKey: '—', invoiced: '—',
      });
      rowNum++;
    } else {
      for (let i = 0; i < rows.length; i++) {
        const row      = rows[i];
        const isLast   = i === rows.length - 1;
        const finKey   = row.finIssues.length ? row.finIssues.map(f => f.key).join('\n')                                                      : '—';
        const invoiced = row.finIssues.length ? row.finIssues.map(f => fmt(f.invoiced)).join('\n')                                            : '—';

        addDataRow(ws, rowNum, i === 0, isLast, {
          ismKey:        i === 0 ? ism.key          : '',
          ismSummary:    i === 0 ? ism.summary       : '',
          serviceAmount: i === 0 ? fmt(ism.serviceAmount) : '',
          linkedKey:     row.key,
          issueType:     row.issueType || '—',
          earned:        fmt(row.earned),
          finKey,
          invoiced,
        });
        rowNum++;
      }

      // Merge ISM columns vertically when there are multiple linked rows
      if (rows.length > 1) {
        [1, 2, 3].forEach((col) => {
          ws.mergeCells(startRow, col, rowNum - 1, col);
          const cell      = ws.getCell(startRow, col);
          cell.fill       = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.ismBg } };
          cell.alignment  = {
            vertical: 'middle',
            horizontal: col === 3 ? 'right' : 'left',
            wrapText: true,
          };
          cell.border     = {
            top: thinGrey, left: thinGrey, right: thinGrey, bottom: mediumBlue,
          };
        });
      }
    }
  }

  // ── Save ──
  const date     = new Date().toISOString().slice(0, 10);
  const filename = `ism-report-${date}.xlsx`;
  await wb.xlsx.writeFile(filename);
  console.log(`Saved → ${filename}`);
}

function addDataRow(ws, rowNum, isIsmRow, isGroupLast, data) {
  const r = ws.getRow(rowNum);
  r.values = [
    data.ismKey, data.ismSummary, data.serviceAmount,
    data.linkedKey, data.issueType, data.earned,
    data.finKey, data.invoiced,
  ];
  r.height = 20;

  r.eachCell({ includeEmpty: true }, (cell, col) => {
    const border = cellBorder(isGroupLast);
    cell.border  = border;
    cell.alignment = { vertical: 'middle', wrapText: true };

    // ISM columns — grey background
    if (col <= 3) {
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF4F5F7' } };
      if (col === 3) cell.alignment = { ...cell.alignment, horizontal: 'right' };
    }
    // Numeric columns — right aligned
    if (col === 6 || col === 8) {
      cell.alignment = { ...cell.alignment, horizontal: 'right' };
    }
  });

  r.commit();
}

main().catch((err) => { console.error(err.message); process.exit(1); });

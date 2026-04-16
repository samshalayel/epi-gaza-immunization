/* ── Excel export with full styling (SheetJS) ─────────────────────────── */

// ── Style helpers ──────────────────────────────────────────────────────────
const S = {
  border: { top:{style:'thin',color:{rgb:'888888'}}, bottom:{style:'thin',color:{rgb:'888888'}},
            left:{style:'thin',color:{rgb:'888888'}}, right:{style:'thin',color:{rgb:'888888'}} },
  hdrFill:  { patternType:'solid', fgColor:{rgb:'1FA2C8'} },
  hdr2Fill: { patternType:'solid', fgColor:{rgb:'3ABCD8'} },
  blueFill: { patternType:'solid', fgColor:{rgb:'D0E8FF'} },
  yelFill:  { patternType:'solid', fgColor:{rgb:'FFF0CC'} },
  grnFill:  { patternType:'solid', fgColor:{rgb:'C6EFCE'} },
  redFill:  { patternType:'solid', fgColor:{rgb:'FFC7CE'} },
  whtFont:  { name:'Arial', sz:9, bold:true, color:{rgb:'FFFFFF'} },
  bldFont:  { name:'Arial', sz:9, bold:true, color:{rgb:'000000'} },
  nrmFont:  { name:'Arial', sz:9 },
  itaFont:  { name:'Arial', sz:9, italic:true, color:{rgb:'0D7FA5'} },
  ctr:      { horizontal:'center', vertical:'center', wrapText:true },
  lft:      { horizontal:'left',   vertical:'center', wrapText:true },
};

function mkCell(v, fill, font, align, fmt, border) {
  const t = typeof v === 'number' ? 'n' : 's';
  const cell = { v, t };
  if (fmt) cell.z = fmt;
  cell.s = {};
  if (fill)   cell.s.fill      = fill;
  if (font)   cell.s.font      = font;
  if (align)  cell.s.alignment = align;
  cell.s.border = border || S.border;
  return cell;
}

function ec(r, c) { return XLSX.utils.encode_cell({r, c}); }

// ── Main builder ────────────────────────────────────────────────────────────
function buildSheet(fac, year, d) {
  const ws = {};
  const merges = [];

  const catchment = d.catchment_pop || 0;
  const si        = d.si_percent != null ? d.si_percent : 3.2;
  const target    = Math.round(catchment * si / 100);

  // Column layout: col 0 = Antigen, cols 1-24 = Jan…Dec (Tot,Cum×12), 25 = Annual, 26 = Coverage%
  const LAST_COL  = 26;
  const ANN_COL   = 25;
  const COV_COL   = 26;
  const COL_START = 1;

  // month → { tot: col, cum: col }
  const mCol = {};
  for (let i = 0; i < 12; i++) {
    mCol[i + 1] = { tot: COL_START + i * 2, cum: COL_START + i * 2 + 1 };
  }

  // ── Row 0: Title ──────────────────────────────────────────────────────────
  ws[ec(0,0)] = mkCell(
    `Routine Immunization Monthly Monitoring  |  ${fac.name}  |  Year: ${year}`,
    S.hdrFill, { name:'Arial', sz:12, bold:true, color:{rgb:'FFFFFF'} }, S.ctr
  );
  merges.push({ s:{r:0,c:0}, e:{r:0,c:LAST_COL} });
  for (let c = 1; c <= LAST_COL; c++) ws[ec(0,c)] = mkCell('', S.hdrFill, S.whtFont, S.ctr);

  // ── Row 1: Facility info ──────────────────────────────────────────────────
  const infoItems = [
    [0,3,  `Governorate: ${fac.governorate}`],
    [4,7,  `Provider: ${fac.provider}`],
    [8,11, `Type: ${fac.type}`],
  ];
  infoItems.forEach(([s, e, txt]) => {
    ws[ec(1,s)] = mkCell(txt, S.hdr2Fill, S.whtFont, S.ctr);
    merges.push({ s:{r:1,c:s}, e:{r:1,c:e} });
    for (let c = s+1; c <= e; c++) ws[ec(1,c)] = mkCell('', S.hdr2Fill, S.whtFont, S.ctr);
  });
  for (let c = 12; c <= LAST_COL; c++) ws[ec(1,c)] = mkCell('', null, S.nrmFont, S.ctr);

  // ── Row 2: Target population ──────────────────────────────────────────────
  ws[ec(2,0)] = mkCell('Catchment Population', S.hdr2Fill, S.whtFont, S.ctr);
  merges.push({ s:{r:2,c:0}, e:{r:2,c:2} });
  ws[ec(2,1)] = mkCell('', S.hdr2Fill, S.whtFont, S.ctr);
  ws[ec(2,2)] = mkCell('', S.hdr2Fill, S.whtFont, S.ctr);
  ws[ec(2,3)] = mkCell(catchment, null, S.bldFont, S.ctr, '#,##0');

  ws[ec(2,4)] = mkCell('% Surviving Infants (SI)', S.hdr2Fill, S.whtFont, S.ctr);
  merges.push({ s:{r:2,c:4}, e:{r:2,c:6} });
  ws[ec(2,5)] = mkCell('', S.hdr2Fill, S.whtFont, S.ctr);
  ws[ec(2,6)] = mkCell('', S.hdr2Fill, S.whtFont, S.ctr);
  ws[ec(2,7)] = mkCell(si / 100, null, S.bldFont, S.ctr, '0.0%');

  ws[ec(2,8)] = mkCell('Target (Under 1 yr)', S.hdr2Fill, S.whtFont, S.ctr);
  merges.push({ s:{r:2,c:8}, e:{r:2,c:10} });
  ws[ec(2,9)]  = mkCell('', S.hdr2Fill, S.whtFont, S.ctr);
  ws[ec(2,10)] = mkCell('', S.hdr2Fill, S.whtFont, S.ctr);
  ws[ec(2,11)] = mkCell(target, null, S.bldFont, S.ctr, '#,##0');
  for (let c = 12; c <= LAST_COL; c++) ws[ec(2,c)] = mkCell('', null, S.nrmFont, S.ctr);

  // ── Row 3: Month headers (merged) ─────────────────────────────────────────
  ws[ec(3,0)] = mkCell('Antigen / Vaccine', S.hdrFill, S.whtFont, S.ctr);
  merges.push({ s:{r:3,c:0}, e:{r:4,c:0} });

  MONTHS.forEach((m, i) => {
    const tc = mCol[i+1].tot;
    ws[ec(3,tc)]   = mkCell(m, S.hdrFill, S.whtFont, S.ctr);
    ws[ec(3,tc+1)] = mkCell('', S.hdrFill, S.whtFont, S.ctr);
    merges.push({ s:{r:3,c:tc}, e:{r:3,c:tc+1} });
  });
  ws[ec(3,ANN_COL)] = mkCell('Annual Total', S.hdrFill, S.whtFont, S.ctr);
  merges.push({ s:{r:3,c:ANN_COL}, e:{r:4,c:ANN_COL} });
  ws[ec(3,COV_COL)] = mkCell('Coverage %',   S.hdrFill, S.whtFont, S.ctr);
  merges.push({ s:{r:3,c:COV_COL}, e:{r:4,c:COV_COL} });

  // ── Row 4: Tot. / Cum. sub-headers ────────────────────────────────────────
  ws[ec(4,0)] = mkCell('', S.hdrFill, S.whtFont, S.ctr);
  MONTHS.forEach((_, i) => {
    const tc = mCol[i+1].tot;
    ws[ec(4,tc)]   = mkCell('Tot.', S.hdr2Fill, S.whtFont, S.ctr);
    ws[ec(4,tc+1)] = mkCell('Cum.', S.hdr2Fill, S.whtFont, S.ctr);
  });
  ws[ec(4,ANN_COL)] = mkCell('', S.hdrFill, S.whtFont, S.ctr);
  ws[ec(4,COV_COL)] = mkCell('', S.hdrFill, S.whtFont, S.ctr);

  // ── Rows 5+: Antigen data ─────────────────────────────────────────────────
  const DATA_START = 5;
  const annuals = {};

  ANTIGENS.forEach((ag, ai) => {
    const r    = DATA_START + ai;
    const fill = ai % 2 === 0 ? S.blueFill : S.yelFill;

    ws[ec(r,0)] = mkCell(ag.label, fill, S.bldFont, S.lft);

    let cum = 0, annual = 0;
    for (let m = 1; m <= 12; m++) {
      const v = d[`${ag.key}_${m}`] || 0;
      cum += v; annual += v;
      ws[ec(r, mCol[m].tot)] = mkCell(v,   fill, S.nrmFont, S.ctr, '#,##0');
      ws[ec(r, mCol[m].cum)] = mkCell(cum, fill, S.itaFont, S.ctr, '#,##0');
    }
    annuals[ag.key] = annual;

    const cov = target > 0 ? annual / target : 0;
    ws[ec(r, ANN_COL)] = mkCell(annual, fill, S.bldFont, S.ctr, '#,##0');

    // Coverage cell with color
    let covFill = fill;
    if (target > 0) {
      const p = cov * 100;
      if (p >= 100) covFill = S.grnFill;
      else if (p >= 75) covFill = { patternType:'solid', fgColor:{rgb:'EBFADB'} };
      else if (p >= 50) covFill = { patternType:'solid', fgColor:{rgb:'FFEB84'} };
      else if (p >= 45) covFill = { patternType:'solid', fgColor:{rgb:'FFD580'} };
      else if (annual > 0) covFill = S.redFill;
    }
    ws[ec(r, COV_COL)] = mkCell(target > 0 ? cov : '', covFill, S.bldFont, S.ctr, '0.0%');
  });

  // ── Drop-out header row ───────────────────────────────────────────────────
  const DO_HDR_R = DATA_START + ANTIGENS.length + 1;
  ws[ec(DO_HDR_R, 0)] = mkCell('Drop-out Rates', S.hdrFill, S.whtFont, S.lft);
  merges.push({ s:{r:DO_HDR_R,c:0}, e:{r:DO_HDR_R,c:LAST_COL} });
  for (let c = 1; c <= LAST_COL; c++) ws[ec(DO_HDR_R,c)] = mkCell('', S.hdrFill, S.whtFont, S.ctr);

  // ── Drop-out rows ─────────────────────────────────────────────────────────
  DROPOUTS.forEach((drp, di) => {
    const r    = DO_HDR_R + 1 + di;
    const fill = di % 2 === 0 ? S.blueFill : S.yelFill;

    ws[ec(r,0)] = mkCell(drp.label, fill, S.bldFont, S.lft);

    // per-month cumulative dropout
    let cumNum = 0, cumDen = 0;
    for (let m = 1; m <= 12; m++) {
      cumNum += d[`${drp.num}_${m}`] || 0;
      cumDen += d[`${drp.den}_${m}`] || 0;
      const pct = cumNum > 0 ? (cumNum - cumDen) / cumNum : 0;
      const tc = mCol[m].tot, cc = mCol[m].cum;
      ws[ec(r,tc)] = mkCell(cumNum > 0 ? pct : '', fill, S.nrmFont, S.ctr, '0.0%');
      ws[ec(r,cc)] = mkCell('', fill, S.nrmFont, S.ctr);
      merges.push({ s:{r,c:tc}, e:{r,c:cc} });
    }

    // Annual dropout
    const nA = annuals[drp.num] || 0, dA = annuals[drp.den] || 0;
    const annPct = nA > 0 ? (nA - dA) / nA : 0;
    ws[ec(r, ANN_COL)] = mkCell(nA > 0 ? annPct : '', fill, S.bldFont, S.ctr, '0.0%');

    // Status
    let statusFill = fill;
    let statusTxt  = '';
    if (nA > 0) {
      statusTxt  = annPct > 0.10 ? 'HIGH ⚠' : 'OK ✓';
      statusFill = annPct > 0.10
        ? { patternType:'solid', fgColor:{rgb:'FFC7CE'} }
        : { patternType:'solid', fgColor:{rgb:'C6EFCE'} };
    }
    ws[ec(r, COV_COL)] = mkCell(statusTxt, statusFill, S.bldFont, S.ctr);
  });

  // ── Coverage legend ───────────────────────────────────────────────────────
  const LEG_R = DO_HDR_R + DROPOUTS.length + 3;
  ws[ec(LEG_R,0)] = mkCell('Coverage Achievement Legend:', null, S.bldFont, S.lft, null, {});
  merges.push({ s:{r:LEG_R,c:0}, e:{r:LEG_R,c:2} });
  ws[ec(LEG_R,1)] = mkCell('', null, S.bldFont, S.ctr, null, {});
  ws[ec(LEG_R,2)] = mkCell('', null, S.bldFont, S.ctr, null, {});

  const legends = [
    [3,  '≥100%', '00B050', 'FFFFFF'],
    [4,  '≥75%',  '92D050', '000000'],
    [5,  '≥50%',  'FFEB84', '000000'],
    [6,  '≥45%',  'FF9900', 'FFFFFF'],
    [7,  '<45%',  'FF0000', 'FFFFFF'],
  ];
  legends.forEach(([c, lbl, bg, fg]) => {
    ws[ec(LEG_R,c)] = mkCell(lbl,
      { patternType:'solid', fgColor:{rgb:bg} },
      { name:'Arial', sz:9, bold:true, color:{rgb:fg} },
      S.ctr
    );
  });

  // ── Formula notes ─────────────────────────────────────────────────────────
  const NOTE_R = LEG_R + 2;
  ws[ec(NOTE_R,0)] = mkCell(
    'Formulas:  Drop-out # = First Dose − Last Dose  |  Drop-out % = Drop-out # ÷ First Dose × 100  |  Coverage % = Annual Doses ÷ Target × 100',
    null,
    { name:'Arial', sz:8, italic:true, color:{rgb:'444444'} },
    S.lft, null, {}
  );
  merges.push({ s:{r:NOTE_R,c:0}, e:{r:NOTE_R,c:LAST_COL} });

  // ── Sheet range, merges, col widths ──────────────────────────────────────
  ws['!ref'] = XLSX.utils.encode_range({ s:{r:0,c:0}, e:{r:NOTE_R,c:LAST_COL} });
  ws['!merges'] = merges;
  ws['!cols'] = [
    { wch:22 },
    ...Array.from({length:24}, () => ({wch:7})),
    { wch:12 }, { wch:12 }
  ];
  ws['!rows'] = [
    { hpt:22 }, {}, {}, { hpt:18 }, { hpt:14 }
  ];

  return ws;
}

// ── Public API ──────────────────────────────────────────────────────────────
function exportSingleFacility(fac, year) {
  const wb = XLSX.utils.book_new();
  const d  = JSON.parse(localStorage.getItem(`data_${fac.id}_${year}`) || '{}');
  XLSX.utils.book_append_sheet(wb, buildSheet(fac, year, d), truncSheet(fac.name));
  XLSX.writeFile(wb, `EPI_${fac.name.replace(/[\/\\:*?"<>|]/g,'-')}_${year}.xlsx`);
}

function exportAllFacilities(facilities, year) {
  const wb = XLSX.utils.book_new();
  facilities.forEach(fac => {
    const d = JSON.parse(localStorage.getItem(`data_${fac.id}_${year}`) || '{}');
    XLSX.utils.book_append_sheet(wb, buildSheet(fac, year, d), truncSheet(fac.name));
  });
  XLSX.writeFile(wb, `EPI_All_Facilities_${year}.xlsx`);
}

function truncSheet(name) {
  return name.replace(/[\/\\:*?"<>|]/g,'-').substring(0,31);
}

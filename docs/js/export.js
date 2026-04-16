/* Excel export using SheetJS */

function exportSingleFacility(fac, year) {
  const wb = XLSX.utils.book_new();
  const saved = localStorage.getItem(`data_${fac.id}_${year}`);
  const d = saved ? JSON.parse(saved) : {};
  const ws = buildSheet(fac, year, d);
  XLSX.utils.book_append_sheet(wb, ws, truncSheet(fac.name));
  XLSX.writeFile(wb, `EPI_${fac.name.replace(/[\/\\:*?"<>|]/g,'-')}_${year}.xlsx`);
}

function exportAllFacilities(facilities, year) {
  const wb = XLSX.utils.book_new();
  facilities.forEach(fac => {
    const saved = localStorage.getItem(`data_${fac.id}_${year}`);
    const d = saved ? JSON.parse(saved) : {};
    const ws = buildSheet(fac, year, d);
    XLSX.utils.book_append_sheet(wb, ws, truncSheet(fac.name));
  });
  XLSX.writeFile(wb, `EPI_All_Facilities_${year}.xlsx`);
}

function truncSheet(name) {
  return name.replace(/[\/\\:*?"<>|]/g,'-').substring(0,31);
}

function buildSheet(fac, year, d) {
  const catchment = d.catchment_pop || 0;
  const si        = d.si_percent != null ? d.si_percent : 3.2;
  const target    = Math.round(catchment * si / 100);

  const aoa = []; // array of arrays

  // Row 1: Title
  aoa.push([`Routine Immunization Monthly Monitoring | ${fac.name} | Year: ${year}`]);

  // Row 2: Facility info
  aoa.push([`Governorate: ${fac.governorate}`, '', '', `Provider: ${fac.provider}`, '', '', `Type: ${fac.type}`]);

  // Row 3: Target population
  aoa.push(['Catchment Population', catchment, '', '% SI', si/100, '', 'Target (U1)', target]);

  // Row 4-5: Headers
  const hdr1 = ['Antigen'];
  const hdr2 = [''];
  MONTHS.forEach(m => { hdr1.push(m,''); hdr2.push('Monthly','Cumul.'); });
  hdr1.push('Annual','Coverage %');
  hdr2.push('','');
  aoa.push(hdr1, hdr2);

  // Data rows
  const annuals = {};
  ANTIGENS.forEach(ag => {
    const row = [ag.label];
    let cum = 0, annual = 0;
    for (let m=1;m<=12;m++) {
      const v = d[`${ag.key}_${m}`]||0;
      cum += v; annual += v;
      row.push(v, cum);
    }
    annuals[ag.key] = annual;
    const cov = target > 0 ? parseFloat((annual/target*100).toFixed(1)) : 0;
    row.push(annual, cov/100);
    aoa.push(row);
  });

  // Blank row
  aoa.push([]);

  // Dropout header
  aoa.push(['Drop-out Rates']);

  // Dropout rows
  DROPOUTS.forEach(dr => {
    const row = [dr.label];
    // Recalculate cumulative per month
    const cumNum = {}, cumDen = {};
    let cn=0, cd=0;
    for(let m=1;m<=12;m++){
      cn += d[`${dr.num}_${m}`]||0;
      cd += d[`${dr.den}_${m}`]||0;
      cumNum[m]=cn; cumDen[m]=cd;
      const pct = cn>0 ? parseFloat(((cn-cd)/cn*100).toFixed(1)) : null;
      // Monthly: show as single merged cell (just put in first, blank in second)
      row.push(pct !== null ? pct/100 : '', '');
    }
    const nA = annuals[dr.num]||0, dA = annuals[dr.den]||0;
    const annPct = nA>0 ? parseFloat(((nA-dA)/nA*100).toFixed(1)) : 0;
    row.push(annPct/100, nA>0 ? (annPct>10?'HIGH ⚠':'OK ✓') : '–');
    aoa.push(row);
  });

  // Legend
  aoa.push([]);
  aoa.push(['Coverage Legend:', '≥100% Excellent','≥75% Good','≥50% Fair','≥45% Weak','<45% Critical']);
  aoa.push(['Formulas: Drop-out% = (First Dose − Last Dose) ÷ First Dose × 100 | Coverage% = Annual ÷ Target × 100']);

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Column widths
  ws['!cols'] = [
    {wch:22},
    ...Array.from({length:26}, () => ({wch:8})),
    {wch:12},{wch:12}
  ];

  // Merge title row
  const lastCol = 1 + 24 + 2;
  ws['!merges'] = ws['!merges'] || [];
  ws['!merges'].push({s:{r:0,c:0}, e:{r:0,c:lastCol}});

  // Format percentage cells for coverage and dropout
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({r:R, c:C});
      if (!ws[addr]) continue;
      if (typeof ws[addr].v === 'number') {
        // Check if it looks like a percentage (between 0 and 2 with decimals from our /100 division)
        const row_offset = R - 5; // data starts at row index 5
        if (row_offset >= 0) {
          const isLastCol = (C === lastCol || C === lastCol-1);
          if (isLastCol && ws[addr].v <= 2 && ws[addr].v >= -1) {
            ws[addr].z = '0.0%';
          }
        }
        if (C === 0) continue;
        ws[addr].z = C >= 27 && ws[addr].v <= 2 ? '0.0%' : '#,##0';
      }
    }
  }

  return ws;
}

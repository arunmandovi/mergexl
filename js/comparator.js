// ══════════════════════════════════════════════════════════════════════════
//  COMPARATOR
// ══════════════════════════════════════════════════════════════════════════
(function () {
  const ROW_H  = 34;
  const BUFFER = 20;

  let cFiles      = [];
  let cMode       = 'row';
  let cAllResults = [];
  const cScrollers = {};

  const dropZone  = document.getElementById('c-dropZone');
  const fileInput = document.getElementById('c-fileInput');

  dropZone.addEventListener('click',    () => fileInput.click());
  fileInput.addEventListener('change',  e  => cAddFiles(Array.from(e.target.files)));
  dropZone.addEventListener('dragover', e  => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop',     e  => {
    e.preventDefault(); dropZone.classList.remove('drag-over');
    cAddFiles(Array.from(e.dataTransfer.files).filter(f => /\.(xlsx|xls|csv)$/i.test(f.name)));
  });

  function cAddFiles(nf) {
    nf.forEach(f => { if (!cFiles.find(x => x.name === f.name && x.size === f.size)) cFiles.push(f); });
    cRenderTags();
  }
  window.cRemoveFile = function (i) { cFiles.splice(i, 1); cRenderTags(); };

  function cRenderTags() {
    document.getElementById('c-fileTags').innerHTML = cFiles.map((f, i) =>
      `<div class="file-tag">📄 ${escHtml(f.name)}<button onclick="cRemoveFile(${i})">✕</button></div>`
    ).join('');
    document.getElementById('c-compareBtn').disabled = cFiles.length < 2;
  }

  window.cSelectMode = function (m) {
    cMode = m;
    document.getElementById('c-modeRow').classList.toggle('c-sel', m === 'row');
    document.getElementById('c-modeKey').classList.toggle('c-sel', m === 'key');
    document.getElementById('c-keyWrap').style.display = m === 'key' ? 'block' : 'none';
  };

  function cReadWb(file) {
    // CSV: plain-text read so original casing ("06-Jun-26") is preserved exactly
    if (/\.csv$/i.test(file.name)) return readCsvAsWorkbook(file);
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = e => {
        try {
          const wb = XLSX.read(e.target.result, { type: 'binary', cellNF: true, cellDates: false, raw: true });
          res(wb);
        } catch (e) { rej(e); }
      };
      r.onerror = rej;
      r.readAsBinaryString(file);
    });
  }

  function cSheetToMatrix(wb) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
  }

  window.cCompareFiles = async function () {
    const btn = document.getElementById('c-compareBtn');
    btn.disabled = true; btn.textContent = '⏳ Comparing...';
    document.getElementById('c-results').style.display = 'none';
    Object.keys(cScrollers).forEach(k => delete cScrollers[k]);

    try {
      const matrices = [];
      for (const f of cFiles) { const wb = await cReadWb(f); matrices.push(cSheetToMatrix(wb)); }

      const pairs = [];
      for (let a = 0; a < cFiles.length - 1; a++)
        for (let b = a + 1; b < cFiles.length; b++)
          pairs.push({ a, b, ma: matrices[a], mb: matrices[b] });

      const keyCol    = document.getElementById('c-keyCol').value.trim();
      cAllResults     = pairs.map(p => cComparePair(p.ma, p.mb, p.a, p.b, keyCol));
      cRenderResults(cAllResults);
    } catch (err) { alert('Error: ' + err.message); }

    btn.disabled = false; btn.textContent = '🔍 Compare Files';
  };

  function cComparePair(ma, mb, ai, bi, keyCol) {
    const headersA   = ma[0] || [];
    const headersB   = mb[0] || [];
    const allHeaders = [...new Set([...headersA, ...headersB])];
    const rowsA = ma.slice(1), rowsB = mb.slice(1);
    const idxA = h => headersA.indexOf(h);
    const idxB = h => headersB.indexOf(h);
    let compRows = [];

    if (cMode === 'key' && keyCol) {
      const kiA = idxA(keyCol), kiB = idxB(keyCol);
      if (kiA === -1 || kiB === -1) return { ai, bi, error: `Key column "${keyCol}" not found.`, allHeaders };
      const mapA = {}, mapB = {};
      rowsA.forEach(r => mapA[r[kiA]] = r); rowsB.forEach(r => mapB[r[kiB]] = r);
      [...new Set([...Object.keys(mapA), ...Object.keys(mapB)])].forEach(k => {
        const ra = mapA[k] || [], rb = mapB[k] || [], diffs = {};
        allHeaders.forEach(h => { if (String(ra[idxA(h)] ?? '') !== String(rb[idxB(h)] ?? '')) diffs[h] = true; });
        compRows.push({ key: k, ra, rb, diffs, onlyA: !mapB[k], onlyB: !mapA[k] });
      });
    } else {
      for (let i = 0; i < Math.max(rowsA.length, rowsB.length); i++) {
        const ra = rowsA[i] || [], rb = rowsB[i] || [], diffs = {};
        allHeaders.forEach(h => { if (String(ra[idxA(h)] ?? '') !== String(rb[idxB(h)] ?? '')) diffs[h] = true; });
        compRows.push({ rowNum: i + 1, ra, rb, diffs, isDiff: Object.keys(diffs).length > 0,
          onlyA: i >= rowsB.length, onlyB: i >= rowsA.length });
      }
    }
    const totalRows = compRows.length;
    const diffRows  = compRows.filter(r => Object.keys(r.diffs).length > 0 || r.onlyA || r.onlyB).length;
    return { ai, bi, allHeaders, headersA, headersB, compRows, totalRows, diffRows, sameRows: totalRows - diffRows };
  }

  function cRenderResults(pairResults) {
    document.getElementById('c-results').style.display = 'block';

    const allSame = pairResults.every(p => !p.error && p.diffRows === 0);
    const badge   = document.getElementById('c-overallBadge');
    badge.className = 'overall-badge ' + (allSame ? 'match' : 'mismatch');
    badge.innerHTML = allSame ? '✅ All files are <strong>identical</strong>' : '⚠️ Files have <strong>differences</strong>';

    let totTotal = 0, totDiff = 0, totSame = 0;
    pairResults.forEach(p => { if (!p.error) { totTotal += p.totalRows; totDiff += p.diffRows; totSame += p.sameRows; } });
    document.getElementById('c-summaryCards').innerHTML = `
      <div class="sumcard info"><div class="num">${totTotal}</div><div class="lbl">Total Rows</div></div>
      <div class="sumcard identical"><div class="num">${totSame}</div><div class="lbl">Identical</div></div>
      <div class="sumcard different"><div class="num">${totDiff}</div><div class="lbl">Different</div></div>`;

    const tabsBar   = document.getElementById('c-tabsBar');
    const tabPanels = document.getElementById('c-tabPanels');
    tabsBar.innerHTML = ''; tabPanels.innerHTML = '';

    pairResults.forEach((p, pi) => {
      const label = `${cFiles[p.ai].name.replace(/\.[^.]+$/, '')} vs ${cFiles[p.bi].name.replace(/\.[^.]+$/, '')}`;
      const tabEl = document.createElement('div');
      tabEl.className = 'ctab' + (pi === 0 ? ' active' : '');
      tabEl.textContent = label;
      tabEl.onclick = () => cSwitchTab(pi);
      tabsBar.appendChild(tabEl);

      const panel = document.createElement('div');
      panel.className = 'ctab-panel' + (pi === 0 ? ' active' : '');
      panel.id = 'c-panel-' + pi;
      panel.innerHTML = cBuildPanelShell(p, pi);
      tabPanels.appendChild(panel);
    });
    requestAnimationFrame(() => cActivateScroller(0));
  }

  function cBuildPanelShell(p, pi) {
    if (p.error) return `<div class="empty-msg">❌ ${escHtml(p.error)}</div>`;
    const fnA = cFiles[p.ai].name, fnB = cFiles[p.bi].name;
    return `
      <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:14px">
        <div class="sumcard info"      style="flex:none;padding:9px 16px"><div class="num" style="font-size:1.3rem">${p.totalRows}</div><div class="lbl">Rows</div></div>
        <div class="sumcard identical" style="flex:none;padding:9px 16px"><div class="num" style="font-size:1.3rem">${p.sameRows}</div><div class="lbl">Same</div></div>
        <div class="sumcard different" style="flex:none;padding:9px 16px"><div class="num" style="font-size:1.3rem">${p.diffRows}</div><div class="lbl">Different</div></div>
        <button class="btn-dl" id="c-dlbtn-${pi}" onclick="cDownloadExcel(${pi})">⬇ Download Excel</button>
      </div>
      <div class="legend">
        <span><span class="dot same"></span> Identical row</span>
        <span><span class="dot diff-row"></span> Row has differences</span>
        <span><span class="dot diff-cell"></span> Changed cell</span>
      </div>
      <div class="vscroll-wrap" id="c-vs-${pi}">
        <table>
          <thead>
            <tr>
              <th style="width:44px">#</th>
              ${p.allHeaders.map(h =>
                `<th class="col-a">${escHtml(String(h))}<br><span style="font-size:.68rem;font-weight:400">📄 ${escHtml(fnA)}</span></th>` +
                `<th class="col-b">${escHtml(String(h))}<br><span style="font-size:.68rem;font-weight:400">📄 ${escHtml(fnB)}</span></th>`
              ).join('')}
              <th style="width:96px">Status</th>
            </tr>
          </thead>
          <tbody id="c-vb-${pi}"></tbody>
        </table>
      </div>`;
  }

  function cSwitchTab(idx) {
    document.querySelectorAll('.ctab').forEach((t, i)       => t.classList.toggle('active', i === idx));
    document.querySelectorAll('.ctab-panel').forEach((p, i) => p.classList.toggle('active', i === idx));
    requestAnimationFrame(() => cActivateScroller(idx));
  }

  function cActivateScroller(pi) {
    const p = cAllResults[pi]; if (!p || p.error) return;
    if (!cScrollers[pi]) {
      const el = document.getElementById('c-vs-' + pi); if (!el) return;
      cScrollers[pi] = { lastStart: -1, lastEnd: -1 };
      el.addEventListener('scroll', () => cPaintRows(pi), { passive: true });
    }
    cPaintRows(pi);
  }

  function cPaintRows(pi) {
    const p  = cAllResults[pi];
    const el = document.getElementById('c-vs-' + pi);
    const tb = document.getElementById('c-vb-' + pi);
    if (!el || !tb || !p) return;

    const scrollTop = el.scrollTop;
    const viewH     = el.clientHeight || 540;
    const total     = p.compRows.length;

    const startIdx = Math.max(0,     Math.floor(scrollTop / ROW_H) - BUFFER);
    const endIdx   = Math.min(total, Math.ceil((scrollTop + viewH) / ROW_H) + BUFFER);

    const st = cScrollers[pi];
    if (st.lastStart === startIdx && st.lastEnd === endIdx) return;
    st.lastStart = startIdx; st.lastEnd = endIdx;

    const topPad  = startIdx * ROW_H;
    const botPad  = Math.max(0, (total - endIdx)) * ROW_H;
    const colSpan = p.allHeaders.length * 2 + 2;

    let html = '';
    if (topPad > 0) html += `<tr><td colspan="${colSpan}" style="height:${topPad}px;padding:0;border:none;background:transparent"></td></tr>`;
    for (let i = startIdx; i < endIdx; i++) html += cBuildRowHtml(p, i);
    if (botPad > 0) html += `<tr><td colspan="${colSpan}" style="height:${botPad}px;padding:0;border:none;background:transparent"></td></tr>`;
    tb.innerHTML = html;
  }

  function cBuildRowHtml(p, i) {
    const row     = p.compRows[i];
    const hasDiff = Object.keys(row.diffs).length > 0 || row.onlyA || row.onlyB;
    const cls     = hasDiff ? 'row-diff' : 'row-same';
    const label   = row.onlyA ? '⬅ Only A' : row.onlyB ? '➡ Only B' : hasDiff ? '⚠ Diff' : '✅ Same';

    let html = `<tr class="${cls}" style="height:${ROW_H}px"><td class="row-num">${escHtml(String(row.rowNum ?? row.key))}</td>`;
    p.allHeaders.forEach(h => {
      const ia = p.headersA.indexOf(h), ib = p.headersB.indexOf(h);
      const va = ia >= 0 ? String(row.ra[ia] ?? '') : '';
      const vb = ib >= 0 ? String(row.rb[ib] ?? '') : '';
      const d  = row.diffs[h];
      html += `<td class="val-a${d ? ' cell-diff' : ''}" title="${escHtml(va)}">${escHtml(va)}</td>`;
      html += `<td class="val-b${d ? ' cell-diff' : ''}" title="${escHtml(vb)}">${escHtml(vb)}</td>`;
    });
    html += `<td>${label}</td></tr>`;
    return html;
  }

  window.cDownloadExcel = async function (pi) {
    const p   = cAllResults[pi]; if (!p || p.error) return;
    const btn = document.getElementById('c-dlbtn-' + pi);
    if (btn) { btn.disabled = true; btn.textContent = '⏳ Generating…'; }

    try {
      const fnA = cFiles[p.ai].name, fnB = cFiles[p.bi].name;
      const wb  = new ExcelJS.Workbook();
      const ws  = wb.addWorksheet('Comparison');

      const hdrFillA  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBEE3F8' } };
      const hdrFillB  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9D8FD' } };
      const hdrFill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      const diffFill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFED7D7' } };
      const onlyAFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF8FF' } };
      const onlyBFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0FFF4' } };
      const sameFill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
      const border    = s => ({ style: s, color: { argb: 'FFD1D5DB' } });
      const cellBdr   = { top: border('thin'), left: border('thin'), bottom: border('thin'), right: border('thin') };

      // Banner row
      const banner = [''];
      p.allHeaders.forEach(() => { banner.push(fnA); banner.push(fnB); });
      banner.push('');
      const bannerRow = ws.addRow(banner);
      bannerRow.height = 20;
      p.allHeaders.forEach((_, hi) => {
        const colA = 2 + hi * 2;
        try { ws.mergeCells(1, colA, 1, colA + 1); } catch (e) {}
        const cell      = bannerRow.getCell(colA);
        cell.value      = `📄 ${fnA}  vs  📄 ${fnB}`;
        cell.font       = { bold: true, size: 10, color: { argb: 'FF2B6CB0' } };
        cell.fill       = hdrFillA;
        cell.alignment  = { horizontal: 'center', vertical: 'middle' };
      });

      // Header row
      const colHdrs = ['#'];
      p.allHeaders.forEach(h => { colHdrs.push(`${h} (A)`); colHdrs.push(`${h} (B)`); });
      colHdrs.push('Status');
      const hRow = ws.addRow(colHdrs);
      hRow.height = 22;
      hRow.eachCell((cell, col) => {
        cell.font = { bold: true, size: 10 }; cell.border = cellBdr;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        const idx = col - 2;
        if (col === 1 || col === colHdrs.length) cell.fill = hdrFill;
        else cell.fill = idx % 2 === 0 ? hdrFillA : hdrFillB;
      });

      ws.views = [{ state: 'frozen', ySplit: 2, xSplit: 1 }];

      // Data rows
      p.compRows.forEach(row => {
        const hasDiff = Object.keys(row.diffs).length > 0 || row.onlyA || row.onlyB;
        const label   = row.onlyA ? 'Only in A' : row.onlyB ? 'Only in B' : hasDiff ? 'Different' : 'Same';
        const rowData = [row.rowNum ?? row.key];
        p.allHeaders.forEach(h => {
          const ia = p.headersA.indexOf(h), ib = p.headersB.indexOf(h);
          rowData.push(ia >= 0 ? (row.ra[ia] ?? '') : '');
          rowData.push(ib >= 0 ? (row.rb[ib] ?? '') : '');
        });
        rowData.push(label);
        const exRow    = ws.addRow(rowData);
        exRow.height   = 18;
        const baseFill = row.onlyA ? onlyAFill : row.onlyB ? onlyBFill : sameFill;
        exRow.eachCell(cell => { cell.fill = baseFill; cell.border = cellBdr; cell.alignment = { vertical: 'middle' }; });
        if (!row.onlyA && !row.onlyB) {
          p.allHeaders.forEach((h, hi) => {
            if (row.diffs[h]) {
              exRow.getCell(2 + hi * 2).fill = diffFill;
              exRow.getCell(3 + hi * 2).fill = diffFill;
            }
          });
        }
        const sc = exRow.getCell(colHdrs.length);
        if (row.onlyA || row.onlyB) sc.font = { bold: true, color: { argb: 'FF2B6CB0' } };
        else if (hasDiff)           sc.font = { bold: true, color: { argb: 'FF9B2C2C' } };
        else                        sc.font = { color: { argb: 'FF276749' } };
      });

      // Column widths
      ws.getColumn(1).width = 8;
      p.allHeaders.forEach((h, hi) => {
        const w = Math.min(30, Math.max(13, String(h).length + 4));
        ws.getColumn(2 + hi * 2).width = w; ws.getColumn(3 + hi * 2).width = w;
      });
      ws.getColumn(colHdrs.length).width = 12;

      const buffer = await wb.xlsx.writeBuffer();
      const blob   = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url    = URL.createObjectURL(blob);
      const a      = document.createElement('a');
      a.href = url; a.download = `comparison_${fnA.replace(/\.[^.]+$/, '')}_vs_${fnB.replace(/\.[^.]+$/, '')}.xlsx`;
      a.click(); URL.revokeObjectURL(url);
    } catch (err) { alert('Download failed: ' + err.message); }

    if (btn) { btn.disabled = false; btn.textContent = '⬇ Download Excel'; }
  };
})();

// ══════════════════════════════════════════════════════════════════════════
//  MERGER
// ══════════════════════════════════════════════════════════════════════════
(function () {
  let mFiles     = [];
  let mMergeMode = 'combine';
  let mSelFmt    = 'xlsx';

  const dropZone  = document.getElementById('m-dropZone');
  const fileInput = document.getElementById('m-fileInput');

  dropZone.addEventListener('click',    () => fileInput.click());
  fileInput.addEventListener('change',  e  => mAddFiles(Array.from(e.target.files)));
  dropZone.addEventListener('dragover', e  => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop',     e  => {
    e.preventDefault(); dropZone.classList.remove('drag-over');
    mAddFiles(Array.from(e.dataTransfer.files).filter(f => /\.(xlsx|xls|csv)$/i.test(f.name)));
  });

  function mAddFiles(nf) {
    nf.forEach(f => { if (!mFiles.find(x => x.name === f.name && x.size === f.size)) mFiles.push(f); });
    mRenderFiles();
  }
  window.mRemoveFile = function (i) { mFiles.splice(i, 1); mRenderFiles(); };

  function mRenderFiles() {
    const container = document.getElementById('m-fileItems');
    const listEl    = document.getElementById('m-fileList');
    document.getElementById('m-fileCount').textContent = mFiles.length;
    if (!mFiles.length) { listEl.style.display = 'none'; document.getElementById('m-mergeBtn').disabled = true; return; }
    listEl.style.display = 'block';
    document.getElementById('m-mergeBtn').disabled = false;
    container.innerHTML = mFiles.map((f, i) => `
      <div class="file-item">
        <div class="fi-name"><span>✅</span><span>${escHtml(f.name)}</span></div>
        <div style="display:flex;align-items:center;gap:11px">
          <span class="fi-size">${(f.size / 1024).toFixed(1)} KB</span>
          <button class="fi-rm" onclick="mRemoveFile(${i})">✕</button>
        </div>
      </div>`).join('');
  }

  window.mSelectOption = function (mode) {
    mMergeMode = mode;
    document.getElementById('m-opt1').classList.toggle('m-sel', mode === 'combine');
    document.getElementById('m-opt2').classList.toggle('m-sel', mode === 'sheets');
    document.getElementById('m-opt3').classList.toggle('m-sel', mode === 'flatten');
  };

  window.mSelectFmt = function (fmt) {
    mSelFmt = fmt;
    document.getElementById('m-fmtXlsx').classList.toggle('m-sel', fmt === 'xlsx');
    document.getElementById('m-fmtXls').classList.toggle('m-sel',  fmt === 'xls');
    document.getElementById('m-fmtCsv').classList.toggle('m-sel',  fmt === 'csv');
  };

  function mSetProgress(pct) {
    document.getElementById('m-progWrap').style.display = 'block';
    document.getElementById('m-progFill').style.width   = pct + '%';
  }
  function mShowStatus(msg, type) {
    const el = document.getElementById('m-status');
    el.textContent = msg; el.className = 'm-status ' + type; el.style.display = 'block';
  }
  function mDownloadBlob(data, filename, mime) {
    const blob = new Blob([data], { type: mime + ';charset=utf-8;' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = filename; a.click(); URL.revokeObjectURL(url);
  }

  function mReadFile(file) {
    // CSV: read as plain text so values like "06-Jun-26" are never touched by SheetJS
    if (/\.csv$/i.test(file.name)) return readCsvAsWorkbook(file);
    // XLSX / XLS: read as binary to preserve cell styles, number formats, column widths
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const wb = XLSX.read(e.target.result, { type: 'binary', cellStyles: true, cellNF: true, cellDates: false, raw: true });
          resolve(wb);
        } catch (err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsBinaryString(file);
    });
  }

  function mGetHeaders(ws, range) {
    const h = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c })];
      h.push(cell ? String(cell.v ?? '') : '');
    }
    return h;
  }

  window.mMergeFiles = async function () {
    if (!mFiles.length) return;
    const btn = document.getElementById('m-mergeBtn');
    btn.disabled = true; document.getElementById('m-status').style.display = 'none';
    mSetProgress(10);
    const fmt      = mSelFmt;
    const baseName = document.getElementById('m-outputName').value.trim() || 'merged_output';

    try {
      if (mMergeMode === 'combine') {

        if (fmt === 'csv') {
          let csvRows = [], parentHeaders = null;
          for (let i = 0; i < mFiles.length; i++) {
            mSetProgress(10 + i / mFiles.length * 70);
            const wb    = await mReadFile(mFiles[i]);
            const ws    = wb.Sheets[wb.SheetNames[0]];
            const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
            const startR = (i > 0 && parentHeaders) ? range.s.r + 1 : range.s.r;
            for (let r = startR; r <= range.e.r; r++) {
              const row = [];
              for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                row.push(cell ? String(cell.v ?? '') : '');
              }
              csvRows.push(row);
            }
            if (i === 0) {
              parentHeaders = [];
              for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c })];
                parentHeaders.push(cell ? String(cell.v ?? '') : '');
              }
            }
          }
          const csv = csvRows.map(row => row.map(v => {
            const s = String(v);
            return (s.includes(',') || s.includes('"') || s.includes('\n')) ? '"' + s.replace(/"/g, '""') + '"' : s;
          }).join(',')).join('\r\n');
          mDownloadBlob(csv, baseName + '.csv', 'text/csv');

        } else {
          // XLSX / XLS — direct cell-object copy to preserve all formatting + dates
          const parentWb    = await mReadFile(mFiles[0]);
          mSetProgress(15);
          const parentWs    = parentWb.Sheets[parentWb.SheetNames[0]];
          const parentRange = XLSX.utils.decode_range(parentWs['!ref'] || 'A1');
          const parentHdrs  = mGetHeaders(parentWs, parentRange);

          let currentLastRow = parentRange.e.r, maxCol = parentRange.e.c;

          for (let i = 1; i < mFiles.length; i++) {
            mSetProgress(15 + i / mFiles.length * 72);
            const wb  = await mReadFile(mFiles[i]);
            const ws  = wb.Sheets[wb.SheetNames[0]];
            const rng = XLSX.utils.decode_range(ws['!ref'] || 'A1');
            const srcHdrs    = mGetHeaders(ws, rng);
            const startSrcRow = JSON.stringify(srcHdrs) === JSON.stringify(parentHdrs) ? rng.s.r + 1 : rng.s.r;

            for (let r = startSrcRow; r <= rng.e.r; r++) {
              currentLastRow++;
              for (let c = rng.s.c; c <= rng.e.c; c++) {
                const srcAddr = XLSX.utils.encode_cell({ r, c });
                const dstCol  = parentRange.s.c + (c - rng.s.c);
                const dstAddr = XLSX.utils.encode_cell({ r: currentLastRow, c: dstCol });
                const srcCell = ws[srcAddr];
                if (srcCell !== undefined) { parentWs[dstAddr] = Object.assign({}, srcCell); if (dstCol > maxCol) maxCol = dstCol; }
              }
            }
          }
          parentWs['!ref'] = XLSX.utils.encode_range({ s: parentRange.s, e: { r: currentLastRow, c: maxCol } });
          mSetProgress(90);
          if (fmt === 'xls') { XLSX.writeFile(parentWb, baseName + '.xls', { bookType: 'biff8' }); }
          else               { XLSX.writeFile(parentWb, baseName + '.xlsx'); }
        }

      } else if (mMergeMode === 'flatten') {
        // ── Flatten ALL sheets from ALL files into one single sheet ──────────
        // Helper: copy every data row of a worksheet into parentWs
        function appendSheetRows(ws, parentWs, parentHdrs, currentLastRow, maxCol, parentStartCol) {
          if (!ws['!ref']) return { currentLastRow, maxCol };
          const rng     = XLSX.utils.decode_range(ws['!ref']);
          const srcHdrs = mGetHeaders(ws, rng);
          // Skip header row if it matches the parent headers
          const startR  = JSON.stringify(srcHdrs) === JSON.stringify(parentHdrs) ? rng.s.r + 1 : rng.s.r;
          for (let r = startR; r <= rng.e.r; r++) {
            currentLastRow++;
            for (let c = rng.s.c; c <= rng.e.c; c++) {
              const srcAddr = XLSX.utils.encode_cell({ r, c });
              const dstCol  = parentStartCol + (c - rng.s.c);
              const dstAddr = XLSX.utils.encode_cell({ r: currentLastRow, c: dstCol });
              const srcCell = ws[srcAddr];
              if (srcCell !== undefined) {
                parentWs[dstAddr] = Object.assign({}, srcCell);
                if (dstCol > maxCol) maxCol = dstCol;
              }
            }
          }
          return { currentLastRow, maxCol };
        }

        if (fmt === 'csv') {
          // CSV flatten: plain text rows
          let csvRows      = [];
          let parentHeaders = null;
          for (let fi = 0; fi < mFiles.length; fi++) {
            mSetProgress(10 + fi / mFiles.length * 78);
            const wb = await mReadFile(mFiles[fi]);
            for (const sheetName of wb.SheetNames) {
              const ws = wb.Sheets[sheetName];
              if (!ws['!ref']) continue;
              const range = XLSX.utils.decode_range(ws['!ref']);
              const hdrs  = mGetHeaders(ws, range);
              const startR = (parentHeaders && JSON.stringify(hdrs) === JSON.stringify(parentHeaders))
                ? range.s.r + 1 : range.s.r;
              // Blank separator row between sheets (skip before very first sheet)
              if (parentHeaders) csvRows.push([]);
              for (let r = startR; r <= range.e.r; r++) {
                const row = [];
                for (let c = range.s.c; c <= range.e.c; c++) {
                  const cell = ws[XLSX.utils.encode_cell({ r, c })];
                  row.push(cell ? String(cell.v ?? '') : '');
                }
                csvRows.push(row);
              }
              if (!parentHeaders) parentHeaders = hdrs;
            }
          }
          const csv = csvRows.map(row => row.map(v => {
            const s = String(v);
            return (s.includes(',') || s.includes('"') || s.includes('\n')) ? '"' + s.replace(/"/g, '""') + '"' : s;
          }).join(',')).join('\r\n');
          mDownloadBlob(csv, baseName + '.csv', 'text/csv');

        } else {
          // XLSX / XLS flatten: cell-object copy to preserve all formatting
          const firstWb     = await mReadFile(mFiles[0]);
          mSetProgress(15);
          const parentWs    = firstWb.Sheets[firstWb.SheetNames[0]];
          const parentRange = XLSX.utils.decode_range(parentWs['!ref'] || 'A1');
          const parentHdrs  = mGetHeaders(parentWs, parentRange);
          let currentLastRow = parentRange.e.r;
          let maxCol         = parentRange.e.c;

          // Remaining sheets of the first file
          for (let si = 1; si < firstWb.SheetNames.length; si++) {
            mSetProgress(15 + (si / firstWb.SheetNames.length) * 20);
            const ws = firstWb.Sheets[firstWb.SheetNames[si]];
            currentLastRow++; // blank separator row
            ({ currentLastRow, maxCol } = appendSheetRows(ws, parentWs, parentHdrs, currentLastRow, maxCol, parentRange.s.c));
          }

          // All sheets of remaining files
          for (let fi = 1; fi < mFiles.length; fi++) {
            mSetProgress(35 + fi / mFiles.length * 55);
            const wb = await mReadFile(mFiles[fi]);
            for (const sheetName of wb.SheetNames) {
              const ws = wb.Sheets[sheetName];
              currentLastRow++; // blank separator row
              ({ currentLastRow, maxCol } = appendSheetRows(ws, parentWs, parentHdrs, currentLastRow, maxCol, parentRange.s.c));
            }
          }

          parentWs['!ref'] = XLSX.utils.encode_range({ s: parentRange.s, e: { r: currentLastRow, c: maxCol } });

          // Remove all sheets except the first — data is now all in parentWs
          const keepName = firstWb.SheetNames[0];
          firstWb.SheetNames.slice(1).forEach(n => { delete firstWb.Sheets[n]; });
          firstWb.SheetNames.splice(1);
          // Also remove extra sheets from any other loaded workbooks (they were already copied)
          // The final workbook only has one sheet now
          firstWb.Sheets[keepName] = parentWs;

          mSetProgress(90);
          if (fmt === 'xls') { XLSX.writeFile(firstWb, baseName + '.xls', { bookType: 'biff8' }); }
          else               { XLSX.writeFile(firstWb, baseName + '.xlsx'); }
        }

      } else {
        // Separate sheets
        const parentWb  = await mReadFile(mFiles[0]);
        const firstName = mFiles[0].name.replace(/\.[^.]+$/, '').substring(0, 31);
        const origName  = parentWb.SheetNames[0];
        parentWb.SheetNames[0]       = firstName;
        parentWb.Sheets[firstName]   = parentWb.Sheets[origName];
        delete parentWb.Sheets[origName];
        const usedNames = { [firstName]: true };

        for (let i = 1; i < mFiles.length; i++) {
          mSetProgress(10 + i / mFiles.length * 80);
          const wb  = await mReadFile(mFiles[i]);
          const ws  = wb.Sheets[wb.SheetNames[0]];
          let shName = mFiles[i].name.replace(/\.[^.]+$/, '').substring(0, 31);
          let base = shName, n = 2;
          while (usedNames[shName]) { shName = base.substring(0, 28) + '_' + n++; }
          usedNames[shName] = true;
          XLSX.utils.book_append_sheet(parentWb, ws, shName);
        }
        mSetProgress(90);
        if (fmt === 'csv')      { mDownloadBlob(XLSX.utils.sheet_to_csv(parentWb.Sheets[parentWb.SheetNames[0]]), baseName + '.csv', 'text/csv'); }
        else if (fmt === 'xls') { XLSX.writeFile(parentWb, baseName + '.xls', { bookType: 'biff8' }); }
        else                    { XLSX.writeFile(parentWb, baseName + '.xlsx'); }
      }

      mSetProgress(100);
      const modeLabel = mMergeMode === 'flatten' ? 'flattened' : 'merged';
      mShowStatus(`✅ ${modeLabel.charAt(0).toUpperCase()+modeLabel.slice(1)} ${mFiles.length} file(s) → "${baseName}.${fmt}"`, 'success');
    } catch (err) {
      mShowStatus('❌ Error: ' + err.message, 'error');
    }
    btn.disabled = false;
  };
})();

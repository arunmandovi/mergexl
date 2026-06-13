// ══════════════════════════════════════════════════════════════════════════
//  PDF TO EXCEL CONVERTER
// ══════════════════════════════════════════════════════════════════════════
(function () {
  // Configure PDF.js worker
  if (typeof pdfjsLib !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }

  let pFiles     = [];
  let pSheetMode = 'combined';

  const dropZone  = document.getElementById('p-dropZone');
  const fileInput = document.getElementById('p-fileInput');

  dropZone.addEventListener('click',    () => fileInput.click());
  fileInput.addEventListener('change',  e  => pAddFiles(Array.from(e.target.files)));
  dropZone.addEventListener('dragover', e  => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop',     e  => {
    e.preventDefault(); dropZone.classList.remove('drag-over');
    pAddFiles(Array.from(e.dataTransfer.files).filter(f => /\.pdf$/i.test(f.name)));
  });

  function pAddFiles(nf) {
    nf.forEach(f => { if (!pFiles.find(x => x.name === f.name && x.size === f.size)) pFiles.push(f); });
    pRenderFiles();
  }
  window.pRemoveFile = function (i) { pFiles.splice(i, 1); pRenderFiles(); };

  function pRenderFiles() {
    const container = document.getElementById('p-fileItems');
    const listEl    = document.getElementById('p-fileList');
    document.getElementById('p-fileCount').textContent = pFiles.length;
    if (!pFiles.length) {
      listEl.style.display = 'none';
      document.getElementById('p-convertBtn').disabled = true;
      return;
    }
    listEl.style.display = 'block';
    document.getElementById('p-convertBtn').disabled = false;
    container.innerHTML = pFiles.map((f, i) => `
      <div class="file-item">
        <div class="fi-name"><span>📄</span><span>${escHtml(f.name)}</span></div>
        <div style="display:flex;align-items:center;gap:11px">
          <span class="fi-size">${(f.size / 1024).toFixed(1)} KB</span>
          <button class="fi-rm" onclick="pRemoveFile(${i})">✕</button>
        </div>
      </div>`).join('');
  }

  window.pSelectSheetMode = function (mode) {
    pSheetMode = mode;
    document.getElementById('p-modeCombined').classList.toggle('p-sel', mode === 'combined');
    document.getElementById('p-modeSheets').classList.toggle('p-sel',   mode === 'sheets');
  };

  function pSetProgress(pct, msg) {
    document.getElementById('p-progWrap').style.display = 'block';
    document.getElementById('p-progFill').style.width   = pct + '%';
    document.getElementById('p-progMsg').textContent    = msg || '';
  }
  function pShowStatus(msg, type) {
    const el = document.getElementById('p-status');
    el.textContent = msg; el.className = 'p-status ' + type; el.style.display = 'block';
  }

  // ── Column detection helpers ────────────────────────────────────────────

  // Merge close x-positions into column centres
  function clusterPositions(positions, threshold) {
    if (!positions.length) return [];
    const sorted = [...new Set(positions.map(p => Math.round(p)))].sort((a, b) => a - b);
    const clusters = [];
    let group = [sorted[0]];
    for (let i = 1; i < sorted.length; i++) {
      if (sorted[i] - sorted[i - 1] <= threshold) {
        group.push(sorted[i]);
      } else {
        clusters.push(group.reduce((a, b) => a + b, 0) / group.length);
        group = [sorted[i]];
      }
    }
    clusters.push(group.reduce((a, b) => a + b, 0) / group.length);
    return clusters;
  }

  function findNearestCol(x, colPositions) {
    let best = 0, bestDist = Infinity;
    colPositions.forEach((cx, i) => {
      const d = Math.abs(x - cx);
      if (d < bestDist) { bestDist = d; best = i; }
    });
    return best;
  }

  // ── Core extraction ─────────────────────────────────────────────────────

  function extractTableFromPage(textContent, viewport) {
    const pageHeight = viewport.height;

    // Collect all non-empty text items with page coordinates (y flipped so 0=top)
    const items = [];
    for (const item of textContent.items) {
      if (!item.str || !item.str.trim()) continue;
      items.push({
        text: item.str.trim(),
        x:    item.transform[4],
        y:    pageHeight - item.transform[5]  // flip y: 0 = top of page
      });
    }
    if (!items.length) return [];

    // Group items into rows by y-position with tolerance
    const ROW_TOL = 5;
    const rowMap  = [];
    for (const item of items) {
      let placed = false;
      for (const row of rowMap) {
        if (Math.abs(row.avgY - item.y) <= ROW_TOL) {
          row.items.push(item);
          row.avgY = row.items.reduce((s, it) => s + it.y, 0) / row.items.length;
          placed = true;
          break;
        }
      }
      if (!placed) rowMap.push({ avgY: item.y, items: [item] });
    }

    // Sort rows top-to-bottom, items left-to-right
    rowMap.sort((a, b) => a.avgY - b.avgY);
    rowMap.forEach(row => row.items.sort((a, b) => a.x - b.x));

    // Detect column positions from all x coordinates
    const allX         = rowMap.flatMap(row => row.items.map(i => i.x));
    const colPositions = clusterPositions(allX, 25);

    // Build 2-D array — concatenate text that lands in the same cell
    return rowMap.map(row => {
      const cells = new Array(colPositions.length).fill('');
      row.items.forEach(item => {
        const ci   = findNearestCol(item.x, colPositions);
        cells[ci]  = cells[ci] ? cells[ci] + ' ' + item.text : item.text;
      });
      return cells;
    });
  }

  // ── Main conversion ─────────────────────────────────────────────────────

  window.pConvertFiles = async function () {
    if (!pFiles.length) return;
    if (typeof pdfjsLib === 'undefined') {
      pShowStatus('❌ PDF.js library failed to load. Check your internet connection.', 'error');
      return;
    }

    const btn      = document.getElementById('p-convertBtn');
    btn.disabled   = true;
    document.getElementById('p-status').style.display = 'none';
    const baseName = document.getElementById('p-outputName').value.trim() || 'pdf_to_excel';

    try {
      pSetProgress(5, 'Loading PDFs…');

      // Load all PDF documents first to get total page count
      const pdfDocs = [];
      for (const f of pFiles) {
        const buf = await f.arrayBuffer();
        const doc = await pdfjsLib.getDocument({ data: buf }).promise;
        pdfDocs.push({ file: f, doc });
      }
      const totalPages = pdfDocs.reduce((s, d) => s + d.doc.numPages, 0);

      const wb           = XLSX.utils.book_new();
      const combinedRows = [];
      let processedPages = 0;
      let usedSheetNames = {};

      for (const { file, doc } of pdfDocs) {
        const fileBase = file.name.replace(/\.pdf$/i, '');

        for (let pageNum = 1; pageNum <= doc.numPages; pageNum++) {
          processedPages++;
          pSetProgress(
            5 + (processedPages / totalPages) * 85,
            `Page ${pageNum} / ${doc.numPages}  —  ${file.name}`
          );

          const page        = await doc.getPage(pageNum);
          const viewport    = page.getViewport({ scale: 1 });
          const textContent = await page.getTextContent();
          const table       = extractTableFromPage(textContent, viewport);
          if (!table.length) continue;

          if (pSheetMode === 'sheets') {
            // One sheet per page
            let shName = pFiles.length > 1
              ? `${fileBase.substring(0, 20)}_p${pageNum}`
              : `Page_${pageNum}`;
            shName = shName.substring(0, 31);
            // Make unique
            let base = shName, n = 2;
            while (usedSheetNames[shName]) { shName = base.substring(0, 28) + '_' + n++; }
            usedSheetNames[shName] = true;

            const ws = XLSX.utils.aoa_to_sheet(table);
            autoWidthColumns(ws, table);
            XLSX.utils.book_append_sheet(wb, ws, shName);

          } else {
            // Combined mode — add separator header between pages
            if (combinedRows.length) combinedRows.push([]);
            if (totalPages > 1) {
              const label = pFiles.length > 1
                ? `[ ${file.name}  —  Page ${pageNum} ]`
                : `[ Page ${pageNum} ]`;
              combinedRows.push([label]);
            }
            combinedRows.push(...table);
          }
        }
      }

      if (pSheetMode === 'combined') {
        if (!combinedRows.length) throw new Error('No text content could be extracted from the PDF(s).');
        const ws = XLSX.utils.aoa_to_sheet(combinedRows);
        autoWidthColumns(ws, combinedRows);
        XLSX.utils.book_append_sheet(wb, ws, 'PDF Data');
      }

      if (!wb.SheetNames.length) throw new Error('No content was extracted.');

      pSetProgress(95, 'Writing Excel file…');
      XLSX.writeFile(wb, baseName + '.xlsx');
      pSetProgress(100, '');
      pShowStatus(
        `✅ Converted ${pFiles.length} PDF(s) (${totalPages} page${totalPages > 1 ? 's' : ''}) → "${baseName}.xlsx"`,
        'success'
      );

    } catch (err) {
      pShowStatus('❌ Error: ' + err.message, 'error');
    }

    btn.disabled = false;
  };

  // Set reasonable column widths based on content
  function autoWidthColumns(ws, data) {
    if (!data.length) return;
    const maxCols = Math.max(...data.map(r => r.length));
    const widths  = new Array(maxCols).fill(10);
    data.forEach(row => {
      row.forEach((cell, ci) => {
        const len = String(cell ?? '').length;
        if (len > widths[ci]) widths[ci] = Math.min(len + 2, 50);
      });
    });
    ws['!cols'] = widths.map(w => ({ wch: w }));
  }
})();

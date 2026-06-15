// ── App-level tab switching ────────────────────────────────────────────────
function switchAppTab(tab) {
  document.querySelectorAll('.app-tab').forEach((t, i) =>
    t.classList.toggle('active', (i === 0 && tab === 'merge') || (i === 1 && tab === 'compare')));
  document.getElementById('mergeSection').classList.toggle('active',   tab === 'merge');
  document.getElementById('compareSection').classList.toggle('active', tab === 'compare');
}

// ── Shared utilities ───────────────────────────────────────────────────────
function escHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Read a CSV file as plain text and build a workbook where every cell is an
// exact string — SheetJS never touches the values, so "06-Jun-26" stays "06-Jun-26".
function readCsvAsWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const text = e.target.result;
        // Normalise line endings, drop trailing blank line
        const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
        while (lines.length && lines[lines.length - 1] === '') lines.pop();

        const ws = {};
        let maxC = 0;
        lines.forEach((line, r) => {
          // RFC-4180 field parser — handles quoted fields with commas/newlines/escaped quotes
          const fields = [];
          let field = '', inQ = false;
          for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (inQ) {
              if (ch === '"' && line[i + 1] === '"') { field += '"'; i++; }
              else if (ch === '"') { inQ = false; }
              else { field += ch; }
            } else {
              if (ch === '"') { inQ = true; }
              else if (ch === ',') { fields.push(field); field = ''; }
              else { field += ch; }
            }
          }
          fields.push(field);
          if (fields.length > maxC) maxC = fields.length;
          fields.forEach((val, c) => {
            // t:'s' forces SheetJS to treat this as a plain string — no date/number parsing
            ws[XLSX.utils.encode_cell({ r, c })] = { t: 's', v: val, w: val };
          });
        });

        if (lines.length > 0 && maxC > 0)
          ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lines.length - 1, c: maxC - 1 } });

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        resolve(wb);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsText(file);   // plain text read — no binary interpretation
  });
}

// Reads the bid template out of the .xlsx and writes results back into the same workbook.
// Cells are edited in the sheet XML directly, so formatting, formulas and other sheets are untouched.
(function (root) {
  'use strict';

  const TRAINING_FILL = 'FFFFFF00'; // yellow shading used for training blocks

  function colToNum(letters) { let n = 0; for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
  function numToCol(n) { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - 1 - r) / 26; } return s; }
  function serialToIso(n) { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000); return d.toISOString().slice(0, 10); }
  function xmlEscape(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

  function parseXml(text) { return new DOMParser().parseFromString(text, 'application/xml'); }
  function kids(el, name) { return Array.from(el.childNodes).filter(n => n.nodeType === 1 && n.localName === name); }
  function kid(el, name) { return kids(el, name)[0] || null; }
  function textOf(el) { return Array.from(el.getElementsByTagNameNS('*', 't')).map(t => t.textContent).join(''); }

  function resolveTarget(target) {
    target = target.replace(/^\//, '');
    return target.startsWith('xl/') ? target : 'xl/' + target;
  }

  async function openWorkbook(data) {
    const zip = await JSZip.loadAsync(data, { createFolders: false });
    const wbXml = parseXml(await zip.file('xl/workbook.xml').async('string'));
    const rels = parseXml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
    const relMap = {};
    for (const r of rels.getElementsByTagNameNS('*', 'Relationship')) relMap[r.getAttribute('Id')] = r.getAttribute('Target');
    const sheets = [];
    for (const s of wbXml.getElementsByTagNameNS('*', 'sheet')) {
      const rid = s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || s.getAttribute('r:id');
      sheets.push({ name: s.getAttribute('name'), path: resolveTarget(relMap[rid]) });
    }
    let shared = [];
    if (zip.file('xl/sharedStrings.xml')) {
      const ss = parseXml(await zip.file('xl/sharedStrings.xml').async('string'));
      shared = Array.from(ss.documentElement.childNodes).filter(n => n.localName === 'si').map(textOf);
    }
    const styles = parseXml(await zip.file('xl/styles.xml').async('string'));
    const fills = kids(kid(styles.documentElement, 'fills'), 'fill').map(f => {
      const p = kid(f, 'patternFill'); const fg = p && kid(p, 'fgColor');
      return p && p.getAttribute('patternType') === 'solid' && fg ? (fg.getAttribute('rgb') || '').toUpperCase() : '';
    });
    const xfFill = kids(kid(styles.documentElement, 'cellXfs'), 'xf').map(x => fills[Number(x.getAttribute('fillId') || 0)] || '');
    // The hidden bid sheet is kept apart so it never shows up as a sheet to bid from.
    const stateSheet = sheets.find(sh => sh.name === STATE_SHEET) || null;
    return { zip, sheets: sheets.filter(sh => sh !== stateSheet), stateSheet, shared, xfFill };
  }

  // Load a sheet into rows: Map(rowNum -> Map(colNum -> {v, fill}))
  async function loadGrid(wb, sheetName) {
    const sheet = wb.sheets.find(s => s.name === sheetName);
    const doc = parseXml(await wb.zip.file(sheet.path).async('string'));
    const rows = new Map();
    for (const row of doc.getElementsByTagNameNS('*', 'row')) {
      const r = Number(row.getAttribute('r'));
      const cells = new Map();
      for (const c of kids(row, 'c')) {
        const ref = c.getAttribute('r');
        const col = colToNum(ref.match(/^[A-Z]+/)[0]);
        const t = c.getAttribute('t');
        const vEl = kid(c, 'v');
        let v = vEl ? vEl.textContent : '';
        if (t === 's' && v !== '') v = wb.shared[Number(v)] ?? '';
        else if (t === 'inlineStr') v = textOf(c);
        cells.set(col, { v, fill: wb.xfFill[Number(c.getAttribute('s') || 0)] || '' });
      }
      rows.set(r, cells);
    }
    return rows;
  }

  const cellV = (rows, r, c) => { const row = rows.get(r); const x = row && row.get(c); return x ? x.v : ''; };
  const label = (rows, r) => String(cellV(rows, r, 1)).trim();

  // Work out the template's structure from a sheet laid out like the bid packet.
  async function readTemplate(wb, sheetName) {
    const rows = await loadGrid(wb, sheetName);
    const rowNums = Array.from(rows.keys()).sort((a, b) => a - b);
    const warnings = [];

    // Date header: first row with a long run of consecutive Excel date serials.
    let dateRow = null, firstCol = null, days = [];
    for (const r of rowNums) {
      const cells = Array.from(rows.get(r).entries()).sort((a, b) => a[0] - b[0]);
      for (const [c, x] of cells) {
        const n = Number(x.v);
        if (!(x.v !== '' && n > 30000 && n < 80000)) continue;
        let k = 0; const run = [];
        while (Number(cellV(rows, r, c + k)) === n + k) { run.push(serialToIso(n + k)); k++; }
        if (run.length >= 28) { dateRow = r; firstCol = c; days = run; break; }
      }
      if (dateRow) break;
    }
    if (!dateRow) throw new Error('Could not find a row of dates in this sheet.');

    // The weekday row under the dates should agree with the dates.
    const wdRow = dateRow + 1;
    const names = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    let mismatch = 0;
    days.slice(0, 14).forEach((d, i) => {
      const shown = String(cellV(rows, wdRow, firstCol + i)).trim().slice(0, 3);
      if (shown && shown !== names[new Date(d + 'T12:00:00Z').getUTCDay()]) mismatch++;
    });
    if (mismatch) warnings.push(`The weekday row does not match the dates in row ${dateRow}. Check the dates are for the right year.`);

    // Lines: numbered rows below the header until the Chief Pilot / "Pilots available" rows.
    const lines = [];
    let endRow = null;
    for (const r of rowNums) {
      if (r <= wdRow) continue;
      const a = label(rows, r).toLowerCase();
      if (a.includes('chief pilot') || a.startsWith('pilots available')) { endRow = r; break; }
      const b = Number(cellV(rows, r, 2));
      if (!Number.isInteger(b) || b <= 0) continue;
      const cells = days.map((_, i) => {
        const raw = String(cellV(rows, r, firstCol + i));
        const norm = raw.trim().toUpperCase();
        if (raw !== '' && raw !== norm) warnings.push(`Line ${b}, ${days[i]}: "${raw}" read as "${norm}". The sheet's own counts will miss it until it's fixed.`);
        return norm === 'X' ? 'HAC' : norm; // the 2026 sheet writes HAC days as X
      });
      lines.push({ lineNo: b, row: r, name: label(rows, r), cells });
    }
    if (!lines.length) throw new Error('No numbered lines found under the date header.');

    // Flag off blocks / work stretches that break the line's usual rhythm.
    for (const l of lines) {
      const runs = [];
      let cur = null;
      l.cells.forEach((c, i) => {
        const off = c === 'OFF';
        if (!cur || cur.off !== off) { cur = { off, start: i, len: 0 }; runs.push(cur); }
        cur.len++;
      });
      const inner = runs.slice(1, -1);
      const mode = arr => { const m = {}; arr.forEach(x => m[x] = (m[x] || 0) + 1); return Number(Object.entries(m).sort((a, b) => b[1] - a[1])[0]?.[0]); };
      const offLen = mode(inner.filter(x => x.off).map(x => x.len));
      const onLen = mode(inner.filter(x => !x.off).map(x => x.len));
      for (const x of inner) {
        const want = x.off ? offLen : onLen;
        if (x.len !== want) warnings.push(`Line ${l.lineNo}: ${x.off ? 'off block' : 'work stretch'} starting ${days[x.start]} is ${x.len} days (usually ${want}).`);
      }
    }

    // Holidays: a row after the lines with H marks.
    const after = rowNums.filter(r => r > lines[lines.length - 1].row);
    let holidays = [];
    for (const r of after) {
      const vals = days.map((_, i) => String(cellV(rows, r, firstCol + i)).trim().toUpperCase());
      if (vals.includes('OFF')) continue; // Chief Pilot row also carries H marks
      const h = days.filter((_, i) => vals[i] === 'H');
      if (h.length >= 2) { holidays = h; break; }
    }

    // Events: rows labelled BOARD DATES or MT (any mark under a date counts, as in the sheet formula).
    const events = new Set();
    for (const r of after) {
      const a = label(rows, r).toLowerCase();
      if (a === 'board dates' || a === 'mt') {
        days.forEach((d, i) => { if (String(cellV(rows, r, firstCol + i)).trim() !== '') events.add(d); });
      }
    }

    // Safety / company training days: any date marked "Safety…" or "training" outside the pilots'
    // recurrent TRAINING DATES row (e.g. "Safety Day", "training" under the dates). Those days use
    // the "Pilots Required on Safety & Training" minimum.
    const safety = new Set();
    for (const r of after) {
      if (label(rows, r).toLowerCase().startsWith('training dates')) continue;
      days.forEach((d, i) => { if (/safety|training/i.test(String(cellV(rows, r, firstCol + i)))) safety.add(d); });
    }

    // Staffing minimums written in column B.
    const required = {};
    for (const r of after) {
      const a = label(rows, r).toLowerCase();
      const n = Number(cellV(rows, r, 2));
      if (!a.startsWith('pilots required') || !Number.isFinite(n)) continue;
      if (a.includes('safety')) required.safety = n;
      else if (a.includes('weekday')) required.weekday = n;
      else if (a.includes('weekend')) required.weekend = n;
      else if (a.includes('event')) required.event = n;
      else if (a.includes('holiday')) required.holiday = n;
    }

    // Training: yellow blocks in the TRAINING DATES row, labelled with pilots' initials.
    const training = [];
    const tRow = after.find(r => label(rows, r).toLowerCase().startsWith('training dates'));
    if (tRow) {
      const yellow = days.map((_, i) => { const x = rows.get(tRow).get(firstCol + i); return !!x && x.fill === TRAINING_FILL; });
      const text = days.map((_, i) => String(cellV(rows, tRow, firstCol + i)));
      for (let i = 0; i < days.length; i++) {
        if (!yellow[i] || (i > 0 && yellow[i - 1])) continue;
        let j = i; while (j + 1 < days.length && yellow[j + 1]) j++;
        // A block can hold several back-to-back sessions; each "TRAINING" label starts one.
        let labels = [];
        for (let k = i; k <= j; k++) if (/training/i.test(text[k])) labels.push(k);
        if (!labels.length) labels = [i + 1];
        labels.forEach((lk, n) => {
          const from = Math.max(i, lk - 1);
          // Sessions sharing a block may overlap: one can end on the next session's first
          // training day (e.g. one session 12/9–12/14 and the next 12/13–12/19), but never run past 7 days.
          const next = n + 1 < labels.length ? labels[n + 1] : null;
          const to = next == null ? j : Math.min(next, from + 6);
          const words = text.slice(from, next == null ? j + 1 : next).join(' ').toUpperCase().replace(/TRAINING/g, ' ').split(/[^A-Z]+/).filter(Boolean);
          const three = words.filter(w => w.length === 3);
          const initials = Array.from(new Set(three.length ? three : words.filter(w => w.length >= 2 && w.length <= 4)));
          // First and last shaded days are the positioning days, the ones between are training.
          if (to - from < 2) { warnings.push(`Training block at ${days[from]} is too short to have position days.`); return; }
          training.push({ initials, start: days[from + 1], end: days[to - 1], note: text.slice(from, to + 1).filter(s => s.trim()).join(' / ').replace(/\s+/g, ' ') });
        });
      }
    }

    return { sheetName, dateRow, firstCol, days, lines, holidays, events: Array.from(events).sort(), safety: Array.from(safety).sort(), required, training, warnings };
  }

  // Names already in column A of earlier year sheets, to seed the roster.
  async function readRoster(wb, sheetName) {
    const t = await readTemplate(wb, sheetName);
    return t.lines.filter(l => l.name && !/new hire/i.test(l.name)).map(l => ({ name: l.name, seniority: l.lineNo }));
  }

  // Rewrite the given cells in one row's XML. changes: Map(col -> text|'' )
  function patchRow(rowXml, rowNum, changes) {
    const open = rowXml.match(/^<row\b[^>]*?(\/?)>/);
    if (open[1] === '/') rowXml = rowXml.replace(/\/>$/, '></row>');
    const head = rowXml.match(/^<row\b[^>]*>/)[0];
    const body = rowXml.slice(head.length, -'</row>'.length);
    const cells = [];
    const re = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
    let m;
    while ((m = re.exec(body))) {
      const ref = m[1].match(/\br="([A-Z]+)\d+"/)[1];
      cells.push({ col: colToNum(ref), attrs: m[1], xml: m[0] });
    }
    const style = attrs => { const s = attrs.match(/\bs="(\d+)"/); return s ? ` s="${s[1]}"` : ''; };
    const build = (col, s, text) => {
      const r = `${numToCol(col)}${rowNum}`;
      return text === '' ? `<c r="${r}"${s}/>` : `<c r="${r}"${s} t="inlineStr"><is><t>${xmlEscape(text)}</t></is></c>`;
    };
    for (const [col, text] of changes) {
      const ex = cells.find(c => c.col === col);
      if (ex) ex.xml = build(col, style(ex.attrs), text);
      else {
        const left = cells.filter(c => c.col < col).pop();
        cells.push({ col, attrs: '', xml: build(col, left ? style(left.attrs) : '', text) });
      }
    }
    cells.sort((a, b) => a.col - b.col);
    return head + cells.map(c => c.xml).join('') + '</row>';
  }

  // Hidden sheet that carries a saved bid inside the workbook, so it can be reopened to continue.
  const STATE_SHEET = 'PilotBid data';
  const CHUNK = 30000; // Excel cells hold up to 32,767 characters

  // Produce the workbook: pilot names on chosen lines plus P/T/V/X/H codes. Cells are compared with
  // what's currently in the file (a saved bid may already have picks in it), so undone picks are cleared.
  // opts.stateJson: the bid to embed in the hidden sheet; without it any hidden bid sheet is removed.
  async function writeResults(wb, template, state, opts = {}) {
    const sheet = wb.sheets.find(s => s.name === template.sheetName);
    let xml = await wb.zip.file(sheet.path).async('string');
    const byRow = new Map();
    const hacText = (state.config.codeText && state.config.codeText.HAC) || 'HAC';
    for (const line of state.lines) {
      const inFile = (template.fileLines || template.lines).find(l => l.row === line.row);
      const changes = new Map();
      const pilot = line.pilotId ? state.byId.get(line.pilotId) : null;
      const name = pilot ? pilot.name : '';
      if (name !== (inFile.name || '')) changes.set(1, name);
      line.cells.forEach((c, i) => { if (c !== inFile.cells[i]) changes.set(template.firstCol + i, c === 'HAC' ? hacText : c); });
      if (changes.size) byRow.set(line.row, changes);
    }
    xml = xml.replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, rowXml => {
      const r = Number(rowXml.match(/\br="(\d+)"/)[1]);
      return byRow.has(r) ? patchRow(rowXml, r, byRow.get(r)) : rowXml;
    });
    const out = await JSZip.loadAsync(await wb.zip.generateAsync({ type: "arraybuffer" }), { createFolders: false });
    out.file(sheet.path, xml, { createFolders: false });
    await setStateSheet(out, opts.stateJson || null);
    // Ask Excel to recalculate the count/availability formulas when the file is opened.
    let wbx = await out.file('xl/workbook.xml').async('string');
    if (/<calcPr\b/.test(wbx)) {
      if (!/fullCalcOnLoad=/.test(wbx)) wbx = wbx.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"');
    } else wbx = wbx.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
    out.file('xl/workbook.xml', wbx, { createFolders: false });
    return out.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', compression: 'DEFLATE' });
  }

  // Add, update or (json = null) remove the hidden bid sheet.
  async function setStateSheet(zip, json) {
    let wbx = await zip.file('xl/workbook.xml').async('string');
    let rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    let ct = await zip.file('[Content_Types].xml').async('string');
    const tag = wbx.match(/<sheet\b[^>]*\bname="PilotBid data"[^>]*\/>/);
    let path;
    if (tag) {
      const rid = tag[0].match(/r:id="([^"]+)"/)[1];
      const rel = rels.match(new RegExp(`<Relationship\\b[^>]*\\bId="${rid}"[^>]*/>`))[0];
      path = resolveTarget(rel.match(/Target="([^"]+)"/)[1]);
      if (!json) {
        wbx = wbx.replace(tag[0], '');
        rels = rels.replace(rel, '');
        ct = ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${path}"[^>]*/>`), '');
        zip.remove(path);
      }
    } else {
      if (!json) return;
      const sid = Math.max(0, ...Array.from(wbx.matchAll(/\bsheetId="(\d+)"/g), m => Number(m[1]))) + 1;
      let k = 1; while (zip.file(`xl/worksheets/sheet${k}.xml`)) k++;
      path = `xl/worksheets/sheet${k}.xml`;
      const rid = 'rIdPilotBid';
      wbx = wbx.replace('</sheets>', `<sheet name="${STATE_SHEET}" sheetId="${sid}" state="hidden" r:id="${rid}"/></sheets>`);
      rels = rels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${k}.xml"/></Relationships>`);
      ct = ct.replace('</Types>', `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
    }
    if (json) {
      const rows = [];
      for (let i = 0, r = 1; i < json.length; i += CHUNK, r++) rows.push(`<row r="${r}"><c r="A${r}" t="inlineStr"><is><t>${xmlEscape(json.slice(i, i + CHUNK))}</t></is></c></row>`);
      zip.file(path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.join('')}</sheetData></worksheet>`, { createFolders: false });
    }
    zip.file('xl/workbook.xml', wbx, { createFolders: false });
    zip.file('xl/_rels/workbook.xml.rels', rels, { createFolders: false });
    zip.file('[Content_Types].xml', ct, { createFolders: false });
  }

  // The bid saved inside a workbook, or null if it's a plain spreadsheet.
  async function readState(wb) {
    const sheet = wb.stateSheet;
    if (!sheet) return null;
    const doc = parseXml(await wb.zip.file(sheet.path).async('string'));
    const text = Array.from(doc.getElementsByTagNameNS('*', 'c')).map(textOf).join('');
    return JSON.parse(text);
  }

  root.Xlsx = { STATE_SHEET, openWorkbook, readTemplate, readRoster, readState, writeResults, patchRow, serialToIso, colToNum, numToCol };
})(this);

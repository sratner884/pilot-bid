// UI for the bid: setup -> bidding turns -> download the filled spreadsheet.
(function () {
  'use strict';
  const E = Engine;
  const $ = s => document.querySelector(s);
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const STORE = 'pilotbid.v1';

  const S = { fileName: '', fileB64: '', sheetName: '', wb: null, template: null, config: null, picks: [], state: null, sel: [], viewAs: 'admin', code: 'V', rank: [], rankFor: '', original: null, origWarnings: null };

  // ---------- persistence ----------
  function toB64(buf) {
    const bytes = new Uint8Array(buf); let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }
  function fromB64(b64) { const s = atob(b64); const a = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) a[i] = s.charCodeAt(i); return a.buffer; }
  function snapshot() { return { version: 1, fileName: S.fileName, fileB64: S.fileB64, sheetName: S.sheetName, config: S.config, picks: S.picks, original: S.original, warnings: S.origWarnings }; }
  // The bid as stored in the hidden sheet of a saved Excel file: settings, picks, and the lines as they
  // were before any picks (the visible sheet in that file already has the picks written in).
  function bidState() {
    return { app: 'PilotBid', version: 2, savedAt: new Date().toISOString(), fileName: S.fileName, sheetName: S.sheetName, config: S.config, picks: S.picks,
      original: S.template.lines.map(l => ({ row: l.row, name: l.name, cells: l.cells })), warnings: S.template.warnings };
  }
  function persist() { try { localStorage.setItem(STORE, JSON.stringify(snapshot())); } catch (e) { /* storage full or blocked: bid file still works */ } }
  function download(blob, name) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }

  async function restore(snap) {
    S.fileName = snap.fileName; S.fileB64 = snap.fileB64; S.picks = snap.picks || [];
    S.original = snap.original || null; S.origWarnings = snap.warnings || null;
    S.wb = await Xlsx.openWorkbook(fromB64(snap.fileB64));
    await useSheet(snap.sheetName, snap.config);
  }

  // Continue a bid saved as an Excel file (it carries the bid in a hidden sheet).
  async function openSavedXlsx(buf, name) {
    const wb = await Xlsx.openWorkbook(buf);
    const saved = await Xlsx.readState(wb);
    if (!saved) throw new Error('This workbook has no saved bid in it. To start a new bid from it, load it under Setup → Workbook.');
    S.fileName = saved.fileName || name; S.fileB64 = toB64(buf); S.picks = saved.picks || [];
    S.original = saved.original; S.origWarnings = saved.warnings || [];
    S.wb = wb;
    await useSheet(saved.sheetName, saved.config);
  }

  // ---------- loading ----------
  async function loadFile(buf, name) {
    S.wb = await Xlsx.openWorkbook(buf);
    if (S.wb.stateSheet) { // a saved bid, not a blank template: carry on with it
      alert('This workbook is a saved bid. Opening it so you can continue where it left off.');
      return openSavedXlsx(buf, name);
    }
    S.fileName = name; S.fileB64 = toB64(buf); S.picks = []; S.config = null; S.original = null; S.origWarnings = null;
    await useSheet(S.wb.sheets[S.wb.sheets.length - 1].name, null);
  }

  async function useSheet(name, config) {
    S.sheetName = name;
    S.template = await Xlsx.readTemplate(S.wb, name);
    if (S.original) {
      // A saved bid: the sheet already has picks written in. Keep what's in the file (to know what to
      // change when saving again) and put the lines back as they were before any picks.
      const t = S.template;
      t.fileLines = t.lines.map(l => ({ row: l.row, name: l.name, cells: l.cells.slice() }));
      t.lines = t.lines.map(l => { const o = S.original.find(x => x.row === l.row); return o ? { ...l, name: o.name, cells: o.cells.slice() } : l; });
      if (S.origWarnings) t.warnings = S.origWarnings;
    }
    if (config) {
      S.config = E.normalizeConfig(config);
      // Setups saved before safety/training days existed: take them from the sheet.
      if (!S.config.safetyDays) S.config.safetyDays = (S.template.safety || []).slice();
      if (S.config.limits.safety == null) S.config.limits.safety = S.template.required.safety ?? S.template.lines.length;
    }
    else {
      const keepPilots = S.config ? S.config.pilots : [];
      S.config = E.defaultConfig(S.template);
      S.config.pilots = keepPilots;
    }
    S.sel = [];
    refresh();
  }

  function refresh() {
    if (S.template) S.state = E.buildState(S.template, S.config, S.picks);
    persist();
    renderAll();
  }

  // ---------- rendering ----------
  function renderAll() {
    $('#fileStatus').textContent = S.fileName ? `${S.fileName} · sheet ${S.sheetName}` : 'No spreadsheet loaded';
    const pilots = S.config ? E.sortedPilots(S.config) : [];
    if (S.viewAs !== 'admin' && S.viewAs !== 'turn' && !pilots.some(p => p.id === S.viewAs)) S.viewAs = 'admin';
    $('#viewAs').innerHTML = `<option value="admin">Admin</option><option value="turn" ${S.viewAs === 'turn' ? 'selected' : ''}>Whoever's turn it is</option>` +
      pilots.map(p => `<option value="${p.id}" ${p.id === S.viewAs ? 'selected' : ''}>${esc(p.name)}</option>`).join('');
    const vp = S.state ? viewedPilot(S.state) : null;
    // When nobody is up, say why: nothing loaded, no pilots yet, or every round is done.
    const nobodyUp = !S.state ? 'No one to show yet: load the spreadsheet in Setup.'
      : !S.state.pilots.length ? 'No one to show yet: add pilots in Setup.'
      : 'Bidding is complete: every round is done. Use Start over below to bid again.';
    $('#viewHint').textContent = S.viewAs === 'admin' ? 'Admin view: you can pick for whoever is up, undo, and auto-pick.'
      : vp ? `Showing what ${S.state.byId.get(vp).name} sees.${S.viewAs === 'turn' ? ' Follows the bid as turns change.' : ''}`
      : nobodyUp;
    // A pilot only ever sees the bidding screen; with no pilot to show, keep the admin tabs.
    const asPilot = !!vp;
    document.querySelectorAll('nav button, #saveBid, #openBidLabel').forEach(b => b.classList.toggle('hidden', asPilot && b.dataset.tab !== 'bid'));
    if (asPilot && !$('#tab-setup').classList.contains('hidden') || asPilot && !$('#tab-results').classList.contains('hidden')) showTab('bid');
    renderSetup();
    renderBid();
    renderResults();
  }

  function shiftIso(iso, n) { const d = new Date(iso + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
  function fmt(iso) { const d = new Date(iso + 'T12:00:00Z'); return `${E.dayName(iso)} ${d.getUTCMonth() + 1}/${d.getUTCDate()}`; }

  function renderSetup() {
    const sel = $('#sheetSel'), rs = $('#rosterSheet');
    const opts = S.wb ? S.wb.sheets.map(s => `<option ${s.name === S.sheetName ? 'selected' : ''}>${esc(s.name)}</option>`).join('') : '';
    sel.innerHTML = opts;
    rs.innerHTML = S.wb ? S.wb.sheets.filter(s => s.name !== S.sheetName).map(s => `<option>${esc(s.name)}</option>`).reverse().join('') : '';
    const t = S.template, c = S.config;
    if (!t) { $('#tplSummary').innerHTML = ''; return; }
    const st = S.state;
    const short = E.shortages(E.buildState(t, c, []));
    $('#tplSummary').innerHTML = `
      <div class="chips">
        <span class="chip">${t.lines.length} lines</span>
        <span class="chip">${fmt(t.days[0])} – ${fmt(t.days[t.days.length - 1])}</span>
        <span class="chip">${c.holidays.length} holidays</span>
        <span class="chip">${c.events.length} event days</span>
        <span class="chip">${(c.safetyDays || []).length} safety/training day${(c.safetyDays || []).length === 1 ? '' : 's'}</span>
        <span class="chip">${c.training.length} training sessions</span>
        ${S.picks.length ? `<span class="chip">${S.picks.length} picks made</span>` : ''}
      </div>
      ${t.warnings.length ? `<div class="warn"><b>Check the template</b><ul>${t.warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul></div>` : ''}
      ${short.length ? `<div class="warn" style="margin-top:8px">${short.length} day(s) are already below the minimum before any bids: ${short.slice(0, 12).map(x => fmt(x.d)).join(', ')}${short.length > 12 ? '…' : ''}</div>` : ''}`;

    const L = c.limits;
    $('#limits').innerHTML = [
      ['weekday', 'Weekdays'], ['weekend', 'Weekends'], ['event', 'Event days'], ['holiday', 'Holidays'], ['safety', 'Safety & training days'],
    ].map(([k, lab]) => `<label class="f">${lab}<input type="number" min="0" data-limit="${k}" value="${esc(L[k])}"></label>`).join('');

    $('#allowV').value = c.allowances.V; $('#allowHAC').value = c.allowances.HAC;
    $('#hacText').value = c.codeText?.HAC || 'HAC';

    const pilots = E.sortedPilots(c);
    const lineOf = st ? st.lineOf : new Map();
    $('#pilots').innerHTML = pilots.length ? `<table class="t"><thead><tr><th>Seniority</th><th>Name</th><th>Initials</th><th>V days</th><th>HAC days</th><th>Training</th><th>Line</th><th></th></tr></thead><tbody>
      ${pilots.map((p, i) => `<tr data-pid="${p.id}">
        <td><input type="number" min="1" max="${pilots.length}" data-p="seniority" value="${i + 1}" style="width:60px" title="Type a number to move this pilot to that seniority"></td>
        <td><input type="text" data-p="name" value="${esc(p.name)}"></td>
        <td><input type="text" data-p="initials" value="${esc(p.initials)}" style="width:70px;min-width:0"></td>
        <td><input type="number" min="0" data-p="V" value="${esc(p.allow?.V ?? '')}" placeholder="${c.allowances.V}"></td>
        <td><input type="number" min="0" data-p="HAC" value="${esc(p.allow?.HAC ?? '')}" placeholder="${c.allowances.HAC}"></td>
        <td>${E.trainingFor(c, p).length}</td>
        <td>${lineOf.get(p.id) ? lineOf.get(p.id).lineNo : '–'}</td>
        <td style="white-space:nowrap"><button class="mini" data-pact="up">↑</button> <button class="mini" data-pact="down">↓</button> <button class="mini" data-pact="del">✕</button></td>
      </tr>`).join('')}</tbody></table>
      ${pilots.length > t.lines.length ? `<p class="warn" style="margin-top:8px">${pilots.length} pilots but only ${t.lines.length} lines. The most junior pilots won't get a line.</p>` : ''}`
      : '<p class="hint">No pilots yet. Add them, or copy names from an earlier sheet.</p>';

    const tot = E.roundTotals(c);
    $('#rounds').innerHTML = `<table class="t"><thead><tr><th>Round</th><th>V days</th><th>HAC days</th><th>Holidays per pilot</th><th>Min in a row</th><th>Min working days between groups</th><th>Must touch block</th><th></th></tr></thead><tbody>
      ${c.rounds.map((r, i) => i === 0
        ? `<tr><td>1</td><td colspan="7">Line pick: every pilot picks a line, most senior first. Training goes on each line as it's picked. Nobody picks days off until every pilot has a line.</td></tr>`
        : E.isHolidayRound(r) ? `<tr data-ri="${i}"><td>${i + 1}</td><td class="hint">–</td><td class="hint">–</td>
          <td><input type="number" min="0" data-r="H" value="${r.days.H}"></td>
          <td colspan="3" class="hint" style="margin:0">Holiday ranking: pilots rank holidays, then they're awarded one at a time by seniority, round robin. Set to 0 to make this a V/HAC round.</td>
          <td style="white-space:nowrap"><button class="mini" data-ract="up" ${i === 1 ? 'disabled' : ''}>↑</button> <button class="mini" data-ract="down" ${i === c.rounds.length - 1 ? 'disabled' : ''}>↓</button> <button class="mini" data-ract="del">✕</button></td></tr>`
        : `<tr data-ri="${i}"><td>${i + 1}</td>
          ${E.DAY_CODES.map(code => `<td><input type="number" min="0" data-r="${code}" value="${r.days[code] || 0}"></td>`).join('')}
          <td><input type="number" min="1" data-r="minRun" value="${r.minRun}"></td>
          <td><input type="number" min="0" data-r="minGap" value="${r.minGap ?? 2}"></td>
          <td><input type="checkbox" data-r="mustAdjoin" ${r.mustAdjoin ? 'checked' : ''}></td>
          <td style="white-space:nowrap"><button class="mini" data-ract="up" ${i === 1 ? 'disabled' : ''}>↑</button> <button class="mini" data-ract="down" ${i === c.rounds.length - 1 ? 'disabled' : ''}>↓</button> <button class="mini" data-ract="del">✕</button></td></tr>`).join('')}
      <tr><td><b>Total</b></td><td><b>${tot.V}</b></td><td><b>${tot.HAC}</b></td><td><b>${tot.H}</b></td><td colspan="4"></td></tr></tbody></table>
      ${tot.V > c.allowances.V || tot.HAC > c.allowances.HAC ? `<p class="warn" style="margin-top:8px">The rounds add up to ${tot.V} V and ${tot.HAC} HAC, more than the yearly allowance (${c.allowances.V} V, ${c.allowances.HAC} HAC). Pilots will run out before the last rounds.</p>` : ''}`;

    const pilotByIni = new Map(c.pilots.filter(p => p.initials).map(p => [p.initials.toUpperCase(), p.name]));
    $('#training').innerHTML = c.training.length ? `<table class="t"><thead><tr><th>Initials</th><th>Pilots</th><th>Position out (P)</th><th>Position back (P)</th><th>Training days</th><th>From sheet</th><th></th></tr></thead><tbody>
      ${c.training.slice().sort((a, b) => a.start.localeCompare(b.start)).map(s => `<tr data-tid="${s.id}">
        <td><input type="text" data-tr="initials" value="${esc(s.initials.join(' '))}" style="width:110px;min-width:0"></td>
        <td>${s.initials.map(x => pilotByIni.get(x.toUpperCase()) ? esc(pilotByIni.get(x.toUpperCase())) : `<span style="color:var(--muted)">${esc(x)}?</span>`).join(', ')}</td>
        <td><input type="date" data-tr="start" value="${shiftIso(s.start, -1)}"></td>
        <td><input type="date" data-tr="end" value="${shiftIso(s.end, 1)}"></td>
        <td>${fmt(s.start)} – ${fmt(s.end)}</td>
        <td class="hint" style="margin:0">${esc(s.note)}</td>
        <td><button class="mini" data-tact="del">✕</button></td></tr>`).join('')}</tbody></table>`
      : '<p class="hint">No training sessions found.</p>';
    $('#tmode').value = c.trainingOnOffDay;
    $('#requireApproval').checked = !!c.requireApproval;
    $('#pcode').value = c.positionCode || 'P';
    $('#holidays').value = c.holidays.join('\n');
    $('#events').value = c.events.join(' ');
    $('#safetyDays').value = (c.safetyDays || []).join('\n');
  }

  function legendHtml() {
    return [['--off', 'Line day off'], ['--hol', 'H (holiday off)'], ['--train', 'Training (T)'], ['--pos', 'Position (P)'], ['--vac', 'V'], ['--hac', 'HAC'], ['--ok-bg', 'Can pick'], ['--wk', "Can't pick"], ['--sel', 'Selected'], ['--event', 'Event'], ['--hol', 'Holiday (dot)'], ['--safety', 'Safety/training day (dot)']]
      .map(([v, l]) => `<span style="--sw:var(${v})">${l}</span>`).join('');
  }

  function cellClass(code) { return code ? 'c-' + code.replace(/\s/g, '') : ''; }
  function cellText(code) { return code === 'OFF' ? '' : code === 'HAC' ? (S.config.codeText?.HAC || 'HAC') : code.replace('OFF ', ''); }

  function renderBid() {
    const st = S.state;
    const turnCard = $('#turnCard'), pickCard = $('#pickCard');
    if (!st) { turnCard.innerHTML = '<p class="hint">Load a spreadsheet in Setup first.</p>'; pickCard.innerHTML = ''; $('#grid').innerHTML = ''; $('#log').innerHTML = ''; return; }
    if (!st.pilots.length) { turnCard.innerHTML = '<p class="hint">Add pilots in Setup first.</p>'; pickCard.innerHTML = ''; }
    const t = st.turn;
    const vp = viewedPilot(st);
    if (vp) return renderPilotView(st, vp);
    const undo = `<button class="btn" id="undo" ${S.picks.length ? '' : 'disabled'}>Undo last</button>`;
    const auto = `<div class="row practice"><span class="hint" style="margin:0">Practice:</span>
      <button class="mini" data-auto="turn" ${t ? '' : 'disabled'}>Auto-pick this turn</button>
      <button class="mini" data-auto="round" ${t ? '' : 'disabled'}>Auto-pick rest of round</button>
      <button class="mini" data-auto="all" ${t ? '' : 'disabled'}>Auto-pick whole bid</button>
      <button class="mini" id="startOver" ${S.picks.length ? '' : 'disabled'}>Start over</button></div>`;
    if (st.pilots.length && !t) {
      turnCard.innerHTML = `<div class="turn"><div><div class="who">Bidding complete</div><div class="sub">All ${st.config.rounds.length} rounds are done. Download the spreadsheet from Results.</div></div><span class="spacer"></span>${undo}</div>${auto}`;
      pickCard.innerHTML = '';
    } else if (t && t.awaiting) {
      const rd = st.config.rounds[t.round];
      const p = st.byId.get(t.pilotId);
      turnCard.innerHTML = `<div class="turn">
        <div><div class="sub">Round ${t.round + 1} of ${st.config.rounds.length} · ${esc(E.roundSummary(rd))} · waiting for your approval</div>
        <div class="who">${esc(p.name)} has finished</div>
        <div class="sub">Check their picks below, then approve to move on to the next pilot, or send the turn back so they can change it.</div></div>
        <span class="spacer"></span><button class="btn" id="sendBack">Send back to pilot</button>
        <button class="btn primary" id="approveTurn">Approve</button></div>${auto}`;
      pickCard.innerHTML = `<h2>${esc(p.name)} · round ${t.round + 1}</h2>${E.isHolidayRound(rd) ? rankingsHtml(st, p.id) : `<div class="hchips">${roundCell(st, t.round, p)}</div>`}`;
    } else if (t) {
      const rd = st.config.rounds[t.round];
      const p = st.byId.get(t.pilotId);
      const order = st.pilots.findIndex(x => x.id === p.id) + 1;
      const extra = turnText(st, t, p);
      turnCard.innerHTML = `<div class="turn">
        <div><div class="sub">Round ${t.round + 1} of ${st.config.rounds.length} · ${esc(E.roundSummary(rd))} · seniority ${order}</div><div class="who">${esc(p.name)}</div><div class="sub">${esc(extra)}</div></div>
        <span class="spacer"></span>${undo}
        <button class="btn" id="endTurn">${E.needsLine(st, t.round, p.id) ? 'Skip pilot' : E.isHolidayRound(rd) ? 'Skip (no holidays)' : 'Finish turn'}</button></div>${auto}`;
      if (E.needsLine(st, t.round, p.id)) renderLinePicker(st, p);
      else if (E.isHolidayRound(rd)) renderRankPicker(st, p, rd);
      else renderOffPicker(st, p, rd);
    }
    renderBoard(st);
  }

  // The pilot whose screen is shown, or null for the admin view.
  function viewedPilot(st) {
    if (S.viewAs === 'admin') return null;
    if (S.viewAs === 'turn') return st.turn ? st.turn.pilotId : null;
    return st.byId.has(S.viewAs) ? S.viewAs : null;
  }

  // "Pick your line, then 7 V · 2 HAC this round (19 V, 4 HAC left for the year) · min 1 in a row"
  function turnText(st, t, p) {
    const rd = st.config.rounds[t.round];
    if (E.isHolidayRound(rd)) return `Rank the holidays you want. After everyone has ranked, up to ${rd.days.H} each are awarded one at a time by seniority.`;
    const left = E.DAY_CODES.map(c => [c, E.remainingFor(st, t.round, p.id, c)]).filter(([, n]) => n > 0);
    const parts = left.map(([c, n]) => `${n} ${E.codeName(c)}`).join(' · ');
    const year = ['V', 'HAC'].filter(c => rd.days[c] > 0).map(c => `${E.allowanceFor(st.config, p, c) - st.used.get(p.id)[c]} ${c}`).join(', ');
    const days = parts ? `${parts} left this round${year ? ` (${year} left for the year)` : ''}${rd.minRun > 1 ? ` · at least ${rd.minRun} days in a row` : ''}${rd.minGap > 1 ? ` · at least ${rd.minGap} working days between groups of days off` : ''}${rd.mustAdjoin ? ' · must touch a block' : ''}` : 'No days left this round';
    return E.needsLine(st, t.round, p.id) ? `Pick a line first${parts ? `, then ${parts}` : ''}.` : days;
  }

  // The code the day picker is adding, kept to one this pilot can still use.
  function currentCode(st) {
    const open = st.turn ? E.openCodes(st, st.turn.round, st.turn.pilotId) : [];
    if (!open.includes(S.code)) S.code = open[0] || 'V';
    return S.code;
  }

  // What a pilot sees: their turn (same pickers as admin), or a waiting screen with their own line.
  function renderPilotView(st, pilotId) {
    const me = st.byId.get(pilotId), t = st.turn;
    const banner = `<div class="sub">Practice · viewing as ${esc(me.name)}</div>`;
    if (t && t.pilotId === me.id && !t.awaiting) {
      const rd = st.config.rounds[t.round];
      const left = turnText(st, t, me);
      $('#turnCard').innerHTML = `<div class="turn"><div>${banner}<div class="who">It's your turn</div>
        <div class="sub">Round ${t.round + 1} of ${st.config.rounds.length} · ${esc(left)}</div></div>
        <span class="spacer"></span>${E.needsLine(st, t.round, me.id) ? '' : `<button class="btn" id="endTurn">${E.isHolidayRound(rd) ? "I don't want any holidays" : 'Finish turn'}</button>`}</div>`;
      if (E.needsLine(st, t.round, me.id)) renderLinePicker(st, me);
      else if (E.isHolidayRound(rd)) renderRankPicker(st, me, rd);
      else renderOffPicker(st, me, rd);
    } else {
      let msg;
      if (!t) msg = 'Bidding is complete.';
      else if (t.awaiting && t.pilotId === me.id) msg = "You've finished your turn. It's waiting for the admin's approval; if they send it back, you'll be able to change it.";
      else if (t.awaiting) msg = `${esc(st.byId.get(t.pilotId).name)} has finished and is waiting for the admin's approval.`;
      else {
        const who = st.byId.get(t.pilotId).name;
        const order = st.pilots.map(p => p.id);
        const ahead = order.indexOf(me.id) - order.indexOf(t.pilotId);
        const later = ahead > 0 ? `${ahead - 1 === 0 ? "You're next" : `${ahead - 1} pilot${ahead - 1 === 1 ? '' : 's'} ahead of you`} this round.`
          : t.round + 1 < st.config.rounds.length ? `Your next turn is in round ${t.round + 2}.` : 'You have no more turns.';
        msg = `${esc(who)} is picking (${esc(E.roundLabel(st.config, t.round))}). ${later}`;
        const mine = st.ranks[t.round] && st.ranks[t.round].get(me.id);
        if (E.isHolidayRound(st.config.rounds[t.round]) && mine) msg += ` You've ranked ${mine.length} holiday${mine.length === 1 ? '' : 's'}; they're awarded once everyone has ranked.`;
      }
      const won = Object.values(st.awards).flat().filter(a => a.pilotId === me.id);
      if (Object.keys(st.awards).length) msg += won.length ? ` Holidays awarded to you: ${won.map(a => esc(fmt(st.days[a.dayIdx]))).join(', ')}.` : ' No holidays were awarded to you.';
      $('#turnCard').innerHTML = `<div class="turn"><div>${banner}<div class="who">${esc(me.name)}</div><div class="sub">${msg}</div></div></div>`;
      if (st.lineOf.get(me.id)) renderOffPicker(st, me, null);
      else $('#pickCard').innerHTML = `<h2>Your line</h2><p class="hint">You haven't picked a line yet.</p>`;
    }
    renderBoard(st);
  }

  // Each pilot's holiday ranking in order, marked with what they were awarded once awards are done.
  // onlyId limits it to one pilot (a pilot sees only their own ranking).
  function rankingsHtml(st, onlyId) {
    const rounds = st.config.rounds.map((rd, r) => [rd, r]).filter(([rd]) => E.isHolidayRound(rd));
    if (!rounds.length) return '';
    return rounds.map(([rd, r]) => {
      const ranks = st.ranks[r] || new Map(), awards = st.awards[r];
      const won = new Map((awards || []).map(a => [`${a.pilotId}:${a.dayIdx}`, a]));
      const rows = st.pilots.filter(p => !onlyId || p.id === onlyId).map(p => {
        const list = ranks.get(p.id);
        let cells;
        if (list && list.length) cells = list.map((i, k) => {
          const a = won.get(`${p.id}:${i}`);
          const cls = awards ? (a ? 'won' : 'lost') : '';
          const name = holidayName(st.days[i]);
          return `<span class="hchip ${cls}" title="${a ? `Awarded in pass ${a.pass}` : awards ? 'Not awarded' : 'Waiting for awards'}">${k + 1}. ${esc(fmt(st.days[i]))}${name ? ` · ${esc(name)}` : ''}${a ? ' ✓' : ''}</span>`;
        }).join('');
        else if (list || st.finished.has(`${r}:${p.id}`)) cells = '<span class="hint" style="margin:0">Doesn\'t want any holidays</span>';
        else cells = `<span class="hint" style="margin:0">${st.lineOf.has(p.id) ? 'Not ranked yet' : 'No line'}</span>`;
        const got = awards ? awards.filter(a => a.pilotId === p.id).length : null;
        return `<tr><td style="white-space:nowrap">${esc(p.name)}</td><td><div class="hchips">${cells}</div></td>${awards ? `<td>${got}</td>` : ''}</tr>`;
      }).join('');
      return `${rounds.length > 1 ? `<h4 style="margin:8px 0 4px">Round ${r + 1}</h4>` : ''}
        <p class="hint" style="margin:0 0 8px">${awards ? 'Green ✓ = awarded; crossed out = not awarded (taken by more senior pilots or the day hit its staffing minimum).' : `Rankings so far. Holidays are awarded once every pilot has ranked (up to ${rd.days.H} each).`}</p>
        <table class="t"><thead><tr><th>Pilot</th><th>Ranking (1 = most wanted)</th>${awards ? '<th>Awarded</th>' : ''}</tr></thead><tbody>${rows}</tbody></table>`;
    }).join('');
  }

  // What a pilot has in one round: their line, awarded holidays, or V/HAC picks, else a status.
  function roundCell(st, r, p) {
    const rd = st.config.rounds[r];
    const chip = (code, text) => `<span class="hchip c-${code}">${code === 'HAC' ? esc(S.config.codeText?.HAC || 'HAC') : code} ${esc(text)}</span>`;
    const order = st.pilots.indexOf(p), t = st.turn;
    const status = () => {
      if (t && t.round === r && t.pilotId === p.id) return `<span class="hint">${t.awaiting ? 'Nothing taken (awaiting approval)' : 'Picking now'}</span>`;
      if (t && (r > t.round || (r === t.round && order > st.pilots.findIndex(x => x.id === t.pilotId)))) return '<span class="hint">Not yet</span>';
      return '<span class="hint">None</span>';
    };
    if (rd.line) {
      const l = st.log.find(pk => pk.type === 'line' && pk.pilotId === p.id);
      return l ? `<b>Line ${l.lineNo}</b>` : status();
    }
    if (E.isHolidayRound(rd)) {
      const aw = st.awards[r];
      if (aw) {
        const mine = aw.filter(a => a.pilotId === p.id);
        return mine.length ? mine.map(a => chip('H', fmt(st.days[a.dayIdx]))).join('') : '<span class="hint">None awarded</span>';
      }
      const rk = st.ranks[r] && st.ranks[r].get(p.id);
      return rk ? `<span class="hint">Ranked ${rk.length}, awaiting awards</span>` : status();
    }
    const picks = st.log.filter(pk => pk.type === 'off' && pk.round === r && pk.pilotId === p.id);
    if (!picks.length) return status();
    const span = pk => pk.dayIdx.length === 1 ? fmt(st.days[pk.dayIdx[0]]) : `${fmt(st.days[pk.dayIdx[0]])}–${fmt(st.days[pk.dayIdx[pk.dayIdx.length - 1]])}`;
    const live = t && t.round === r && t.pilotId === p.id;
    return picks.map(pk => chip(pk.code, `${span(pk)} (${pk.dayIdx.length})`)).join('') + (live ? ` <span class="hint">${t.awaiting ? 'awaiting approval' : 'in progress'}</span>` : '');
  }

  // Admin: every pilot (rows) by round (columns), with totals.
  function picksGridHtml(st) {
    if (!st.pilots.length) return '<p class="hint">No pilots yet.</p>';
    const rounds = st.config.rounds;
    return `<div class="scroll"><table class="t picks"><thead><tr><th>Pilot</th>
      ${rounds.map((rd, r) => `<th>Round ${r + 1}<br><span class="hint" style="font-weight:400">${esc(E.roundSummary(rd))}</span></th>`).join('')}
      <th>Totals</th></tr></thead><tbody>
      ${st.pilots.map(p => `<tr class="${st.turn && st.turn.pilotId === p.id ? 'me' : ''}"><td style="white-space:nowrap">${esc(p.name)}</td>
        ${rounds.map((rd, r) => `<td><div class="hchips">${roundCell(st, r, p)}</div></td>`).join('')}
        <td style="white-space:nowrap">V ${st.used.get(p.id).V}/${E.allowanceFor(st.config, p, 'V')}<br>${esc(S.config.codeText?.HAC || 'HAC')} ${st.used.get(p.id).HAC}/${E.allowanceFor(st.config, p, 'HAC')}<br>H ${st.used.get(p.id).H}</td></tr>`).join('')}
      </tbody></table></div>`;
  }

  // Pilot: their own results round by round.
  function myBidHtml(st, p) {
    return `<table class="t"><thead><tr><th>Round</th><th>What you have</th></tr></thead><tbody>
      ${st.config.rounds.map((rd, r) => `<tr><td style="white-space:nowrap">${r + 1} · ${esc(E.roundSummary(rd))}</td><td><div class="hchips">${roundCell(st, r, p)}</div></td></tr>`).join('')}
      </tbody></table>
      <p class="hint" style="margin:8px 0 0">So far: ${st.used.get(p.id).V} of ${E.allowanceFor(st.config, p, 'V')} V, ${st.used.get(p.id).HAC} of ${E.allowanceFor(st.config, p, 'HAC')} HAC (${esc(S.config.codeText?.HAC || 'HAC')}), ${st.used.get(p.id).H} holiday${st.used.get(p.id).H === 1 ? '' : 's'}.</p>`;
  }

  function renderBoard(st) {
    const viewing = viewedPilot(st);
    $('#summaryCard').innerHTML = viewing
      ? `<h2>Your bid so far</h2>${myBidHtml(st, st.byId.get(viewing))}`
      : `<h2>Picks by pilot</h2><p class="hint">Everything each pilot has picked or been awarded, round by round. The highlighted row is the pilot who's up.</p>${picksGridHtml(st)}`;
    // Holiday rankings: shown once the bid reaches a holiday round; a pilot sees only their own.
    const firstHol = st.config.rounds.findIndex(E.isHolidayRound);
    const reached = firstHol >= 0 && (!st.turn || st.turn.round >= firstHol);
    const vp = viewedPilot(st);
    const t = st.turn;
    const rankingNow = vp && t && !t.awaiting && t.pilotId === vp && E.isHolidayRound(st.config.rounds[t.round]); // the editor above already shows it
    $('#rankCard').classList.toggle('hidden', !reached || rankingNow);
    if (reached) $('#rankCard').innerHTML = `<h2>${vp ? 'Your holiday ranking' : 'Holiday rankings'}</h2>${rankingsHtml(st, vp)}`;
    $('#legend2').innerHTML = legendHtml();
    renderGrid(st);
    $('#log').innerHTML = st.log.length ? st.log.slice().reverse().map(pk => {
      if (pk.type === 'award') {
        const per = st.pilots.map(p => [p, pk.awards.filter(a => a.pilotId === p.id)]).filter(([, a]) => a.length);
        return `<div><b>R${pk.round + 1} · Holidays awarded</b>${per.length ? per.map(([p, a]) => `<br>${esc(p.name)}: ${a.map(x => `${fmt(st.days[x.dayIdx])} (choice ${x.rank})`).join(', ')}`).join('') : '<br>None could be awarded.'}</div>`;
      }
      const name = esc(st.byId.get(pk.pilotId)?.name);
      if (pk.type === 'approve') return `<div>R${pk.round + 1} · Admin approved ${name}'s turn</div>`;
      if (pk.type === 'rank') return `<div>R${pk.round + 1} · ${name} submitted a holiday ranking (${pk.dayIdx.length} holiday${pk.dayIdx.length === 1 ? '' : 's'})</div>`;
      if (pk.type === 'line') return `<div>R${pk.round + 1} · ${name} picked line ${pk.lineNo}</div>`;
      if (pk.type === 'off' && pk.code === 'H') return `<div>R${pk.round + 1} · ${name}: holiday ${pk.dayIdx.map(i => fmt(st.days[i])).join(', ')}</div>`;
      if (pk.type === 'off') return `<div>R${pk.round + 1} · ${name}: ${pk.code} ${fmt(st.days[pk.dayIdx[0]])} – ${fmt(st.days[pk.dayIdx[pk.dayIdx.length - 1]])} (${pk.dayIdx.length} days)</div>`;
      return `<div>R${pk.round + 1} · ${name} finished the turn</div>`;
    }).join('') : '<p class="hint">No picks yet.</p>';
  }

  function renderLinePicker(st, pilot) {
    $('#pickCard').innerHTML = `<h2>Available lines</h2><p class="hint">"Training clashes" counts this pilot's training or position days that would fall on the line's days off.</p><div class="lines">
      ${st.lines.map(l => {
        const info = E.lineInfo(st, l, pilot);
        const taken = l.pilotId ? st.byId.get(l.pilotId).name : null;
        return `<div class="linecard ${taken ? 'taken' : ''}"><b>Line ${l.lineNo}</b>
          <span>Off ${esc(info.pattern)} · ${info.offDays} days off${info.holidays ? ` · ${info.holidays} holiday${info.holidays === 1 ? '' : 's'} off` : ''}</span>
          <span class="hint" style="margin:0">First off ${info.firstOff ? fmt(info.firstOff) : '–'}</span>
          ${taken ? `<span>Taken by ${esc(taken)}</span>` : `<span style="color:${info.clash ? 'var(--danger)' : 'var(--muted)'}">${info.clash} training clash${info.clash === 1 ? '' : 'es'}</span>
          <button class="btn primary" data-line="${l.lineNo}">Pick line ${l.lineNo}</button>`}</div>`;
      }).join('')}</div>`;
  }

  // rd = null shows the pilot's line read-only (waiting screen).
  function renderOffPicker(st, pilot, rd) {
    const line = st.lineOf.get(pilot.id);
    const code0 = rd ? currentCode(st) : null;
    const avail = E.bidAvailability(st);
    const selSet = new Set(S.sel);
    const months = [];
    st.days.forEach((iso, i) => {
      const m = iso.slice(0, 7);
      if (!months.length || months[months.length - 1].key !== m) months.push({ key: m, cells: [] });
      months[months.length - 1].cells.push(i);
    });
    const html = months.map(mo => {
      const first = st.days[mo.cells[0]];
      const lead = E.dow(first);
      const [y, mm] = mo.key.split('-').map(Number);
      return `<div class="month"><h4>${MONTHS[mm - 1]} ${y}</h4><div class="mgrid">
        ${['S', 'M', 'T', 'W', 'T', 'F', 'S'].map(d => `<div class="wd">${d}</div>`).join('')}
        ${'<div></div>'.repeat(lead)}
        ${mo.cells.map(i => {
          const code = line.cells[i];
          const iso = st.days[i];
          const dots = (st.sets.holidays.has(iso) ? '<i class="hd"></i>' : '') + (st.sets.events.has(iso) ? '<i class="ev"></i>' : '') + (st.sets.safety.has(iso) ? '<i class="sd"></i>' : '');
          const n = Number(iso.slice(8));
          if (code) return `<div class="d ${cellClass(code)}" title="${esc(fmt(iso))}: ${esc(code)}">${n}${dots}</div>`;
          if (!rd) return `<div class="d plain" title="${esc(fmt(iso))}: working">${n}${dots}</div>`;
          const why = E.dayBlockReason(st, pilot.id, i, avail, code0);
          const cls = selSet.has(i) ? 'sel' : why ? 'blocked' : 'work';
          return `<div class="d ${cls}" data-day="${i}" title="${esc(fmt(iso))}: ${esc(why || `${avail[i]} available, ${st.required[i]} required`)}">${n}${dots}</div>`;
        }).join('')}</div></div>`;
    }).join('');
    if (!rd) {
      $('#pickCard').innerHTML = `<h2>Your line · Line ${line.lineNo}</h2><div class="legend">${legendHtml()}</div><div class="cal">${html}</div>`;
      return;
    }
    const errs = S.sel.length ? E.validateOffPick(st, S.sel, code0) : [];
    const open = E.openCodes(st, st.turn.round, pilot.id);
    // Days taken this turn, each removable until the pilot finishes the turn.
    const mine = S.picks.map((pk, k) => [pk, k]).filter(([pk]) => pk.type === 'off' && pk.round === st.turn.round && pk.pilotId === pilot.id && st.log.includes(pk));
    const span = pk => pk.dayIdx.length === 1 ? fmt(st.days[pk.dayIdx[0]]) : `${fmt(st.days[pk.dayIdx[0]])} – ${fmt(st.days[pk.dayIdx[pk.dayIdx.length - 1]])}`;
    const turnList = `<div class="turnlist"><h4>Your days this turn</h4>
      ${mine.length ? `<ul>${mine.map(([pk, k]) => `<li><span><b>${pk.code}</b> ${esc(span(pk))} · ${pk.dayIdx.length} day${pk.dayIdx.length === 1 ? '' : 's'}</span><button class="btn" data-unpick="${k}">✕ Remove</button></li>`).join('')}</ul>` : '<p class="hint" style="margin:0">None yet.</p>'}
      <p class="hint" style="margin:6px 0 0">You can remove days until you click <b>Finish turn</b>.</p></div>`;
    const codeBtns = open.length > 1 ? `<div class="row" style="gap:6px;margin:4px 0 10px">Adding: ${open.map(c => `<button class="btn ${c === code0 ? 'primary' : ''}" data-code="${c}">${c === 'H' ? 'Holidays' : c} (${E.remainingFor(st, st.turn.round, pilot.id, c)} left)</button>`).join('')}</div>` : '';
    const selText = code0 === 'H'
      ? (S.sel.length ? S.sel.map(i => fmt(st.days[i])).join(', ') : 'Click the holidays to bid for.')
      : S.sel.length ? `${fmt(st.days[S.sel[0]])} – ${fmt(st.days[S.sel[S.sel.length - 1]])} · ${S.sel.length} day(s)` : 'Click a first day, then a last day.';
    if (!open.length) {
      $('#pickCard').innerHTML = `<h2>Line ${line.lineNo}</h2>${turnList}
        <p class="okbox" style="margin-top:12px">You've used all your days for this round. Remove some to change them, or click <b>Finish turn</b>.</p>`;
      return;
    }
    $('#pickCard').innerHTML = `<h2>Line ${line.lineNo} · ${code0 === 'H' ? 'pick holidays' : `pick ${code0} days`}</h2>
      ${turnList}
      ${codeBtns}
      <div class="legend">${legendHtml()}</div>
      <div class="cal">${html}</div>
      <div class="selbox"><div class="row" style="align-items:center">
        <b>${esc(selText)}</b><span class="spacer"></span>
        <button class="btn" id="clearSel" ${S.sel.length ? '' : 'disabled'}>Clear</button>
        <button class="btn primary" id="addDays" ${S.sel.length && !errs.length ? '' : 'disabled'}>${code0 === 'H' ? `Award ${S.sel.length || ''} holiday(s)` : `Add ${S.sel.length || ''} ${code0} day(s)`}</button></div>
        ${errs.length ? `<div class="warn" style="margin-top:8px"><ul>${errs.slice(0, 6).map(e => `<li>${esc(e)}</li>`).join('')}</ul></div>` : ''}
      </div>`;
  }

  // US holiday names for the ranking screen; dates not recognised just show the date.
  function holidayName(iso) {
    const d = new Date(iso + 'T12:00:00Z'), m = d.getUTCMonth() + 1, day = d.getUTCDate(), w = d.getUTCDay();
    const nth = Math.ceil(day / 7), last = day + 7 > new Date(Date.UTC(d.getUTCFullYear(), m, 0)).getUTCDate();
    if (m === 1 && day === 1) return "New Year's Day";
    if (m === 1 && w === 1 && nth === 3) return 'Martin Luther King Jr. Day';
    if (m === 2 && w === 1 && nth === 3) return "Presidents' Day";
    if (m === 5 && w === 1 && last) return 'Memorial Day';
    if (m === 6 && day === 19) return 'Juneteenth';
    if (m === 7 && day === 4) return 'Independence Day';
    if (m === 9 && w === 1 && nth === 1) return 'Labor Day';
    if (m === 11 && w === 4 && nth === 4) return 'Thanksgiving';
    if (m === 11 && w === 5 && day >= 23 && day <= 29) return 'Day after Thanksgiving';
    if (m === 12 && day === 24) return 'Christmas Eve';
    if (m === 12 && day === 25) return 'Christmas Day';
    if (m === 12 && day === 31) return "New Year's Eve";
    return '';
  }

  // Holiday ranking for the pilot whose turn it is: build an ordered list, then submit it.
  function renderRankPicker(st, pilot, rd) {
    $('#pickCard').innerHTML = `<h2>Rank your holidays</h2>
      <p class="hint">Click holidays in the order you want them; 1 is the one you want most. Use ↑ ↓ to change the order. Once every pilot has ranked, holidays are handed out one at a time: the most senior pilot gets their top choice, then the next pilot, and so on round robin, up to ${rd.days.H} each. A holiday is skipped once it would leave fewer than ${st.config.limits.holiday} pilots available, so you may not get every one you rank. <b>Check the order before you submit: once submitted, your ranking can't be changed.</b></p>
      ${rankEditorHtml(st, pilot)}`;
  }

  // Ranking editor for the pilot whose turn it is. The ranking is final once submitted.
  function rankEditorHtml(st, pilot) {
    const key = `${st.turn.round}:${pilot.id}`;
    if (S.rankFor !== key) { S.rankFor = key; S.rankPilot = pilot.id; S.rank = []; }
    const eligible = E.eligibleHolidays(st, pilot.id);
    const line = st.lineOf.get(pilot.id);
    const label = i => `${fmt(st.days[i])}${holidayName(st.days[i]) ? ` · ${holidayName(st.days[i])}` : ''}`;
    const notRanked = eligible.filter(i => !S.rank.includes(i));
    const other = st.days.map((d, i) => i).filter(i => st.sets.holidays.has(st.days[i]) && !eligible.includes(i));
    return `<div class="rank">
        <div><h4>Your ranking</h4>${S.rank.length ? `<ol>${S.rank.map((i, k) => `<li><span><b>${k + 1}.</b> ${esc(label(i))}</span>
          <span class="mv"><button class="btn" data-rmove="up" data-ri="${k}" ${k ? '' : 'disabled'} title="Move up">↑</button><button class="btn" data-rmove="down" data-ri="${k}" ${k < S.rank.length - 1 ? '' : 'disabled'} title="Move down">↓</button><button class="btn" data-rmove="del" data-ri="${k}" title="Remove">✕</button></span></li>`).join('')}</ol>`
          : '<p class="hint">Nothing ranked yet.</p>'}</div>
        <div><h4>Holidays</h4>
          ${notRanked.map(i => `<button class="btn rankadd" data-radd="${i}">+ ${esc(label(i))}</button>`).join('') || '<p class="hint">All ranked.</p>'}
          ${other.map(i => `<div class="hint" style="margin:4px 0">${esc(label(i))}: ${/T|P/.test(line.cells[i]) ? 'in training' : 'already off on your line'}</div>`).join('')}
        </div>
      </div>
      <div class="row" style="margin-top:12px"><span class="spacer"></span>
        <button class="btn" id="rankAll" ${notRanked.length ? '' : 'disabled'}>Add the rest in date order</button>
        <button class="btn primary" id="submitRank" ${S.rank.length ? '' : 'disabled'}>Submit ranking (${S.rank.length})</button></div>`;
  }

  function renderGrid(st) {
    const avail = E.availability(st);
    const me = st.turn ? st.turn.pilotId : null;
    const monthHead = [];
    st.days.forEach((iso, i) => {
      const m = Number(iso.slice(5, 7));
      if (!monthHead.length || monthHead[monthHead.length - 1].m !== m) monthHead.push({ m, n: 0 });
      monthHead[monthHead.length - 1].n++;
    });
    const we = st.days.map(d => E.isWeekend(d));
    const rows = st.lines.map(l => {
      const name = l.pilotId ? st.byId.get(l.pilotId).name : '';
      return `<tr class="${l.pilotId && l.pilotId === me ? 'me' : ''}"><td class="lab">${l.lineNo} · ${esc(name || '(open)')}</td>
        ${l.cells.map((c, i) => `<td class="${c ? cellClass(c) : we[i] ? 'we' : ''}">${esc(cellText(c))}</td>`).join('')}</tr>`;
    }).join('');
    $('#grid').innerHTML = `<table class="grid">
      <tr class="mh"><th class="lab"></th>${monthHead.map(h => `<th colspan="${h.n}">${MONTHS[h.m - 1].slice(0, 3)}</th>`).join('')}</tr>
      <tr><th class="lab">Day</th>${st.days.map((d, i) => `<th class="${we[i] ? 'we' : ''}" title="${fmt(d)}">${Number(d.slice(8))}</th>`).join('')}</tr>
      ${rows}
      <tr><td class="lab">Available</td>${avail.map((a, i) => `<td class="${a < st.required[i] ? 'short' : ''}" title="${fmt(st.days[i])}: ${a} available, ${st.required[i]} required">${a}</td>`).join('')}</tr>
      <tr><td class="lab">Required</td>${st.required.map(r => `<td>${r}</td>`).join('')}</tr>
    </table>`;
  }

  function renderResults() {
    const st = S.state;
    if (!st) { $('#resultStatus').innerHTML = ''; $('#usage').innerHTML = ''; $('#short').innerHTML = ''; $('#awards').innerHTML = ''; $('#resultGrid').innerHTML = ''; $('#download').disabled = true; return; }
    $('#download').disabled = false;
    $('#resultStatus').innerHTML = st.turn
      ? `<div class="warn">Bidding isn't finished yet (${esc(E.roundLabel(st.config, st.turn.round))}, ${esc(st.byId.get(st.turn.pilotId).name)} is up). You can still download what's been picked so far.</div>`
      : st.pilots.length ? '<div class="okbox">All rounds complete.</div>' : '';
    $('#usage').innerHTML = st.pilots.length ? `<table class="t"><thead><tr><th>Pilot</th><th>Line</th><th>V used / allowed</th><th>HAC used / allowed</th><th>Holidays awarded</th></tr></thead><tbody>
      ${st.pilots.map(p => `<tr><td>${esc(p.name)}</td><td>${st.lineOf.get(p.id)?.lineNo ?? '–'}</td>
        <td>${st.used.get(p.id).V} / ${E.allowanceFor(st.config, p, 'V')}</td><td>${st.used.get(p.id).HAC} / ${E.allowanceFor(st.config, p, 'HAC')}</td><td>${st.used.get(p.id).H}</td></tr>`).join('')}</tbody></table>` : '';
    $('#awards').innerHTML = rankingsHtml(st, null) || '<p class="hint">There is no holiday round in this bid.</p>';
    $('#resultGrid').innerHTML = picksGridHtml(st);
    const sh = E.shortages(st);
    $('#short').innerHTML = sh.length ? `<table class="t"><thead><tr><th>Date</th><th>Available</th><th>Required</th></tr></thead><tbody>
      ${sh.map(x => `<tr><td>${fmt(x.d)} ${x.d.slice(0, 4)}</td><td>${x.avail}</td><td>${x.req}</td></tr>`).join('')}</tbody></table>
      <p class="hint" style="margin-top:8px">The app never lets a pick cause one of these. They come from the template lines and training themselves.</p>`
      : '<div class="okbox">Every day meets its staffing minimum.</div>';
  }

  // ---------- actions ----------
  function addPick(pk) { S.picks.push(pk); S.sel = []; refresh(); }

  function guessInitials(name, config) {
    const parts = name.trim().split(/\s+/);
    const f = (parts[0] || '')[0]?.toUpperCase() || '', l = (parts[parts.length - 1] || '')[0]?.toUpperCase() || '';
    const known = new Set(config.training.flatMap(t => t.initials));
    const hits = Array.from(known).filter(x => x.length === 3 && x[0] === f && x[2] === l);
    return hits.length === 1 ? hits[0] : f + l;
  }

  document.addEventListener('click', e => {
    const b = e.target.closest('button, [data-day]');
    if (!b) return;
    const tab = b.dataset.tab;
    if (tab) return showTab(tab);
    const c = S.config;
    if (b.id === 'toBid') return showTab('bid');
    if (b.id === 'addPilot') {
      const n = c.pilots.length + 1;
      c.pilots.push({ id: E.uid(), name: `Pilot ${n}`, initials: '', seniority: n, allow: {} });
      return refresh();
    }
    if (b.id === 'copyRoster') {
      const from = $('#rosterSheet').value; if (!from) return;
      Xlsx.readRoster(S.wb, from).then(list => {
        if (c.pilots.length && !confirm(`Replace the ${c.pilots.length} pilot(s) listed with the names from ${from}?`)) return;
        c.pilots = list.map((p, i) => ({ id: E.uid(), name: p.name, initials: guessInitials(p.name, c), seniority: i + 1, allow: {} }));
        refresh();
      }).catch(err => alert(err.message));
      return;
    }
    if (b.dataset.pact) {
      const id = b.closest('tr').dataset.pid; const list = E.sortedPilots(c); const i = list.findIndex(p => p.id === id);
      if (b.dataset.pact === 'del') list.splice(i, 1);
      else { const j = b.dataset.pact === 'up' ? i - 1 : i + 1; if (j < 0 || j >= list.length) return; [list[i], list[j]] = [list[j], list[i]]; }
      list.forEach((p, k) => p.seniority = k + 1); c.pilots = list;
      return refresh();
    }
    if (b.id === 'addRound') {
      c.rounds.push({ line: false, days: { V: 5, HAC: 0, H: 0 }, minRun: 2, minGap: 2, mustAdjoin: false });
      return refresh();
    }
    if (b.dataset.ract === 'del') { c.rounds.splice(Number(b.closest('tr').dataset.ri), 1); return refresh(); }
    if (b.dataset.ract === 'up' || b.dataset.ract === 'down') {
      // Round 1 (line pick) stays first; the others can be put in any order.
      const i = Number(b.closest('tr').dataset.ri), j = b.dataset.ract === 'up' ? i - 1 : i + 1;
      if (j < 1 || j >= c.rounds.length) return;
      [c.rounds[i], c.rounds[j]] = [c.rounds[j], c.rounds[i]];
      return refresh();
    }
    if (b.id === 'addTraining') { const d = S.template.days; c.training.push({ id: E.uid(), initials: [], start: d[1], end: d[5], note: 'added' }); return refresh(); }
    if (b.dataset.tact === 'del') { c.training = c.training.filter(t => t.id !== b.closest('tr').dataset.tid); return refresh(); }

    const st = S.state; if (!st) return;
    if (b.id === 'undo') { S.picks.pop(); S.sel = []; return refresh(); }
    if (b.dataset.auto) return autoRun(b.dataset.auto);
    if (b.id === 'startOver') {
      if (!confirm('Clear all picks and start the bid over? Your settings and pilots are kept.')) return;
      S.picks = []; S.sel = []; return refresh();
    }
    if (b.id === 'approveTurn' && st.turn && st.turn.awaiting) return addPick({ round: st.turn.round, pilotId: st.turn.pilotId, type: 'approve' });
    if (b.id === 'sendBack' && st.turn && st.turn.awaiting) {
      // Reopen the turn by taking back the step that finished it (line pick, ranking, or Finish turn).
      const t = st.turn;
      for (let k = S.picks.length - 1; k >= 0; k--) {
        const pk = S.picks[k];
        if (pk.round === t.round && pk.pilotId === t.pilotId && ['done', 'line', 'rank'].includes(pk.type)) { S.picks.splice(k, 1); break; }
      }
      S.sel = []; S.rankFor = '';
      return refresh();
    }
    if (b.id === 'endTurn' && st.turn && !st.turn.awaiting) {
      const t = st.turn, rd = st.config.rounds[t.round];
      if (!rd.line && !E.isHolidayRound(rd)) {
        const mine = st.log.filter(pk => pk.type === 'off' && pk.round === t.round && pk.pilotId === t.pilotId);
        const list = mine.map(pk => `${pk.code} ${fmt(st.days[pk.dayIdx[0]])}${pk.dayIdx.length > 1 ? ' – ' + fmt(st.days[pk.dayIdx[pk.dayIdx.length - 1]]) : ''} (${pk.dayIdx.length})`).join('\n');
        const left = E.DAY_CODES.map(c => [c, E.remainingFor(st, t.round, t.pilotId, c)]).filter(([, n]) => n > 0).map(([c, n]) => `${n} ${c}`).join(', ');
        if (!confirm(`Finish this turn? These days can't be changed afterwards.\n\n${list || 'No days taken.'}${left ? `\n\nUnused this round: ${left}` : ''}`)) return;
      } else if (E.isHolidayRound(rd) && !confirm("Skip holidays? This pilot won't be awarded any holidays this round.")) return;
      return addPick({ round: t.round, pilotId: t.pilotId, type: 'done' });
    }
    if (b.dataset.unpick != null && st.turn && !st.turn.awaiting) {
      const k = Number(b.dataset.unpick), pk = S.picks[k];
      if (!pk || pk.type !== 'off' || pk.round !== st.turn.round || pk.pilotId !== st.turn.pilotId) return;
      S.picks.splice(k, 1); S.sel = [];
      return refresh();
    }
    if (b.dataset.line && st.turn) {
      const lineNo = Number(b.dataset.line);
      if (E.validateLinePick(st, lineNo).length) return;
      return addPick({ round: st.turn.round, pilotId: st.turn.pilotId, type: 'line', lineNo });
    }
    if (b.dataset.day != null) {
      const i = Number(b.dataset.day);
      if (st.turn && currentCode(st) === 'H') { // holidays: toggle each day
        S.sel = S.sel.includes(i) ? S.sel.filter(x => x !== i) : S.sel.concat(i).sort((a, z) => a - z);
        return renderBid();
      }
      if (S.sel.length !== 1) S.sel = [i];
      else { const a = Math.min(S.sel[0], i), z = Math.max(S.sel[0], i); S.sel = []; for (let k = a; k <= z; k++) S.sel.push(k); }
      return renderBid();
    }
    if (b.id === 'clearSel') { S.sel = []; return renderBid(); }
    if (b.dataset.radd) { S.rank.push(Number(b.dataset.radd)); return renderBid(); }
    if (b.id === 'rankAll') { E.eligibleHolidays(st, S.rankPilot).forEach(i => { if (!S.rank.includes(i)) S.rank.push(i); }); return renderBid(); }
    if (b.dataset.rmove) {
      const k = Number(b.dataset.ri), j = b.dataset.rmove === 'up' ? k - 1 : k + 1;
      if (b.dataset.rmove === 'del') S.rank.splice(k, 1); else [S.rank[k], S.rank[j]] = [S.rank[j], S.rank[k]];
      return renderBid();
    }
    if (b.id === 'submitRank' && st.turn) {
      if (E.validateRank(st, S.rank).length) return alert(E.validateRank(st, S.rank).join('\n'));
      const summary = S.rank.map((i, k) => `${k + 1}. ${fmt(st.days[i])}${holidayName(st.days[i]) ? ' ' + holidayName(st.days[i]) : ''}`).join('\n');
      if (!confirm(`Submit this holiday ranking? It can't be changed afterwards.\n\n${summary}`)) return;
      const order = S.rank.slice(); S.rank = [];
      return addPick({ round: st.turn.round, pilotId: st.turn.pilotId, type: 'rank', dayIdx: order });
    }
    if (b.dataset.code) { S.code = b.dataset.code; S.sel = []; return renderBid(); }
    if (b.id === 'addDays' && st.turn) {
      const code = currentCode(st);
      if (E.validateOffPick(st, S.sel, code).length) return;
      return addPick({ round: st.turn.round, pilotId: st.turn.pilotId, type: 'off', code, dayIdx: S.sel.slice() });
    }
    if (b.id === 'download') {
      Xlsx.writeResults(S.wb, S.template, st).then(blob => download(blob, S.fileName.replace(/\.xlsx$/i, '') + ` - ${S.sheetName} bid results.xlsx`)).catch(err => alert('Could not write the spreadsheet: ' + err.message));
      return;
    }
    if (b.id === 'saveBid') {
      if (!S.fileB64 || !st) return alert('Nothing to save yet.');
      const stamp = new Date().toISOString().slice(0, 16).replace('T', ' ').replace(':', '');
      Xlsx.writeResults(S.wb, S.template, st, { stateJson: JSON.stringify(bidState()) })
        .then(blob => download(blob, `${S.fileName.replace(/\.xlsx$/i, '').replace(/ - .* bid( results| saved.*)?$/i, '')} - ${S.sheetName} bid saved ${stamp}.xlsx`))
        .catch(err => alert('Could not save the bid: ' + err.message));
      return;
    }
  });

  document.addEventListener('change', e => {
    const el = e.target; const c = S.config;
    if (el.id === 'viewAs') { S.viewAs = el.value; S.sel = []; if (el.value !== 'admin') showTab('bid'); return renderAll(); }
    if (el.id === 'fileIn' && el.files[0]) {
      if (S.picks.length && !confirm('Loading a new spreadsheet clears the picks made so far. Continue?')) { el.value = ''; return; }
      el.files[0].arrayBuffer().then(buf => loadFile(buf, el.files[0].name)).catch(err => alert('Could not read that spreadsheet: ' + err.message));
      return;
    }
    if (el.id === 'openBid' && el.files[0]) {
      const f = el.files[0];
      const opening = /\.json$/i.test(f.name)
        ? f.text().then(txt => restore(JSON.parse(txt)))          // older bid files
        : f.arrayBuffer().then(buf => openSavedXlsx(buf, f.name)); // Excel saves
      opening.then(() => showTab('bid')).catch(err => alert('Could not open that saved bid: ' + err.message));
      el.value = '';
      return;
    }
    if (!c) return;
    if (el.id === 'sheetSel') {
      if (S.picks.length && !confirm('Switching sheets clears the picks made so far. Continue?')) return renderSetup();
      S.picks = []; S.original = null; S.origWarnings = null; useSheet(el.value, null).catch(err => alert(err.message)); return;
    }
    if (el.dataset.limit) { c.limits[el.dataset.limit] = Number(el.value) || 0; return refresh(); }
    if (el.id === 'hacText') { c.codeText = { ...(c.codeText || {}), HAC: el.value.trim() || 'HAC' }; return refresh(); }
    if (el.id === 'allowV' || el.id === 'allowHAC') { c.allowances[el.id.slice(5)] = Number(el.value) || 0; return refresh(); }
    if (el.dataset.p) {
      const p = c.pilots.find(x => x.id === el.closest('tr').dataset.pid);
      if (el.dataset.p === 'seniority') {
        // Move the pilot to that position; everyone else shifts and is renumbered 1, 2, 3…
        const list = E.sortedPilots(c).filter(x => x.id !== p.id);
        const pos = Math.min(Math.max(1, Math.round(Number(el.value) || 1)), list.length + 1);
        list.splice(pos - 1, 0, p);
        list.forEach((x, k) => { x.seniority = k + 1; });
        c.pilots = list;
        if (S.picks.length) alert('Seniority changed after bidding started. The picks made so far are replayed in the new order, and picks that no longer fit are dropped. Use Start over if anything looks off.');
        return refresh();
      }
      if (el.dataset.p === 'name') p.name = el.value;
      else if (el.dataset.p === 'initials') p.initials = el.value.trim().toUpperCase();
      else { p.allow = p.allow || {}; p.allow[el.dataset.p] = el.value === '' ? '' : Number(el.value); }
      return refresh();
    }
    if (el.dataset.r) {
      const r = c.rounds[Number(el.closest('tr').dataset.ri)];
      const k = el.dataset.r;
      if (el.type === 'checkbox') r[k] = el.checked;
      else if (E.DAY_CODES.includes(k)) {
        r.days[k] = Math.max(0, Number(el.value) || 0);
        if (k === 'H' && r.days.H > 0) { r.days.V = 0; r.days.HAC = 0; } // holidays get a round of their own
      }
      else r[k] = k === 'minGap' ? Math.max(0, Number(el.value) || 0) : Math.max(1, Number(el.value) || 1);
      return refresh();
    }
    if (el.dataset.tr) {
      const t = c.training.find(x => x.id === el.closest('tr').dataset.tid);
      if (el.dataset.tr === 'initials') t.initials = el.value.toUpperCase().split(/[^A-Z]+/).filter(Boolean);
      // Inputs hold the P-to-P window; sessions store the training days inside it.
      else if (el.value) t[el.dataset.tr] = shiftIso(el.value, el.dataset.tr === 'start' ? 1 : -1);
      return refresh();
    }
    if (el.id === 'requireApproval') { c.requireApproval = el.checked; return refresh(); }
    if (el.id === 'tmode') { c.trainingOnOffDay = el.value; return refresh(); }
    if (el.id === 'pcode') { c.positionCode = el.value; return refresh(); }
    if (el.id === 'holidays' || el.id === 'events' || el.id === 'safetyDays') {
      const list = Array.from(new Set(el.value.match(/\d{4}-\d{2}-\d{2}/g) || [])).sort();
      c[el.id] = list; return refresh();
    }
  });

  // Practice: let the app make picks until this turn, this round, or the whole bid is done.
  function autoRun(scope) {
    let st = S.state;
    const start = st.turn;
    if (!start) return;
    for (let guard = 0; st.turn && guard < 2000; guard++) {
      if (scope === 'turn' && (st.turn.round !== start.round || st.turn.pilotId !== start.pilotId)) break;
      if (scope === 'round' && st.turn.round !== start.round) break;
      const pk = E.autoPick(st);
      if (!pk) break;
      S.picks.push(pk);
      st = E.buildState(S.template, S.config, S.picks);
    }
    S.sel = [];
    refresh();
  }

  function showTab(name) {
    document.querySelectorAll('nav button').forEach(b => b.classList.toggle('active', b.dataset.tab === name));
    ['setup', 'bid', 'results'].forEach(t => $('#tab-' + t).classList.toggle('hidden', t !== name));
    window.scrollTo(0, 0);
  }

  // ---------- start ----------
  (async () => {
    let snap = null;
    try { snap = JSON.parse(localStorage.getItem(STORE) || 'null'); } catch (e) { snap = null; }
    if (snap && snap.fileB64) {
      try { await restore(snap); return; } catch (e) { console.warn('Could not restore saved bid', e); }
    }
    renderAll();
  })();

  window.PilotBidApp = { S, loadFile, refresh, addPick };
})();

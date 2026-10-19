// Bidding engine: pure logic, no DOM. Works on day indexes into template.days.
(function (root) {
  'use strict';

  const DAY_CODES = ['V', 'HAC', 'H']; // what a round can hand out, in the order they're offered
  // Codes that count as "an existing off block" when a round requires picks to adjoin one.
  const BLOCK_CODES = new Set(['OFF', 'V', 'HAC', 'H']);

  function dow(iso) { return new Date(iso + 'T12:00:00Z').getUTCDay(); }
  function isWeekend(iso) { const d = dow(iso); return d === 0 || d === 6; }

  function uid() { return Math.random().toString(36).slice(2, 10); }

  // A round: optional line pick plus a per-pilot limit for each day code.
  // Defaults: everyone picks a line (and training goes on), then holidays, then two rounds of
  // 7 V + 2 HAC and one of 5 V, as last year.
  function defaultRounds(holidayCount = 9) {
    return [
      { line: true, days: { V: 0, HAC: 0, H: 0 }, minRun: 2, minGap: 2, mustAdjoin: false },
      { line: false, days: { V: 0, HAC: 0, H: holidayCount }, minRun: 2, minGap: 2, mustAdjoin: false },
      { line: false, days: { V: 7, HAC: 2, H: 0 }, minRun: 2, minGap: 2, mustAdjoin: false },
      { line: false, days: { V: 7, HAC: 2, H: 0 }, minRun: 2, minGap: 2, mustAdjoin: false },
      { line: false, days: { V: 5, HAC: 0, H: 0 }, minRun: 2, minGap: 2, mustAdjoin: false },
    ];
  }

  // Bring rounds saved in the older one-code-per-round format up to date.
  function normalizeRound(r) {
    // Setups saved before the "days between groups" rule had min-in-a-row 1 by default; both rules are 2.
    const old = r.minGap == null;
    const minRun = old ? Math.max(2, Number(r.minRun) || 2) : r.minRun;
    const minGap = old ? 2 : r.minGap;
    if (r.days) return { line: !!r.line, days: { V: 0, HAC: 0, H: 0, ...r.days }, minRun, minGap, mustAdjoin: !!r.mustAdjoin };
    if (r.type === 'line') return { line: true, days: { V: 0, HAC: 0, H: 0 }, minRun, minGap, mustAdjoin: false };
    return { line: false, days: { V: 0, HAC: 0, H: 0, [r.code]: r.maxDays }, minRun, minGap, mustAdjoin: !!r.mustAdjoin };
  }
  function normalizeConfig(config) {
    const rounds = config.rounds.map(normalizeRound);
    // Round 1 is always the line pick on its own; any other line pick or days found there move to round 2.
    const first = rounds[0] && rounds[0].line ? rounds.shift() : null;
    if (first && Object.values(first.days).some(n => n > 0)) rounds.unshift({ ...first, line: false });
    rounds.forEach(r => { r.line = false; });
    config.rounds = [{ line: true, days: { V: 0, HAC: 0, H: 0 }, minRun: 2, minGap: 2, mustAdjoin: false }, ...rounds];
    config.codeText = config.codeText || { HAC: 'X' };
    return config;
  }

  function roundTotals(config) {
    const t = { V: 0, HAC: 0, H: 0 };
    config.rounds.forEach(r => DAY_CODES.forEach(c => { t[c] += Number(r.days[c]) || 0; }));
    return t;
  }

  // Build a starting configuration from what the reader detected in the template.
  function defaultConfig(template) {
    const req = template.required || {};
    const rounds = defaultRounds(template.holidays.length);
    const totals = roundTotals({ rounds });
    return {
      limits: {
        weekday: req.weekday ?? 5,
        weekend: req.weekend ?? 5,
        event: req.event ?? 7,
        holiday: req.holiday ?? 4,
        safety: req.safety ?? template.lines.length, // no number in the sheet: everyone works
      },
      holidays: template.holidays.slice(),
      events: template.events.slice(),
      safetyDays: (template.safety || []).slice(),
      allowances: { V: totals.V, HAC: totals.HAC },
      pilots: [],
      training: template.training.map(t => ({ id: uid(), initials: t.initials.slice(), start: t.start, end: t.end, note: t.note || '' })),
      trainingOnOffDay: 'mark', // 'mark' writes "OFF T"/"OFF P"; 'replace' writes "T"/"P"
      positionCode: 'P',
      codeText: { HAC: 'X' }, // how HAC days are written in the spreadsheet
      rounds,
    };
  }

  function sortedPilots(config) {
    return config.pilots.slice().sort((a, b) => a.seniority - b.seniority);
  }

  function allowanceFor(config, pilot, code) {
    if (code === 'H') return Infinity; // holidays are limited per round and by holiday staffing, not a yearly allowance
    const own = pilot.allow && pilot.allow[code];
    return own === '' || own == null ? (config.allowances[code] ?? 0) : Number(own);
  }

  function requiredFor(config, iso, sets) {
    // Safety / company training days first, then the sheet's own order: event, holiday, weekday/weekend.
    if (sets.safety.has(iso)) return config.limits.safety;
    if (sets.events.has(iso)) return config.limits.event;
    if (sets.holidays.has(iso)) return config.limits.holiday;
    return isWeekend(iso) ? config.limits.weekend : config.limits.weekday;
  }

  function trainingFor(config, pilot) {
    const ini = (pilot.initials || '').trim().toUpperCase();
    if (!ini) return [];
    return config.training.filter(t => t.initials.map(s => s.toUpperCase()).includes(ini));
  }

  // Put P/T codes for each of the pilot's training sessions onto a line's cells.
  // posCode: 'P' writes position days as P (2025 sheet); 'T' writes the whole window as T (2026 sheet).
  function applyTraining(cells, days, sessions, mode, posCode = 'P') {
    const index = new Map(days.map((d, i) => [d, i]));
    const put = (i, code) => {
      if (i == null || i < 0 || i >= cells.length) return;
      const cur = cells[i];
      if (cur === '') cells[i] = code;
      else if (cur === 'OFF') cells[i] = mode === 'replace' ? code : 'OFF ' + code;
    };
    for (const s of sessions) {
      const a = index.get(s.start), b = index.get(s.end);
      if (a == null || b == null) continue;
      put(a - 1, posCode);
      for (let i = a; i <= b; i++) put(i, 'T');
      put(b + 1, posCode);
    }
  }

  function codeName(code) { return code === 'H' ? 'holidays' : code; }

  // A holiday round: pilots rank holidays, then holidays are awarded round robin by seniority.
  function isHolidayRound(rd) { return !!rd && !rd.line && Number(rd.days.H) > 0; }

  // "line + 7 V + 2 HAC"
  function roundSummary(rd) {
    if (isHolidayRound(rd)) return `holiday ranking (up to ${rd.days.H} each)`;
    const parts = rd.line ? ['line'] : [];
    DAY_CODES.forEach(c => { if (rd.days[c] > 0) parts.push(`${rd.days[c]} ${codeName(c)}`); });
    return parts.join(' + ') || 'nothing';
  }

  function roundLabel(config, r) {
    const rd = config.rounds[r];
    return rd ? `Round ${r + 1}: ${roundSummary(rd)}` : '';
  }

  // Replay the pick log and derive the whole bid state.
  function buildState(template, config, picks) {
    const days = template.days;
    const sets = { events: new Set(config.events), holidays: new Set(config.holidays), safety: new Set(config.safetyDays || []) };
    const lines = template.lines.map(l => ({ lineNo: l.lineNo, row: l.row, pilotId: null, cells: l.cells.slice() }));
    const pilots = sortedPilots(config);
    const byId = new Map(pilots.map(p => [p.id, p]));
    const lineOf = new Map();
    const used = new Map(pilots.map(p => [p.id, { V: 0, HAC: 0, H: 0 }]));
    const roundUsed = {}; // `${round}:${pilotId}:${code}` -> days
    const finished = new Set(); // `${round}:${pilotId}`
    const log = [];

    const required = days.map(d => requiredFor(config, d, sets));

    const ranks = {}; // round -> Map(pilotId -> holiday day indexes, most wanted first)
    const awards = {}; // round -> [{pilotId, dayIdx, pass, rank}]
    const acted = new Set(); // `${round}:${pilotId}` turns where the pilot did something
    const approved = new Set(); // turns the admin has approved (when approval is required)
    const st = { template, config, days, sets, lines, pilots, byId, lineOf, used, roundUsed, finished, required, log, ranks, awards, acted, approved, turn: null };

    let pos = { round: 0, idx: 0 };
    advance(st, pos);
    for (const pk of picks) {
      const t = st.turn;
      if (!t) break;
      if (pk.type === 'rerank') continue; // rankings are final once submitted; ignore edits from older saved bids
      if (t.awaiting) { // a finished turn waiting for the admin: only an approval moves the bid on
        if (pk.type === 'approve' && pk.round === t.round && pk.pilotId === t.pilotId) {
          approved.add(`${t.round}:${t.pilotId}`);
          log.push(pk);
          advance(st, pos);
        }
        continue;
      }
      if (pk.type === 'approve') continue;
      if (pk.round !== t.round || pk.pilotId !== t.pilotId) continue; // stale entry, ignore
      applyPick(st, pk);
      log.push(pk);
      acted.add(`${t.round}:${t.pilotId}`);
      if (pk.type === 'done' || pk.type === 'rank' || turnDone(st, t.round, t.pilotId)) finished.add(`${t.round}:${t.pilotId}`);
      advance(st, pos);
    }
    return st;
  }

  function applyPick(st, pk) {
    if (pk.type === 'line') {
      const line = st.lines.find(l => l.lineNo === pk.lineNo);
      const pilot = st.byId.get(pk.pilotId);
      line.pilotId = pk.pilotId;
      st.lineOf.set(pk.pilotId, line);
      applyTraining(line.cells, st.days, trainingFor(st.config, pilot), st.config.trainingOnOffDay, st.config.positionCode || 'P');
    } else if (pk.type === 'rank') {
      (st.ranks[pk.round] = st.ranks[pk.round] || new Map()).set(pk.pilotId, pk.dayIdx.slice());
    } else if (pk.type === 'off') {
      const line = st.lineOf.get(pk.pilotId);
      for (const i of pk.dayIdx) line.cells[i] = pk.code;
      st.used.get(pk.pilotId)[pk.code] += pk.dayIdx.length;
      const k = `${pk.round}:${pk.pilotId}:${pk.code}`;
      st.roundUsed[k] = (st.roundUsed[k] || 0) + pk.dayIdx.length;
    }
  }

  // Days of this code the pilot can still take this round (round limit and yearly allowance).
  function remainingFor(st, round, pilotId, code) {
    const rd = st.config.rounds[round];
    const pilot = st.byId.get(pilotId);
    const inRound = (Number(rd.days[code]) || 0) - (st.roundUsed[`${round}:${pilotId}:${code}`] || 0);
    const allow = allowanceFor(st.config, pilot, code) - st.used.get(pilotId)[code];
    return Math.max(0, Math.min(inRound, allow));
  }

  function minFor(rd, code) { return code === 'H' ? 1 : Math.max(1, Number(rd.minRun) || 1); }

  // Codes this pilot can still pick this round.
  function openCodes(st, round, pilotId) {
    const rd = st.config.rounds[round];
    // Even 1 day left can be used, by adding it next to days off the pilot already has.
    return DAY_CODES.filter(c => remainingFor(st, round, pilotId, c) >= 1);
  }

  function needsLine(st, round, pilotId) {
    return st.config.rounds[round].line && !st.lineOf.has(pilotId) && st.lines.some(l => !l.pilotId);
  }

  function turnDone(st, round, pilotId) {
    if (st.finished.has(`${round}:${pilotId}`)) return true;
    if (isHolidayRound(st.config.rounds[round])) return !st.lineOf.has(pilotId); // waits for their ranking
    if (needsLine(st, round, pilotId)) return false;
    if (!st.lineOf.has(pilotId)) return true; // no line, nothing to add days to
    // Once a pilot has taken days this round, the turn stays open (so they can still remove some)
    // until they finish it. A pilot with nothing they could take is skipped.
    const madeAny = DAY_CODES.some(c => st.roundUsed[`${round}:${pilotId}:${c}`] > 0);
    return !madeAny && openCodes(st, round, pilotId).length === 0;
  }

  function advance(st, pos) {
    const n = st.pilots.length;
    while (pos.round < st.config.rounds.length) {
      while (pos.idx < n) {
        const p = st.pilots[pos.idx];
        if (!turnDone(st, pos.round, p.id)) {
          st.turn = { round: pos.round, pilotId: p.id };
          return;
        }
        const key = `${pos.round}:${p.id}`;
        if (st.config.requireApproval && st.acted.has(key) && !st.approved.has(key)) {
          st.turn = { round: pos.round, pilotId: p.id, awaiting: true }; // finished, waiting for the admin
          return;
        }
        pos.idx++;
      }
      // Everyone has ranked: hand out this round's holidays before the next round starts.
      if (isHolidayRound(st.config.rounds[pos.round])) awardHolidays(st, pos.round);
      pos.round++;
      pos.idx = 0;
    }
    st.turn = null;
  }

  // Round robin by seniority: each pass, every pilot still under the limit gets their highest-ranked
  // holiday that's still open (not training or already off on their line, and the day stays at or
  // above the holiday staffing minimum). Stops when everyone has their number or nothing more fits.
  function awardHolidays(st, round) {
    const target = Number(st.config.rounds[round].days.H) || 0;
    const ranked = st.ranks[round] || new Map();
    const prefs = new Map(st.pilots.map(p => [p.id, (ranked.get(p.id) || []).slice()]));
    const got = new Map(st.pilots.map(p => [p.id, 0]));
    const avail = availability(st);
    const list = [];
    for (let pass = 1, changed = true; changed; pass++) {
      changed = false;
      for (const p of st.pilots) {
        const line = st.lineOf.get(p.id);
        if (!line || got.get(p.id) >= target) continue;
        const want = prefs.get(p.id);
        while (want.length) {
          const i = want.shift();
          if (line.cells[i] !== '' || avail[i] - 1 < st.required[i]) continue; // can't have it; try their next choice
          line.cells[i] = 'H';
          avail[i]--;
          got.set(p.id, got.get(p.id) + 1);
          st.used.get(p.id).H++;
          list.push({ pilotId: p.id, dayIdx: i, pass, rank: ranked.get(p.id).indexOf(i) + 1 });
          changed = true;
          break;
        }
      }
    }
    st.awards[round] = list;
    st.log.push({ type: 'award', round, awards: list });
  }

  // Holidays this pilot can rank: ones where their line is a workday (not training or already off).
  function eligibleHolidays(st, pilotId) {
    const line = st.lineOf.get(pilotId);
    if (!line) return [];
    return st.days.map((d, i) => i).filter(i => st.sets.holidays.has(st.days[i]) && line.cells[i] === '');
  }

  function validateRank(st, order, pilotId) {
    const t = st.turn;
    if (!t || !isHolidayRound(st.config.rounds[t.round])) return ['Not a holiday ranking round'];
    if (t.awaiting) return ['This turn is waiting for admin approval'];
    const ok = new Set(eligibleHolidays(st, pilotId || t.pilotId));
    if (new Set(order).size !== order.length) return ['A holiday is listed twice'];
    const bad = order.filter(i => !ok.has(i));
    return bad.length ? bad.map(i => `${st.days[i]}: not a holiday you can take`) : [];
  }

  // Pilots available per day: line rows with a blank cell, exactly like the sheet's COUNTBLANK row.
  function availability(st) {
    return st.days.map((_, i) => st.lines.reduce((n, l) => n + (l.cells[i] === '' ? 1 : 0), 0));
  }

  // Availability for checking a day off: also holds back pilots who haven't picked a line yet but
  // will be in training (or positioning) that day, since that training lands on whatever line they pick.
  function bidAvailability(st) {
    const avail = availability(st);
    const index = new Map(st.days.map((d, i) => [d, i]));
    const pending = st.days.map(() => 0);
    for (const p of st.pilots) {
      if (st.lineOf.has(p.id)) continue;
      const busy = new Set();
      for (const s of trainingFor(st.config, p)) {
        const a = index.get(s.start), b = index.get(s.end);
        if (a == null || b == null) continue;
        for (let i = Math.max(0, a - 1); i <= Math.min(st.days.length - 1, b + 1); i++) busy.add(i);
      }
      busy.forEach(i => pending[i]++);
    }
    const openBlank = st.days.map((_, i) => st.lines.reduce((n, l) => n + (!l.pilotId && l.cells[i] === '' ? 1 : 0), 0));
    return avail.map((a, i) => a - Math.min(pending[i], openBlank[i]));
  }

  // Can this day be taken as `code` by this pilot? Returns null if OK, else a reason.
  function dayBlockReason(st, pilotId, i, avail, code) {
    const line = st.lineOf.get(pilotId);
    if (!line) return 'No line';
    const c = line.cells[i];
    if (code === 'H' && !st.sets.holidays.has(st.days[i])) return 'Not a holiday';
    if (c !== '') return c.includes('T') || c.includes('P') ? 'Training' : 'Already off';
    if (avail[i] - 1 < st.required[i]) return st.sets.safety.has(st.days[i])
      ? `Safety/training day: ${st.required[i]} pilots required`
      : `Staffing: ${avail[i]} available (counting training still to be placed), ${st.required[i]} required`;
    return null;
  }

  // Validate days the current pilot wants to take as `code`.
  function validateOffPick(st, dayIdx, code) {
    const t = st.turn;
    const errors = [];
    if (!t) return ['Bidding is finished'];
    if (t.awaiting) return ['This turn is waiting for admin approval'];
    const rd = st.config.rounds[t.round];
    if (needsLine(st, t.round, t.pilotId)) return ['Pick a line first'];
    if (!dayIdx.length) return ['Pick at least one day'];
    const sorted = dayIdx.slice().sort((a, b) => a - b);
    if (code !== 'H') { // holidays are picked one by one, no run length rules
      for (let k = 1; k < sorted.length; k++) if (sorted[k] !== sorted[k - 1] + 1) { errors.push('Days must be consecutive'); break; }
      // The minimum applies to the group of days off the pick ends up in, so a short pick is fine
      // when it's added onto line days off, V, HAC or a holiday the pilot already has.
      const minRun = minFor(rd, code);
      if (sorted.length < minRun) {
        const cells = st.lineOf.get(t.pilotId).cells;
        let a = sorted[0], z = sorted[sorted.length - 1];
        while (a - 1 >= 0 && BLOCK_CODES.has(cells[a - 1])) a--;
        while (z + 1 < cells.length && BLOCK_CODES.has(cells[z + 1])) z++;
        if (z - a + 1 < minRun) errors.push(`At least ${minRun} days in a row, unless they're added next to days off you already have`);
      }
    }
    const rem = remainingFor(st, t.round, t.pilotId, code);
    if (sorted.length > rem) errors.push(`Only ${rem} ${codeName(code)} left this round`);
    const avail = bidAvailability(st);
    for (const i of sorted) {
      const why = dayBlockReason(st, t.pilotId, i, avail, code);
      if (why) errors.push(`${st.days[i]}: ${why}`);
    }
    // Groups of days off (line days off, V, HAC, holidays) must either touch or be separated by at
    // least minGap working days, so a pick can't leave a single working day stranded between two.
    const gap = Number(rd.minGap) || 0;
    if (gap > 1 && code !== 'H') {
      const cells = st.lineOf.get(t.pilotId).cells;
      const check = (from, step) => {
        let n = 0, k = from;
        while (k >= 0 && k < cells.length && !BLOCK_CODES.has(cells[k])) { n++; k += step; }
        if (n > 0 && n < gap && k >= 0 && k < cells.length)
          errors.push(`Leaves only ${n} working day${n === 1 ? '' : 's'} between these days and your days off on ${dayName(st.days[k])} ${Number(st.days[k].slice(5, 7))}/${Number(st.days[k].slice(8))}. Leave at least ${gap}, or make them touch.`);
      };
      check(sorted[0] - 1, -1);
      check(sorted[sorted.length - 1] + 1, 1);
    }
    if (rd.mustAdjoin && code !== 'H') {
      const cells = st.lineOf.get(t.pilotId).cells;
      const before = cells[sorted[0] - 1], after = cells[sorted[sorted.length - 1] + 1];
      if (!BLOCK_CODES.has(before) && !BLOCK_CODES.has(after)) errors.push('Must touch an existing block of days off');
    }
    return errors;
  }

  function validateLinePick(st, lineNo) {
    const t = st.turn;
    if (!t) return ['Bidding is finished'];
    if (t.awaiting) return ['This turn is waiting for admin approval'];
    if (!needsLine(st, t.round, t.pilotId)) return ['No line pick due now'];
    const line = st.lines.find(l => l.lineNo === lineNo);
    if (!line) return ['No such line'];
    if (line.pilotId) return ['Line already taken'];
    return [];
  }

  // Summary of a line for the line picker, including clashes with this pilot's training.
  function lineInfo(st, line, pilot) {
    const base = st.template.lines.find(l => l.row === line.row).cells; // before training/picks
    const offIdx = [];
    base.forEach((c, i) => { if (c === 'OFF') offIdx.push(i); });
    let pattern = '';
    for (let k = 0; k < offIdx.length; k++) {
      if (k === 0 || offIdx[k] !== offIdx[k - 1] + 1) {
        // first full block start weekday
        const startIso = st.days[offIdx[k]];
        let len = 1; while (offIdx[k + len] === offIdx[k] + len) len++;
        if (len >= 5) { pattern = `${dayName(startIso)}–${dayName(st.days[offIdx[k] + len - 1])}`; break; }
      }
    }
    let clash = 0;
    if (pilot) {
      const tmp = base.slice();
      applyTraining(tmp, st.days, trainingFor(st.config, pilot), 'mark');
      tmp.forEach(c => { if (c.startsWith('OFF ')) clash++; });
    }
    const firstOff = offIdx.length ? st.days[offIdx[0]] : null;
    const holidays = base.filter(c => c === 'H').length;
    return { offDays: offIdx.length, pattern, clash, firstOff, holidays };
  }

  function dayName(iso) { return ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][dow(iso)]; }

  // Days below the staffing minimum.
  function shortages(st) {
    const avail = availability(st);
    return st.days.map((d, i) => ({ d, i, avail: avail[i], req: st.required[i] })).filter(x => x.avail < x.req);
  }

  // Practice mode: a reasonable pick for whoever's turn it is. Line: the free line with the fewest
  // training clashes. Days: the longest allowed run for the first code with days left, preferring
  // to extend an existing block, starting at a different point for each pilot so picks spread out.
  function autoPick(st) {
    const t = st.turn;
    if (!t) return null;
    if (t.awaiting) return { round: t.round, pilotId: t.pilotId, type: 'approve' }; // practice: approve and move on
    const rd = st.config.rounds[t.round];
    const pilot = st.byId.get(t.pilotId);
    const base = { round: t.round, pilotId: t.pilotId };
    if (isHolidayRound(rd)) {
      // Practice ranking: the pilot's eligible holidays in a shuffled order that differs per pilot.
      let seed = st.pilots.indexOf(pilot) * 7919 + 17;
      const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
      const order = eligibleHolidays(st, t.pilotId).map(i => [rnd(), i]).sort((a, b) => a[0] - b[0]).map(x => x[1]);
      return { ...base, type: 'rank', dayIdx: order };
    }
    if (needsLine(st, t.round, t.pilotId)) {
      const free = st.lines.filter(l => !l.pilotId);
      free.sort((a, b) => lineInfo(st, a, pilot).clash - lineInfo(st, b, pilot).clash || a.lineNo - b.lineNo);
      return { ...base, type: 'line', lineNo: free[0].lineNo };
    }
    const avail = bidAvailability(st);
    const n = st.days.length;
    const cells = st.lineOf.get(t.pilotId).cells;
    const touches = (a, z) => BLOCK_CODES.has(cells[a - 1]) || BLOCK_CODES.has(cells[z + 1]);
    for (const code of openCodes(st, t.round, t.pilotId)) {
      const rem = remainingFor(st, t.round, t.pilotId, code);
      const ok = st.days.map((_, i) => !dayBlockReason(st, t.pilotId, i, avail, code));
      const pick = dayIdx => ({ ...base, type: 'off', code, dayIdx });
      if (code === 'H') {
        const hol = st.days.map((d, i) => i).filter(i => ok[i]).slice(0, rem);
        if (hol.length && !validateOffPick(st, hol, code).length) return pick(hol);
        continue;
      }
      const seed = (st.pilots.indexOf(pilot) * 53 + t.round * 97 + (code === 'HAC' ? 180 : 0)) % n;
      for (let len = rem; len >= 1; len--) {
        let fallback = null;
        for (let k = 0; k < n; k++) {
          const a = (seed + k) % n, z = a + len - 1;
          if (z >= n) continue;
          let good = true;
          for (let i = a; i <= z; i++) if (!ok[i]) { good = false; break; }
          if (!good) continue;
          const run = []; for (let i = a; i <= z; i++) run.push(i);
          if (touches(a, z)) { if (!validateOffPick(st, run, code).length) return pick(run); }
          else if (!fallback && !rd.mustAdjoin) fallback = run;
        }
        if (fallback && !validateOffPick(st, fallback, code).length) return pick(fallback);
      }
    }
    return { ...base, type: 'done' };
  }

  const api = {
    autoPick, isHolidayRound, eligibleHolidays, validateRank, DAY_CODES, uid, dow, isWeekend, dayName, defaultConfig, defaultRounds, normalizeConfig,
    roundTotals, roundSummary, codeName, sortedPilots, allowanceFor, trainingFor, applyTraining, buildState,
    availability, bidAvailability, dayBlockReason, validateOffPick, validateLinePick, remainingFor, openCodes, needsLine,
    lineInfo, roundLabel, shortages,
  };
  if (typeof module !== 'undefined') module.exports = api; else root.Engine = api;
})(this);

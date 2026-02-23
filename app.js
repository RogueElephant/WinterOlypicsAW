/* global XLSX */

const STORAGE_KEY = 'aw-winter-olympics-v1';
const RESULTS_FORMAT_VERSION = 1;

function uid(prefix = 'id') {
  return `${prefix}_${Math.random().toString(16).slice(2)}_${Date.now().toString(16)}`;
}

function normalizeHeader(header) {
  return String(header || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function parseMaybeNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  const n = Number(String(value).replace(',', '.').trim());
  return Number.isFinite(n) ? n : null;
}

function isCsvFile(file) {
  const name = String(file?.name || '').toLowerCase();
  const type = String(file?.type || '').toLowerCase();
  return name.endsWith('.csv') || type.includes('text/csv') || type.includes('application/vnd.ms-excel');
}

function parseCsv(text) {
  // Minimal CSV parser with quotes support.
  const rows = [];
  let row = [];
  let cell = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        const next = text[i + 1];
        if (next === '"') {
          cell += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        cell += ch;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      continue;
    }

    if (ch === ',') {
      row.push(cell);
      cell = '';
      continue;
    }

    if (ch === '\n') {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
      continue;
    }

    if (ch === '\r') {
      // ignore
      continue;
    }

    cell += ch;
  }

  // Last line
  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }

  if (!rows.length) return [];

  const headers = rows[0].map((h) => String(h || '').trim());
  const out = [];
  for (const r of rows.slice(1)) {
    if (!r.some((x) => String(x || '').trim().length)) continue;
    const obj = {};
    for (let i = 0; i < headers.length; i += 1) {
      const key = headers[i] || `Column ${i + 1}`;
      obj[key] = r[i] ?? '';
    }
    out.push(obj);
  }
  return out;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

function csvEscape(value) {
  const s = String(value ?? '');
  if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function toCsv(rows, headers) {
  const out = [];
  out.push(headers.map(csvEscape).join(','));
  for (const r of rows) {
    out.push(headers.map((h) => csvEscape(r[h] ?? '')).join(','));
  }
  return out.join('\n');
}

function defaultState() {
  return {
    teams: [], // {id, name, members: [..]}
    events: {
      1: { matches: [] }, // match: {id, aTeamId, bTeamId, winnerTeamId|null}  — Bobsled
      2: { matches: [] }, // match: {id, aTeamId, bTeamId, winnerTeamId|null}
      3: { matches: [] },
      4: { times: {} }, // teamId -> seconds
      5: { matches: [] },
    },
    settings: {
      scoringMode: 'default',
      activeView: 'teams',
    },
  };
}

function getTimestampForFilename() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
}

function buildResultsExportRows() {
  const headers = [
    'RecordType',
    'Event',
    'GroupId',
    'Slot',
    'TeamId',
    'TeamName',
    'Member1',
    'Member2',
    'Member3',
    'Member4',
    'ATeamId',
    'BTeamId',
    'WinnerTeamId',
    'Place',
    'Seconds',
    'Key',
    'Value',
  ];

  const rows = [];
  rows.push({ RecordType: 'meta', Key: 'app', Value: 'Afterwork Winter Olympics' });
  rows.push({ RecordType: 'meta', Key: 'formatVersion', Value: String(RESULTS_FORMAT_VERSION) });
  rows.push({ RecordType: 'meta', Key: 'exportedAt', Value: new Date().toISOString() });

  for (const t of state.teams) {
    rows.push({
      RecordType: 'team',
      TeamId: t.id,
      TeamName: t.name,
      Member1: t.members?.[0] || '',
      Member2: t.members?.[1] || '',
      Member3: t.members?.[2] || '',
      Member4: t.members?.[3] || '',
    });
  }

  // Event 1/2/3/5 matches
  for (const eventId of [1, 2, 3, 5]) {
    for (const m of state.events[eventId].matches || []) {
      rows.push({
        RecordType: `e${eventId}_match`,
        Event: String(eventId),
        GroupId: m.id || uid(`match${eventId}`),
        ATeamId: m.aTeamId || '',
        BTeamId: m.bTeamId || '',
        WinnerTeamId: m.winnerTeamId || '',
      });
    }
  }

  // Event 4 times
  for (const eventId of [4]) {
    const times = state.events[eventId].times || {};
    for (const [teamId, seconds] of Object.entries(times)) {
      rows.push({
        RecordType: 'e4_time',
        Event: String(eventId),
        TeamId: teamId,
        Seconds: Number.isFinite(seconds) ? String(seconds) : '',
      });
    }
  }

  for (const [key, value] of Object.entries(state.settings || {})) {
    rows.push({ RecordType: 'setting', Key: key, Value: String(value ?? '') });
  }

  return { headers, rows };
}

function exportResultsToCsv() {
  const { headers, rows } = buildResultsExportRows();
  const csv = toCsv(rows, headers);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
  downloadBlob(blob, `afterwork-winter-olympics_results_${getTimestampForFilename()}.csv`);
}

function exportResultsToXlsx() {
  if (typeof XLSX === 'undefined') {
    alert('Excel export is not available (XLSX library not loaded). Use Save (.csv) instead.');
    return;
  }

  const wb = XLSX.utils.book_new();

  const meta = [
    { Key: 'app', Value: 'Afterwork Winter Olympics' },
    { Key: 'formatVersion', Value: RESULTS_FORMAT_VERSION },
    { Key: 'exportedAt', Value: new Date().toISOString() },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(meta), 'Meta');

  const teams = state.teams.map((t) => ({
    TeamId: t.id,
    TeamName: t.name,
    Member1: t.members?.[0] || '',
    Member2: t.members?.[1] || '',
    Member3: t.members?.[2] || '',
    Member4: t.members?.[3] || '',
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(teams), 'Teams');

  for (const eventId of [1, 2, 3, 5]) {
    const eMatches = (state.events[eventId].matches || []).map((m) => ({
      MatchId: m.id,
      ATeamId: m.aTeamId || '',
      BTeamId: m.bTeamId || '',
      WinnerTeamId: m.winnerTeamId || '',
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(eMatches), `Event${eventId}_Matches`);
  }

  const e4 = Object.entries(state.events[4].times || {}).map(([teamId, seconds]) => ({ TeamId: teamId, Seconds: seconds }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(e4), 'Event4_Times');

  const settings = Object.entries(state.settings || {}).map(([key, value]) => ({ Key: key, Value: String(value ?? '') }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(settings), 'Settings');

  XLSX.writeFile(wb, `afterwork-winter-olympics_results_${getTimestampForFilename()}.xlsx`);
}

function importResultsFromCsvRows(rows) {
  const next = defaultState();
  next.teams = [];

  // team id -> team
  const teamsById = new Map();

  for (const r of rows) {
    const recordType = String(r.RecordType || '').trim().toLowerCase();
    if (!recordType) continue;

    if (recordType === 'team') {
      const teamId = String(r.TeamId || '').trim() || uid('team');
      const teamName = String(r.TeamName || '').trim();
      if (!teamName) continue;
      const members = [r.Member1, r.Member2, r.Member3, r.Member4]
        .map((x) => String(x || '').trim())
        .filter(Boolean);
      const team = { id: teamId, name: teamName, members };
      teamsById.set(teamId, team);
      continue;
    }

    if (recordType === 'e1_match' || recordType === 'e2_match' || recordType === 'e3_match' || recordType === 'e5_match') {
      const eventId = recordType === 'e1_match' ? 1 : recordType === 'e2_match' ? 2 : recordType === 'e3_match' ? 3 : 5;
      const matchId = String(r.GroupId || '').trim() || uid(`match${eventId}`);
      const aTeamId = String(r.ATeamId || '').trim();
      const bTeamId = String(r.BTeamId || '').trim();
      const winnerTeamId = String(r.WinnerTeamId || '').trim();
      next.events[eventId].matches.push({ id: matchId, aTeamId, bTeamId, winnerTeamId });
      continue;
    }

    if (recordType === 'e4_time') {
      const eventId = 4;
      const teamId = String(r.TeamId || '').trim();
      const seconds = parseMaybeNumber(r.Seconds);
      if (!teamId || seconds === null) continue;
      next.events[eventId].times[teamId] = seconds;
      continue;
    }

    if (recordType === 'setting') {
      const key = String(r.Key || '').trim();
      if (!key) continue;
      next.settings[key] = String(r.Value ?? '');
      continue;
    }
  }

  next.teams = ensureUniqueTeamNames(Array.from(teamsById.values()));

  return next;
}

function importResultsFromWorkbook(wb) {
  const next = defaultState();

  const sheet = (name) => wb.Sheets[name];
  const read = (name) => (sheet(name) ? XLSX.utils.sheet_to_json(sheet(name), { defval: '' }) : []);

  const meta = read('Meta');
  const metaMap = new Map(meta.map((r) => [String(r.Key || '').toLowerCase().trim(), String(r.Value || '').trim()]));
  if (metaMap.get('app') !== 'Afterwork Winter Olympics') throw new Error('Not a results workbook');

  const teamsRows = read('Teams');
  const teams = [];
  for (const r of teamsRows) {
    const teamName = String(r.TeamName || r.Team || r['Team Name'] || '').trim();
    if (!teamName) continue;
    const teamId = String(r.TeamId || '').trim() || uid('team');
    const members = [r.Member1, r.Member2, r.Member3, r.Member4]
      .map((x) => String(x || '').trim())
      .filter(Boolean);
    teams.push({ id: teamId, name: teamName, members });
  }
  next.teams = ensureUniqueTeamNames(teams);

  // Event 1/2/3/5 matches
  for (const eventId of [1, 2, 3, 5]) {
    for (const r of read(`Event${eventId}_Matches`)) {
      next.events[eventId].matches.push({
        id: String(r.MatchId || '').trim() || uid(`match${eventId}`),
        aTeamId: String(r.ATeamId || '').trim(),
        bTeamId: String(r.BTeamId || '').trim(),
        winnerTeamId: String(r.WinnerTeamId || '').trim(),
      });
    }
  }

  for (const r of read('Event4_Times')) {
    const teamId = String(r.TeamId || '').trim();
    const seconds = parseMaybeNumber(r.Seconds);
    if (!teamId || seconds === null) continue;
    next.events[4].times[teamId] = seconds;
  }
  for (const r of read('Settings')) {
    const key = String(r.Key || '').trim();
    if (!key) continue;
    next.settings[key] = String(r.Value ?? '');
  }

  return next;
}

async function importSavedResultsFile(file) {
  const name = String(file?.name || '').toLowerCase();
  const useCsv = isCsvFile(file) || name.endsWith('.csv');

  if (useCsv) {
    const text = await file.text();
    const rows = parseCsv(text);
    const recordTypes = new Set(rows.map((r) => String(r.RecordType || '').trim().toLowerCase()).filter(Boolean));
    if (!recordTypes.has('team') || (!recordTypes.has('e1_match') && !recordTypes.has('e2_match') && !recordTypes.has('e4_time') && !recordTypes.has('e5_match'))) {
      throw new Error('Not a saved results CSV');
    }
    return importResultsFromCsvRows(rows);
  }

  if (typeof XLSX === 'undefined') {
    throw new Error('Excel import not available (XLSX library not loaded)');
  }

  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: 'array' });

  // If it looks like our results workbook, import it; otherwise treat as team import (existing flow)
  if (wb.Sheets.Meta && wb.Sheets.Teams) {
    return importResultsFromWorkbook(wb);
  }

  throw new Error('Not a saved results workbook');
}

let state = loadState();

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    const parsed = JSON.parse(raw);
    const merged = {
      ...defaultState(),
      ...parsed,
      events: {
        ...defaultState().events,
        ...(parsed.events || {}),
      },
      settings: {
        ...defaultState().settings,
        ...(parsed.settings || {}),
      },
    };

    // Backward/forward compatibility for event shapes
    if (!merged.events[1]?.matches) merged.events[1] = { ...(merged.events[1] || {}), matches: [] };
    if (!merged.events[2]?.matches) merged.events[2] = { ...(merged.events[2] || {}), matches: [] };
    if (!merged.events[3]?.matches) merged.events[3] = { ...(merged.events[3] || {}), matches: [] };
    if (!merged.events[4]?.times) merged.events[4] = { ...(merged.events[4] || {}), times: {} };
    if (!merged.events[5]?.matches) merged.events[5] = { ...(merged.events[5] || {}), matches: [] };

    return merged;
  } catch {
    return defaultState();
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function setStatus(msg, kind = 'info') {
  const el = document.getElementById('importStatus');
  if (!el) return;
  el.textContent = msg;
  el.dataset.kind = kind;
}

function getActiveViewId() {
  const v = state.settings?.activeView;
  return v === 'teams' || v === 'points' || v === 'events' ? v : 'teams';
}

function setActiveView(viewId) {
  const id = viewId === 'points' || viewId === 'events' ? viewId : 'teams';
  state.settings.activeView = id;
  saveState();
  document.body.dataset.activeView = id;

  for (const b of document.querySelectorAll('#viewTabs .tab')) b.classList.remove('is-active');
  const btn = document.querySelector(`#viewTabs .tab[data-view="${id}"]`);
  if (btn) btn.classList.add('is-active');

  for (const v of document.querySelectorAll('.view')) v.classList.remove('is-active');
  const viewEl = document.getElementById(`view${id[0].toUpperCase()}${id.slice(1)}`);
  if (viewEl) viewEl.classList.add('is-active');

  const teamsOnlyVisible = id === 'teams';
  const resetBtn = document.getElementById('resetBtn');
  const createMatchupsBtn = document.getElementById('createMatchupsBtn');
  if (resetBtn) resetBtn.hidden = !teamsOnlyVisible;
  if (createMatchupsBtn) createMatchupsBtn.hidden = !teamsOnlyVisible;
}

function getTeamsById() {
  const map = new Map();
  for (const t of state.teams) map.set(t.id, t);
  return map;
}

function ensureUniqueTeamNames(teams) {
  const seen = new Set();
  const out = [];
  for (const team of teams) {
    let name = team.name.trim();
    if (!name) continue;
    if (!seen.has(name.toLowerCase())) {
      out.push({ ...team, name });
      seen.add(name.toLowerCase());
      continue;
    }
    let i = 2;
    while (seen.has(`${name} (${i})`.toLowerCase())) i += 1;
    out.push({ ...team, name: `${name} (${i})` });
    seen.add(`${name} (${i})`.toLowerCase());
  }
  return out;
}

function shuffleArray(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function createRandomMatchupsForEvents() {
  const matchEventIds = [1, 2, 3, 5];
  const teamIds = state.teams.map((t) => t.id);

  if (teamIds.length < 2) {
    alert('Need at least 2 teams to create matchups.');
    return;
  }

  const byeTeamsByEvent = [];

  for (const eventId of matchEventIds) {
    const shuffled = shuffleArray(teamIds);
    const matches = [];

    for (let i = 0; i + 1 < shuffled.length; i += 2) {
      matches.push({
        id: uid(`match${eventId}`),
        aTeamId: shuffled[i],
        bTeamId: shuffled[i + 1],
        winnerTeamId: '',
      });
    }

    if (shuffled.length % 2 === 1) {
      const byeTeamId = shuffled[shuffled.length - 1];
      const byeTeamName = getTeamsById().get(byeTeamId)?.name || 'Unknown team';
      byeTeamsByEvent.push(`Event ${eventId}: ${byeTeamName} gets a bye`);
    }

    state.events[eventId].matches = matches;
  }

  saveState();
  renderAll();

  const byeText = byeTeamsByEvent.length ? ` ${byeTeamsByEvent.join(' | ')}` : '';
  setStatus(`Random matchups created for Events 1, 2, 3 and 5.${byeText}`);
}

function importTeamsFromRows(rows) {
  if (!rows.length) return { teams: [], warnings: ['No rows found.'] };

  const headers = Object.keys(rows[0] || {});
  const normalizedToKey = new Map();
  for (const k of headers) normalizedToKey.set(normalizeHeader(k), k);

  const teamKey =
    normalizedToKey.get('team') ||
    normalizedToKey.get('team name') ||
    normalizedToKey.get('teamname') ||
    headers[0];

  const memberKeys = [];
  const candidates = [
    'member 1',
    'member 2',
    'member 3',
    'member 4',
    'player 1',
    'player 2',
    'player 3',
    'player 4',
    'name 1',
    'name 2',
    'name 3',
    'name 4',
    'team member 1',
    'team member 2',
    'team member 3',
    'team member 4',
  ];

  for (const c of candidates) {
    const key = normalizedToKey.get(c);
    if (key) memberKeys.push(key);
  }

  // If headers are not named, assume next 4 columns after team.
  if (memberKeys.length < 4 && headers.length >= 5) {
    const idx = headers.indexOf(teamKey);
    const fallback = headers.slice(idx + 1, idx + 5);
    for (const f of fallback) {
      if (!memberKeys.includes(f)) memberKeys.push(f);
    }
  }

  // Detect open pool column
  const openPoolKey =
    normalizedToKey.get('open team pool') ||
    normalizedToKey.get('open pool') ||
    normalizedToKey.get('pool') ||
    null;

  const warnings = [];
  if (!teamKey) warnings.push('Could not detect team name column.');

  const teams = [];
  const openPoolNames = [];

  for (const row of rows) {
    // Collect named teams
    const rawName = row[teamKey];
    const name = String(rawName || '').trim();
    if (name) {
      const members = [];
      for (const mk of memberKeys.slice(0, 4)) {
        const val = row[mk];
        const m = String(val || '').trim();
        if (m) members.push(m);
      }
      teams.push({ id: uid('team'), name, members });
    }

    // Collect open pool people
    if (openPoolKey) {
      const poolName = String(row[openPoolKey] || '').trim();
      if (poolName) openPoolNames.push(poolName);
    }
  }

  // Randomly assign open pool people into teams of 3 or 4
  if (openPoolNames.length > 0) {
    const shuffled = shuffleArray(openPoolNames);
    const poolTeams = [];
    let remaining = shuffled.length;

    // Decide team sizes: prefer 4, but use 3 if needed to avoid leftovers of 1 or 2
    const teamSizes = [];
    while (remaining > 0) {
      if (remaining === 3 || remaining === 6) {
        teamSizes.push(3);
        remaining -= 3;
      } else if (remaining <= 4) {
        teamSizes.push(remaining);
        remaining = 0;
      } else {
        teamSizes.push(4);
        remaining -= 4;
      }
    }

    let offset = 0;
    for (let i = 0; i < teamSizes.length; i += 1) {
      const size = teamSizes[i];
      const members = shuffled.slice(offset, offset + size);
      const teamName = i === 0 ? 'Enter Team Name' : `Enter Team Name ${i + 1}`;
      poolTeams.push({ id: uid('team'), name: teamName, members });
      offset += size;
    }

    teams.push(...poolTeams);
    warnings.push(`${openPoolNames.length} open pool player(s) randomly assigned to ${poolTeams.length} team(s).`);
  }

  const unique = ensureUniqueTeamNames(teams);
  if (!unique.length) warnings.push('No teams imported. Check your sheet format.');
  if (unique.length !== teams.length) warnings.push('Duplicate team names were renamed automatically.');

  return { teams: unique, warnings };
}

function computePoints() {
  const teamIds = state.teams.map((t) => t.id);

  const points = {}; // teamId -> {1..5, total}
  for (const id of teamIds) {
    points[id] = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, total: 0 };
  }

  // Event 1, 2, 3 & 5: 1v1 matches — winner=3, loser=0
  for (const eventId of [1, 2, 3, 5]) {
    const matches = state.events[eventId].matches || [];
    for (const m of matches) {
      if (!m.aTeamId || !m.bTeamId) continue;
      if (!m.winnerTeamId) continue;
      points[m.winnerTeamId][eventId] += 3;
    }
  }

  // Event 4: timed
  // Top 4 score: 1st=5, 2nd=3, 3rd=2, 4th=1
  for (const eventId of [4]) {
    const times = state.events[eventId].times || {};
    const entries = Object.entries(times)
      .map(([teamId, seconds]) => ({ teamId, seconds }))
      .filter((x) => Number.isFinite(x.seconds));

    entries.sort((a, b) => a.seconds - b.seconds);

    const top4Score = [5, 3, 2, 1];
    const limit = Math.min(4, entries.length);
    for (let i = 0; i < limit; i += 1) {
      const teamId = entries[i].teamId;
      points[teamId][eventId] += top4Score[i];
    }
  }

  for (const id of teamIds) {
    points[id].total = points[id][1] + points[id][2] + points[id][3] + points[id][4] + points[id][5];
  }

  return points;
}

function renderTeamsTable() {
  const tbody = document.querySelector('#teamsTable tbody');
  tbody.innerHTML = '';

  for (const t of state.teams) {
    const tr = document.createElement('tr');

    const tdName = document.createElement('td');
    tdName.textContent = t.name;

    const tdMembers = document.createElement('td');
    tdMembers.textContent = (t.members || []).join(', ');

    const tdActions = document.createElement('td');
    tdActions.style.whiteSpace = 'nowrap';

    const editBtn = document.createElement('button');
    editBtn.className = 'btn btn-edit';
    editBtn.type = 'button';
    editBtn.textContent = 'Edit';
    editBtn.addEventListener('click', () => editTeamInline(t.id, tr));

    const delBtn = document.createElement('button');
    delBtn.className = 'btn btn-danger';
    delBtn.type = 'button';
    delBtn.textContent = 'Remove';
    delBtn.addEventListener('click', () => {
      if (!confirm(`Remove team "${t.name}"?`)) return;
      removeTeam(t.id);
    });

    tdActions.appendChild(editBtn);
    tdActions.appendChild(delBtn);

    tr.appendChild(tdName);
    tr.appendChild(tdMembers);
    tr.appendChild(tdActions);
    tbody.appendChild(tr);
  }

  if (!state.teams.length) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = 3;
    td.className = 'muted';
    td.textContent = 'No teams loaded yet.';
    tr.appendChild(td);
    tbody.appendChild(tr);
  }
}

function renderPointsTable() {
  const tbody = document.querySelector('#pointsTable tbody');
  tbody.innerHTML = '';

  const points = computePoints();

  const rows = state.teams
    .map((t) => ({ team: t, p: points[t.id] }))
    .sort((a, b) => b.p.total - a.p.total || a.team.name.localeCompare(b.team.name));

  for (const r of rows) {
    const tr = document.createElement('tr');

    const tdName = document.createElement('td');
    tdName.textContent = r.team.name;

    const td1 = document.createElement('td');
    td1.className = 'num';
    td1.textContent = String(r.p[1] || 0);

    const td2 = document.createElement('td');
    td2.className = 'num';
    td2.textContent = String(r.p[2] || 0);

    const td3 = document.createElement('td');
    td3.className = 'num';
    td3.textContent = String(r.p[3] || 0);

    const td4 = document.createElement('td');
    td4.className = 'num';
    td4.textContent = String(r.p[4] || 0);

    const td5 = document.createElement('td');
    td5.className = 'num';
    td5.textContent = String(r.p[5] || 0);

    const tdTotal = document.createElement('td');
    tdTotal.className = 'num';
    tdTotal.textContent = String(r.p.total || 0);

    tr.appendChild(tdName);
    tr.appendChild(td1);
    tr.appendChild(td2);
    tr.appendChild(td3);
    tr.appendChild(td4);
    tr.appendChild(td5);
    tr.appendChild(tdTotal);

    tbody.appendChild(tr);
  }

  if (!state.teams.length) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = 7;
    td.className = 'muted';
    td.textContent = 'Points will appear after teams are loaded.';
    tr.appendChild(td);
    tbody.appendChild(tr);
  }
}

function editTeamInline(teamId, tr) {
  const team = state.teams.find((t) => t.id === teamId);
  if (!team) return;

  tr.innerHTML = '';

  // Name input
  const tdName = document.createElement('td');
  const nameInput = document.createElement('input');
  nameInput.type = 'text';
  nameInput.value = team.name;
  nameInput.placeholder = 'Team name';
  nameInput.style.width = '100%';
  tdName.appendChild(nameInput);

  // Members inputs
  const tdMembers = document.createElement('td');
  const memberInputs = [];
  for (let i = 0; i < 4; i++) {
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.value = (team.members && team.members[i]) || '';
    inp.placeholder = `Member ${i + 1}`;
    inp.style.width = '100%';
    inp.style.marginBottom = i < 3 ? '4px' : '0';
    memberInputs.push(inp);
    tdMembers.appendChild(inp);
  }

  // Action buttons
  const tdActions = document.createElement('td');
  tdActions.style.whiteSpace = 'nowrap';

  const saveBtn = document.createElement('button');
  saveBtn.className = 'btn';
  saveBtn.type = 'button';
  saveBtn.textContent = 'Save';
  saveBtn.addEventListener('click', () => {
    const newName = nameInput.value.trim();
    if (!newName) { alert('Team name cannot be empty.'); return; }
    team.name = newName;
    team.members = memberInputs.map((inp) => inp.value.trim()).filter(Boolean);
    saveState();
    renderAll();
  });

  const cancelBtn = document.createElement('button');
  cancelBtn.className = 'btn btn-secondary';
  cancelBtn.type = 'button';
  cancelBtn.textContent = 'Cancel';
  cancelBtn.addEventListener('click', () => renderTeamsTable());

  tdActions.appendChild(saveBtn);
  tdActions.appendChild(cancelBtn);

  tr.appendChild(tdName);
  tr.appendChild(tdMembers);
  tr.appendChild(tdActions);

  nameInput.focus();
}

function removeTeam(teamId) {
  state.teams = state.teams.filter((t) => t.id !== teamId);

  for (const eventId of [1, 2, 3, 5]) {
    state.events[eventId].matches = (state.events[eventId].matches || []).filter(
      (m) => m.aTeamId !== teamId && m.bTeamId !== teamId
    );
  }
  delete state.events[4].times[teamId];

  saveState();
  renderAll();
}

function addTeamDialog() {
  const name = prompt('Team name?');
  if (!name || !name.trim()) return;
  const membersRaw = prompt('Enter 4 team members (comma-separated). Optional.');
  const members = String(membersRaw || '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean)
    .slice(0, 4);

  state.teams = ensureUniqueTeamNames([...state.teams, { id: uid('team'), name: name.trim(), members }]);
  saveState();
  renderAll();
}

function teamSelectOptions(selectedId = '') {
  const frag = document.createDocumentFragment();
  const opt0 = document.createElement('option');
  opt0.value = '';
  opt0.textContent = '— Select —';
  frag.appendChild(opt0);

  for (const t of state.teams.slice().sort((a, b) => a.name.localeCompare(b.name))) {
    const opt = document.createElement('option');
    opt.value = t.id;
    opt.textContent = t.name;
    if (t.id === selectedId) opt.selected = true;
    frag.appendChild(opt);
  }

  return frag;
}

function renderEventPanels() {
  const container = document.getElementById('eventPanels');
  container.innerHTML = '';

  container.appendChild(renderMatchesPanel(1));
  container.appendChild(renderMatchesPanel(2));
  container.appendChild(renderMatchesPanel(3));
  container.appendChild(renderTimedPanel(4));
  container.appendChild(renderMatchesPanel(5));

  setActiveEventPanel(getActiveEventId());
}

function getActiveEventId() {
  const active = document.querySelector('.tab.is-active');
  return active ? Number(active.dataset.event) : 1;
}

function setActiveEventPanel(eventId) {
  for (const el of document.querySelectorAll('.panel')) {
    el.style.display = Number(el.dataset.event) === eventId ? 'block' : 'none';
  }
}

function renderMatchesPanel(eventId) {
  const panel = document.createElement('div');
  panel.className = 'panel';
  panel.dataset.event = String(eventId);

  const eventNames = { 1: 'Bobsled', 2: 'Ice hockey', 3: 'Curling', 5: 'Skijump' };
  const title = document.createElement('h3');
  title.textContent = `Event ${eventId}: ${eventNames[eventId] || 'Match event'}`;
  panel.appendChild(title);

  const hint = document.createElement('p');
  hint.className = 'muted small';
  hint.innerHTML = 'Add matches between two teams. Winner gets <span class="kbd">3</span> points.';
  panel.appendChild(hint);

  const addBtn = document.createElement('button');
  addBtn.className = 'btn';
  addBtn.type = 'button';
  addBtn.textContent = 'Add match';
  addBtn.addEventListener('click', () => {
    if (state.teams.length < 2) {
      alert('Need at least 2 teams to create a match.');
      return;
    }
    state.events[eventId].matches.push({ id: uid(`match${eventId}`), aTeamId: '', bTeamId: '', winnerTeamId: '' });
    saveState();
    renderAll();
  });
  panel.appendChild(addBtn);

  const matches = state.events[eventId].matches || [];
  if (!matches.length) {
    const empty = document.createElement('p');
    empty.className = 'muted';
    empty.textContent = 'No matches yet.';
    panel.appendChild(empty);
    return panel;
  }

  const list = document.createElement('div');
  list.className = 'split';

  matches.forEach((m, idx) => {
    const card = document.createElement('div');
    card.className = 'card';

    const head = document.createElement('div');
    head.className = 'card-title';
    const h = document.createElement('h3');
    h.textContent = `Match ${idx + 1}`;
    const del = document.createElement('button');
    del.className = 'btn btn-danger';
    del.type = 'button';
    del.textContent = 'Remove';
    del.addEventListener('click', () => {
      state.events[eventId].matches = (state.events[eventId].matches || []).filter((x) => x.id !== m.id);
      saveState();
      renderAll();
    });
    head.appendChild(h);
    head.appendChild(del);
    card.appendChild(head);

    const row1 = document.createElement('div');
    row1.className = 'row';

    const selA = document.createElement('select');
    selA.className = 'select';
    selA.appendChild(teamSelectOptions(m.aTeamId || ''));
    selA.addEventListener('change', () => {
      m.aTeamId = selA.value;
      if (m.winnerTeamId && m.winnerTeamId !== m.aTeamId && m.winnerTeamId !== m.bTeamId) m.winnerTeamId = '';
      saveState();
      renderAll();
    });

    const selB = document.createElement('select');
    selB.className = 'select';
    selB.appendChild(teamSelectOptions(m.bTeamId || ''));
    selB.addEventListener('change', () => {
      m.bTeamId = selB.value;
      if (m.winnerTeamId && m.winnerTeamId !== m.aTeamId && m.winnerTeamId !== m.bTeamId) m.winnerTeamId = '';
      saveState();
      renderAll();
    });

    row1.appendChild(selA);
    row1.appendChild(selB);
    card.appendChild(row1);

    const row2 = document.createElement('div');
    row2.className = 'row';

    const winner = document.createElement('select');
    winner.className = 'select';
    const w0 = document.createElement('option');
    w0.value = '';
    w0.textContent = 'Winner —';
    winner.appendChild(w0);

    const aOk = m.aTeamId && m.aTeamId !== m.bTeamId;
    const bOk = m.bTeamId && m.bTeamId !== m.aTeamId;
    if (aOk) {
      const opt = document.createElement('option');
      opt.value = m.aTeamId;
      opt.textContent = getTeamsById().get(m.aTeamId)?.name || 'Team A';
      winner.appendChild(opt);
    }
    if (bOk) {
      const opt = document.createElement('option');
      opt.value = m.bTeamId;
      opt.textContent = getTeamsById().get(m.bTeamId)?.name || 'Team B';
      winner.appendChild(opt);
    }

    winner.value = m.winnerTeamId || '';
    winner.addEventListener('change', () => {
      m.winnerTeamId = winner.value;
      saveState();
      renderAll();
    });

    row2.appendChild(winner);
    card.appendChild(row2);

    list.appendChild(card);
  });

  panel.appendChild(list);
  return panel;
}

function renderTimedPanel(eventId) {
  const panel = document.createElement('div');
  panel.className = 'panel';
  panel.dataset.event = String(eventId);

  const eventNames = { 4: 'Biathlon' };
  const title = document.createElement('h3');
  title.textContent = `Event ${eventId}: ${eventNames[eventId] || 'Timed event'}`;
  panel.appendChild(title);

  const hint = document.createElement('p');
  hint.className = 'muted small';
  hint.textContent = 'Enter time in seconds (lower is better). Points: 1st=5, 2nd=3, 3rd=2, 4th=1 (top 4 score).';
  panel.appendChild(hint);

  if (!state.teams.length) {
    const empty = document.createElement('p');
    empty.className = 'muted';
    empty.textContent = 'Load teams first.';
    panel.appendChild(empty);
    return panel;
  }

  const wrap = document.createElement('div');
  wrap.className = 'table-wrap';

  const table = document.createElement('table');
  table.className = 'table';
  table.style.minWidth = '520px';

  const thead = document.createElement('thead');
  thead.innerHTML = '<tr><th>Team</th><th class="num">Time (seconds)</th></tr>';
  table.appendChild(thead);

  const tbody = document.createElement('tbody');

  const times = state.events[eventId].times || {};
  const sortedTeams = state.teams.slice().sort((a, b) => a.name.localeCompare(b.name));

  for (const t of sortedTeams) {
    const tr = document.createElement('tr');

    const tdName = document.createElement('td');
    tdName.textContent = t.name;

    const tdTime = document.createElement('td');
    tdTime.className = 'num';

    const inp = document.createElement('input');
    inp.type = 'number';
    inp.min = '0';
    inp.step = '0.01';
    inp.value = Number.isFinite(times[t.id]) ? String(times[t.id]) : '';
    inp.addEventListener('input', () => {
      const n = parseMaybeNumber(inp.value);
      if (n === null) {
        delete state.events[eventId].times[t.id];
      } else {
        state.events[eventId].times[t.id] = n;
      }
      saveState();
      renderPointsTable();
    });

    tdTime.appendChild(inp);

    tr.appendChild(tdName);
    tr.appendChild(tdTime);

    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  wrap.appendChild(table);
  panel.appendChild(wrap);

  const clearBtn = document.createElement('button');
  clearBtn.className = 'btn btn-secondary';
  clearBtn.type = 'button';
  clearBtn.textContent = 'Clear times';
  clearBtn.addEventListener('click', () => {
    if (!confirm('Clear all times for this event?')) return;
    state.events[eventId].times = {};
    saveState();
    renderAll();
  });
  panel.appendChild(document.createElement('div')).className = 'row';
  panel.lastChild.appendChild(clearBtn);

  return panel;
}

function renderAll() {
  renderTeamsTable();
  renderPointsTable();
  renderEventPanels();
}

function bindUI() {
  document.getElementById('exportXlsxBtn').addEventListener('click', () => {
    try {
      exportResultsToXlsx();
    } catch (err) {
      console.error(err);
      alert('Failed to export Excel file. Try Save (.csv) instead.');
    }
  });

  document.getElementById('exportCsvBtn').addEventListener('click', () => {
    try {
      exportResultsToCsv();
    } catch (err) {
      console.error(err);
      alert('Failed to export CSV file.');
    }
  });

  const importInput = document.getElementById('importResultsInput');
  document.getElementById('importResultsBtn').addEventListener('click', () => {
    importInput.value = '';
    importInput.click();
  });

  importInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      setStatus('Loading saved results…');
      const imported = await importSavedResultsFile(file);
      state = imported;
      saveState();
      renderAll();
      setActiveView(getActiveViewId());
      setStatus(`Loaded saved results from "${file.name}".`);
    } catch (err) {
      console.error(err);
      setStatus('Could not load saved results. Please pick a file created by Save (.xlsx) or Save (.csv).', 'error');
    }
  });

  document.getElementById('resetBtn').addEventListener('click', () => {
    if (!confirm('Reset everything (teams + results)?')) return;
    state = defaultState();
    saveState();
    setStatus('Reset complete.');
    setActiveView('teams');
    renderAll();
  });

  document.getElementById('createMatchupsBtn').addEventListener('click', () => {
    if (!confirm('Are you sure you want to create random matchups for Events 1, 2, 3 and 5? Existing matchups in those events will be replaced.')) {
      return;
    }
    createRandomMatchupsForEvents();
  });

  document.getElementById('addTeamBtn').addEventListener('click', addTeamDialog);

  document.getElementById('loadSampleBtn').addEventListener('click', async () => {
    try {
      const defaultTeamsFile = 'I&IAW_Feb_2026(Team list).csv';
      setStatus('Loading default teams file…');
      const response = await fetch(encodeURI(defaultTeamsFile));
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const text = await response.text();
      const rows = parseCsv(text);

      const { teams, warnings } = importTeamsFromRows(rows);
      state.teams = teams;
      state.events = defaultState().events;

      saveState();
      renderAll();
      const warningText = warnings.length ? ` (${warnings.join(' ')})` : '';
      setStatus(`Loaded default teams from "${defaultTeamsFile}".${warningText}`);
    } catch (err) {
      console.error(err);
      setStatus('Could not load default teams file.', 'error');
    }
  });

  document.getElementById('fileInput').addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const name = String(file.name || '').toLowerCase();
      const useCsv = isCsvFile(file) || name.endsWith('.csv');
      let rows = [];

      if (useCsv) {
        setStatus('Reading CSV file…');
        const text = await file.text();
        rows = parseCsv(text);
      } else {
        if (typeof XLSX === 'undefined') {
          throw new Error('XLSX library not available');
        }
        setStatus('Reading Excel file…');
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: 'array' });
        const firstSheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[firstSheetName];
        rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      }

      const { teams, warnings } = importTeamsFromRows(rows);
      state.teams = teams;

      // Reset events on import to avoid mismatched team IDs
      state.events = defaultState().events;

      saveState();
      renderAll();

      const msg = `Imported ${teams.length} team(s) from "${file.name}".`;
      setStatus(warnings.length ? `${msg} Warnings: ${warnings.join(' ')}` : msg);
    } catch (err) {
      console.error(err);
      setStatus(
        'Failed to import. If Excel import is blocked/offline, export your sheet as CSV and import the .csv instead.',
        'error'
      );
    }
  });

  document.getElementById('eventTabs').addEventListener('click', (e) => {
    const btn = e.target.closest('button.tab');
    if (!btn) return;
    for (const t of document.querySelectorAll('.tab')) t.classList.remove('is-active');
    btn.classList.add('is-active');
    setActiveEventPanel(Number(btn.dataset.event));
  });

  document.getElementById('viewTabs').addEventListener('click', (e) => {
    const btn = e.target.closest('button.tab');
    if (!btn) return;
    const viewId = btn.dataset.view;
    setActiveView(viewId);
  });

  const scoringMode = document.getElementById('scoringMode');
  scoringMode.value = state.settings.scoringMode || 'default';
  scoringMode.addEventListener('change', () => {
    state.settings.scoringMode = scoringMode.value;
    saveState();
    // (Reserved for future: alternative scoring.)
  });
}

function main() {
  bindUI();
  renderAll();
  setActiveView(getActiveViewId());
  setStatus('Ready. Upload an Excel (.xlsx/.xls) or CSV (.csv) file to begin.');
}

main();

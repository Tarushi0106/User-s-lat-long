/* ── All logic runs in the browser. No backend server required. ── */

let masterSites = [];   // merged from PAN India + Site Lat/Long
let currentRows = [];

// ── Boot ───────────────────────────────────────────────────────────────────
document.getElementById('headerDate').textContent = new Date().toLocaleDateString('en-IN', {
  weekday: 'short', day: '2-digit', month: 'short', year: 'numeric'
});
document.getElementById('footerYear').textContent = new Date().getFullYear();

// ── Toast ──────────────────────────────────────────────────────────────────
function toast(msg, ms = 3500) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), ms);
}

// ── Haversine distance (metres) ────────────────────────────────────────────
function haversineMeters(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// ── Column index finder ────────────────────────────────────────────────────
function makeCi(headers) {
  return (...names) => {
    for (const n of names) {
      const i = headers.findIndex(h => h === n.toLowerCase());
      if (i !== -1) return i;
    }
    for (const n of names) {
      const i = headers.findIndex(h => h.includes(n.toLowerCase()));
      if (i !== -1) return i;
    }
    return -1;
  };
}

// ── Parse PAN India Dashboard "Site master" sheet ──────────────────────────
// Rows 0–1 are dept/sub-dept labels; row 2 has actual column names; data from row 3.
function parsePanIndiaMaster(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (rows.length < 4) return {};

  const headers = rows[2].map(h => String(h || '').trim().toLowerCase().replace(/\r?\n/g, ' '));
  const ci = makeCi(headers);

  const iId   = ci('stpl site id', 'site id');
  const iName = ci('site name');
  const iCirc = ci('circle name', 'circle');
  const iDist = ci('district');
  const iPers = ci('site acquisition person name');
  const iLat  = ci('latitude');
  const iLng  = ci('longitude');
  const iStat = ci('tower status');

  const map = {};
  for (let i = 3; i < rows.length; i++) {
    const r  = rows[i];
    const id = String(r[iId] || '').trim();
    if (!id) continue;
    const lat = parseFloat(r[iLat]);
    const lng = parseFloat(r[iLng]);
    map[id] = {
      stsId:  id,
      name:   String(r[iName] || '').trim(),
      circle: String(r[iCirc] || '').trim(),
      dist:   String(r[iDist] || '').trim(),
      person: String(r[iPers] || '').trim(),
      status: String(r[iStat] || '').trim(),
      lat:    (!isNaN(lat) && lat) ? lat : null,
      lng:    (!isNaN(lng) && lng) ? lng : null,
    };
  }
  return map;
}

// ── Parse Site Lat/Long file ───────────────────────────────────────────────
// Tries sheet "DPR" first, then "Site master", then first sheet.
function parseSiteLatLong(wb) {
  const sheetPriority = ['DPR', 'Site master', wb.SheetNames[0]];
  let ws = null;
  for (const name of sheetPriority) {
    if (wb.Sheets[name]) { ws = wb.Sheets[name]; break; }
  }
  if (!ws) return {};

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (rows.length < 2) return {};

  const headers = rows[0].map(h => String(h || '').trim().toLowerCase().replace(/\r?\n/g, ' '));
  const ci = makeCi(headers);

  const iId   = ci('sts site id', 'stpl site id', 'site id');
  const iName = ci('site name');
  const iCirc = ci('circle name', 'circle');
  const iDist = ci('district');
  const iPers = ci('site acquisition person name');
  const iLat  = ci('lat', 'latitude');
  const iLng  = ci('long', 'longitude');
  const iStat = ci('updated status', 'tower status');

  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const r   = rows[i];
    const id  = String(r[iId] || '').trim();
    const lat = parseFloat(r[iLat]);
    const lng = parseFloat(r[iLng]);
    if (!id || isNaN(lat) || isNaN(lng) || !lat || !lng) continue;
    map[id] = {
      stsId:  id,
      name:   String(r[iName] || '').trim(),
      circle: String(r[iCirc] || '').trim(),
      dist:   String(r[iDist] || '').trim(),
      person: String(r[iPers] || '').trim(),
      status: String(r[iStat] || '').trim(),
      lat, lng,
    };
  }
  return map;
}

// ── Merge both maps ────────────────────────────────────────────────────────
// Site Lat/Long coordinates override PAN India coordinates when both have the same ID.
// All sites from both files are included. Only sites with valid GPS are kept.
function mergeMasters(panMap, llMap) {
  const merged = { ...panMap };

  for (const [id, site] of Object.entries(llMap)) {
    if (merged[id]) {
      merged[id].lat = site.lat;
      merged[id].lng = site.lng;
      if (!merged[id].name   && site.name)   merged[id].name   = site.name;
      if (!merged[id].circle && site.circle) merged[id].circle = site.circle;
      if (!merged[id].dist   && site.dist)   merged[id].dist   = site.dist;
      if (!merged[id].person && site.person) merged[id].person = site.person;
    } else {
      merged[id] = { ...site };
    }
  }

  return Object.values(merged).filter(s => s.lat && s.lng && !isNaN(s.lat) && !isNaN(s.lng));
}

// ── Build fast lookup index ────────────────────────────────────────────────
function buildMasterIndex() {
  const byId = new Map();
  for (const site of masterSites) {
    if (site.stsId) byId.set(site.stsId.trim().toUpperCase(), site);
  }
  return { byId, list: masterSites };
}

// ── Update UI master status ────────────────────────────────────────────────
function updateMasterStatus() {
  const pill     = document.getElementById('headerMasterPill');
  const statEl   = document.getElementById('statMasterSites');
  const bannerEl = document.getElementById('masterBanner');

  statEl.textContent = masterSites.length;

  if (masterSites.length) {
    pill.textContent = `✔ ${masterSites.length} Sites Loaded`;
    pill.className   = 'master-pill loaded';
    if (bannerEl) bannerEl.style.display = 'none';
  } else {
    pill.textContent = '⚠ Master Loading…';
    pill.className   = 'master-pill none';
    if (bannerEl) bannerEl.style.display = '';
  }
}

// ── Auto-load both master files on startup ─────────────────────────────────
async function autoLoadMasters() {
  // Try cache first for instant startup
  const cached = localStorage.getItem('masterSites');
  if (cached) {
    try {
      masterSites = JSON.parse(cached);
      updateMasterStatus();
      loadStats();
    } catch (_) { /* ignore bad cache */ }
  }

  // Always re-fetch fresh copies in the background
  try {
    const [panBuf, llBuf] = await Promise.all([
      fetch('pan-india-master.xlsb').then(r => { if (!r.ok) throw new Error('pan-india-master.xlsb not found'); return r.arrayBuffer(); }),
      fetch('site-latlong-master.xlsx').then(r => { if (!r.ok) throw new Error('site-latlong-master.xlsx not found'); return r.arrayBuffer(); }),
    ]);

    const panWb = XLSX.read(new Uint8Array(panBuf), { type: 'array' });
    const llWb  = XLSX.read(new Uint8Array(llBuf),  { type: 'array' });

    const panSheet = panWb.Sheets['Site master'];
    if (!panSheet) throw new Error('PAN India Dashboard: "Site master" sheet not found');

    const panMap = parsePanIndiaMaster(panSheet);
    const llMap  = parseSiteLatLong(llWb);

    masterSites = mergeMasters(panMap, llMap);
    localStorage.setItem('masterSites', JSON.stringify(masterSites));

    updateMasterStatus();
    loadStats();
    toast(`✔ Master loaded — ${masterSites.length} sites ready`);
  } catch (err) {
    console.error('Master load error:', err);
    if (!masterSites.length) {
      const pill = document.getElementById('headerMasterPill');
      pill.textContent = '✘ Master Load Failed';
      pill.className   = 'master-pill none';
      toast('⚠ Could not load master files: ' + err.message, 6000);
    }
  }
}

// ── Stats ──────────────────────────────────────────────────────────────────
function loadStats() {
  const reports        = getReports();
  const totalMatched   = reports.reduce((s, r) => s + r.matchedCount,   0);
  const totalUnmatched = reports.reduce((s, r) => s + r.unmatchedCount, 0);
  document.getElementById('statMasterSites').textContent = masterSites.length;
  document.getElementById('statUploads').textContent     = reports.length;
  document.getElementById('statRows').textContent        = totalMatched + totalUnmatched;
  document.getElementById('statMatched').textContent     = totalMatched;
  document.getElementById('statUnmatched').textContent   = totalUnmatched;
  renderRecentReports(reports.slice(0, 5));
}

// ── Reports ────────────────────────────────────────────────────────────────
function getReports() {
  return JSON.parse(localStorage.getItem('reports') || '[]');
}
function saveReport(report) {
  const reports = getReports();
  reports.unshift(report);
  if (reports.length > 20) reports.length = 20;
  localStorage.setItem('reports', JSON.stringify(reports));
}

function renderRecentReports(reports) {
  const el = document.getElementById('recentReports');
  if (!reports.length) {
    el.innerHTML = '<div class="empty-msg">No uploads yet. Upload a field person\'s GPS tracker file to begin.</div>';
    return;
  }
  el.innerHTML = reports.map(r => `
    <div class="report-list-item">
      <div onclick="viewReport('${r.id}')" style="flex:1;cursor:pointer;">
        <div class="report-name">${esc(r.fileName)}</div>
        <div class="report-meta">By ${esc(r.uploadedBy)} &nbsp;·&nbsp; ${r.createdAt} &nbsp;·&nbsp; ${r.totalRows} rows processed</div>
      </div>
      <div class="report-badges">
        <span class="badge badge-green">✔ ${r.matchedCount} Work Done, Verified</span>
        <span class="badge badge-red">✘ ${r.unmatchedCount} Not Verified</span>
      </div>
      <button class="btn-delete" onclick="deleteReport(event,'${r.id}')" title="Delete">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/></svg>
      </button>
    </div>
  `).join('');
}

function deleteReport(e, id) {
  e.stopPropagation();
  if (!confirm('Delete this report?')) return;
  const reports = getReports().filter(r => r.id !== id);
  localStorage.setItem('reports', JSON.stringify(reports));
  loadStats();
  document.getElementById('reportDetail').style.display = 'none';
  toast('Report deleted.');
}

function viewReport(id) {
  const report = getReports().find(r => r.id === id);
  if (!report) return;
  currentRows = report.rows;
  document.getElementById('reportTitle').textContent = report.fileName + ' — ' + report.createdAt;
  document.getElementById('reportDetail').style.display = '';
  document.getElementById('searchInput').value  = '';
  document.getElementById('statusFilter').value = '';
  renderTable(currentRows);
  setTimeout(() => document.getElementById('reportDetail').scrollIntoView({ behavior: 'smooth' }), 80);
}

// ── File input label ───────────────────────────────────────────────────────
document.getElementById('fileInput').addEventListener('change', function () {
  document.getElementById('fileName').textContent = this.files[0]?.name || 'Choose file…';
});

// ── Upload & Match ─────────────────────────────────────────────────────────
// Each row in the user's file = one GPS ping by that person.
// Look up the site (by Site ID then by nearest GPS) in the merged master.
// Compare the person's GPS against the master site GPS — ≤ 500 m = Verified.
document.getElementById('uploadForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const file = document.getElementById('fileInput').files[0];
  if (!file) return;

  if (!masterSites.length) {
    document.getElementById('uploadMsg').textContent = 'Master data is still loading — please wait a moment and try again.';
    document.getElementById('uploadMsg').className = 'error';
    return;
  }

  const btn   = document.getElementById('uploadBtn');
  const msgEl = document.getElementById('uploadMsg');
  btn.disabled = true;
  btn.textContent = 'Processing…';
  msgEl.textContent = '';
  msgEl.className = '';

  try {
    const buf  = await file.arrayBuffer();
    const wb   = XLSX.read(new Uint8Array(buf), { type: 'array' });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    if (rows.length < 2) throw new Error('File has no data rows');

    const headers = rows[0].map(h => String(h || '').trim().toLowerCase());
    const ci = makeCi(headers);

    const colLat     = ci('lat', 'latitude');
    const colLng     = ci('lng', 'long', 'longitude');
    const colSiteId  = ci('site id', 'stpl site id', 'sts site id', 'site_id', 'siteid');
    const colSiteNm  = ci('site name', 'sitename', 'site_name');
    const colTracker = ci('tracker_id', 'tracker', 'device');
    const colTime    = ci('time (gmt', 'time', 'timestamp', 'date');
    const colPerson  = ci('person', 'name', 'employee', 'user', 'field');

    if (colLat === -1 || colLng === -1) throw new Error('Lat/Lng columns not found in the uploaded file');

    // Determine person name
    let personName = document.getElementById('uploadedBy').value.trim() || '';
    if (!personName) {
      for (let i = 1; i < rows.length; i++) {
        if (colPerson !== -1 && rows[i][colPerson]) {
          personName = String(rows[i][colPerson]).trim(); break;
        }
        if (colTracker !== -1 && rows[i][colTracker]) {
          const raw = String(rows[i][colTracker]).trim();
          personName = raw.includes('@') ? raw.split('@')[1] : raw; break;
        }
      }
    }
    if (!personName) personName = 'Unknown';

    const { byId, list: masterList } = buildMasterIndex();
    const TOLERANCE = 500; // metres

    const resultRows   = [];
    let matchedCount   = 0;
    let unmatchedCount = 0;
    let skippedCount   = 0;
    let rowNumber      = 0;

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];

      // Person's actual GPS for this row
      const userLat = parseFloat(r[colLat]);
      const userLng = parseFloat(r[colLng]);
      if (!userLat || !userLng || isNaN(userLat) || isNaN(userLng)) {
        skippedCount++;
        continue;
      }

      const rawSiteId = colSiteId !== -1 ? String(r[colSiteId] || '').trim() : '';
      const rawSiteNm = colSiteNm !== -1 ? String(r[colSiteNm] || '').trim() : '';
      const timeOfRow = colTime   !== -1 ? String(r[colTime]   || '') : '';

      // Step 1: Look up site by Site ID in master
      let masterSite   = rawSiteId ? byId.get(rawSiteId.toUpperCase()) : null;
      let lookupMethod = 'Site ID';

      // Step 2: No ID match — find nearest master site by GPS proximity
      if (!masterSite) {
        lookupMethod = 'GPS';
        let bestDist = Infinity;
        const degTol = TOLERANCE / 111000;
        for (const site of masterList) {
          if (Math.abs(userLat - site.lat) > degTol) continue;
          if (Math.abs(userLng - site.lng) > degTol) continue;
          const d = haversineMeters(userLat, userLng, site.lat, site.lng);
          if (d < bestDist) { bestDist = d; masterSite = site; }
        }
      }

      rowNumber++;

      if (!masterSite) {
        // Location not found in either master file
        unmatchedCount++;
        resultRows.push({
          rowNumber,
          personName,
          timeOfVisit:     timeOfRow,
          userLat, userLng,
          userSiteId:      rawSiteId,
          userSiteName:    rawSiteNm,
          matchedSiteId:   rawSiteId || '–',
          matchedSiteName: 'Not Found in Master',
          district: '–', circle: '–',
          masterLat: null, masterLng: null,
          distanceMeters: null,
          matched: false, lookupMethod: '–',
          status: 'Not Matched',
        });
        continue;
      }

      // Step 3: Compare person's GPS vs master site GPS
      const dist    = haversineMeters(userLat, userLng, masterSite.lat, masterSite.lng);
      const matched = dist <= TOLERANCE;

      if (matched) matchedCount++;
      else         unmatchedCount++;

      resultRows.push({
        rowNumber,
        personName,
        timeOfVisit:     timeOfRow,
        userLat, userLng,
        userSiteId:      rawSiteId,
        userSiteName:    rawSiteNm,
        matchedSiteId:   masterSite.stsId,
        matchedSiteName: masterSite.name,
        district:        masterSite.dist,
        circle:          masterSite.circle,
        masterLat:       masterSite.lat,
        masterLng:       masterSite.lng,
        distanceMeters:  Math.round(dist),
        matched, lookupMethod,
        status: matched ? 'Work Done - Verified' : 'Not Matched',
      });
    }

    if (!rowNumber && skippedCount > 0)
      throw new Error('No valid GPS coordinates found — all rows had 0,0 or blank coordinates');
    if (!rowNumber)
      throw new Error('No data rows found in the uploaded file');

    const report = {
      id:            Date.now().toString(),
      fileName:      file.name,
      uploadedBy:    personName,
      createdAt:     new Date().toLocaleDateString('en-IN', {
        day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit'
      }),
      matchedCount,
      unmatchedCount,
      totalRows:     rowNumber,
      rows:          resultRows,
    };

    saveReport(report);
    loadStats();

    msgEl.textContent = `Done — ${rowNumber} rows processed · ${matchedCount} verified (within 500 m of master site) · ${unmatchedCount} not verified${skippedCount ? ` · ${skippedCount} rows skipped (no GPS)` : ''}.`;
    msgEl.className = 'success';
    viewReport(report.id);
  } catch (err) {
    document.getElementById('uploadMsg').textContent = 'Error: ' + err.message;
    document.getElementById('uploadMsg').className   = 'error';
  }

  btn.disabled = false;
  btn.textContent = 'Upload & Match';
});

// ── Detail table ───────────────────────────────────────────────────────────
document.getElementById('closeDetail').addEventListener('click', () => {
  document.getElementById('reportDetail').style.display = 'none';
});
document.getElementById('searchInput').addEventListener('input', filterTable);
document.getElementById('statusFilter').addEventListener('change', filterTable);

function filterTable() {
  const q  = document.getElementById('searchInput').value.toLowerCase();
  const st = document.getElementById('statusFilter').value;
  const filtered = currentRows.filter(r => {
    const hay = `${r.personName} ${r.matchedSiteId} ${r.matchedSiteName} ${r.district} ${r.circle}`.toLowerCase();
    return (!q || hay.includes(q)) && (!st || r.status === st);
  });
  document.getElementById('resultCount').textContent = filtered.length + ' of ' + currentRows.length + ' rows';
  renderTable(filtered);
}

function renderTable(rows) {
  document.getElementById('resultCount').textContent = rows.length + ' rows';
  const tbody = document.getElementById('reportBody');
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:28px;color:var(--muted)">No results.</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map(r => {
    const userCoords   = `${r.userLat?.toFixed(5) ?? '–'}, ${r.userLng?.toFixed(5) ?? '–'}`;
    const masterCoords = r.masterLat != null
      ? `${r.masterLat.toFixed(5)}, ${r.masterLng.toFixed(5)}`
      : '–';
    const distCell = r.distanceMeters != null
      ? `<span class="dist-tag${r.matched ? '' : ' dist-far'}">${r.distanceMeters} m</span>`
      : '–';
    return `
    <tr>
      <td style="color:var(--muted)">${r.rowNumber}</td>
      <td><strong>${esc(r.personName)}</strong></td>
      <td class="mono">${esc(r.matchedSiteId) || '–'}</td>
      <td>${esc(r.matchedSiteName) || '–'}</td>
      <td style="font-size:.82rem">${esc(r.district) || '–'}</td>
      <td style="font-size:.82rem;color:var(--muted)">${esc(r.circle) || '–'}</td>
      <td class="coords-cell">
        <span class="coord-user">${userCoords}</span>
        <span class="coord-sep">↕ master</span>
        <span class="coord-master">${masterCoords}</span>
      </td>
      <td>${distCell}</td>
      <td style="font-size:.82rem">${esc(r.timeOfVisit) || '–'}</td>
      <td>${r.matched
        ? `<span class="status-pill pill-green">✔ Work Done, Verified</span>`
        : `<span class="status-pill pill-red">✘ Not Verified</span>`
      }</td>
    </tr>`;
  }).join('');
}

// ── Helpers ────────────────────────────────────────────────────────────────
function esc(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

const style = document.createElement('style');
style.textContent = '@keyframes spin { to { transform: rotate(360deg); } }';
document.head.appendChild(style);

// ── Init ───────────────────────────────────────────────────────────────────
loadStats();
autoLoadMasters();

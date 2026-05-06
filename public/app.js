/* ── All logic runs in the browser. No backend server required. ── */

// Separate maps so we can check each master independently
let panIdMap  = new Map();   // Site ID → site object  (PAN India Dashboard)
let llIdMap   = new Map();   // Site ID → site object  (Site Lat/Long)
let masterSites = [];        // merged array (for stats / header count)
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
// Rows 0–1 are dept labels; row 2 = actual headers; data from row 3.
function parsePanIndia(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (rows.length < 4) return new Map();

  const headers = rows[2].map(h => String(h || '').trim().toLowerCase().replace(/\r?\n/g, ' '));
  const ci = makeCi(headers);

  const iId   = ci('stpl site id', 'site id');
  const iName = ci('site name');
  const iCirc = ci('circle name', 'circle');
  const iDist = ci('district');
  const iPers = ci('site acquisition person name');
  const iLat  = ci('latitude');
  const iLng  = ci('longitude');

  const map = new Map();
  for (let i = 3; i < rows.length; i++) {
    const r  = rows[i];
    const id = String(r[iId] || '').trim();
    if (!id) continue;
    const lat = parseFloat(r[iLat]);
    const lng = parseFloat(r[iLng]);
    map.set(id.toUpperCase(), {
      stsId:  id,
      name:   String(r[iName] || '').trim(),
      circle: String(r[iCirc] || '').trim(),
      dist:   String(r[iDist] || '').trim(),
      person: String(r[iPers] || '').trim(),
      lat:    (!isNaN(lat) && lat) ? lat : null,
      lng:    (!isNaN(lng) && lng) ? lng : null,
      source: 'PAN India',
    });
  }
  return map;
}

// ── Parse one sheet from Site Lat/Long file ────────────────────────────────
function parseLLSheet(ws, map) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (rows.length < 2) return;

  const headers = rows[0].map(h => String(h || '').trim().toLowerCase().replace(/\r?\n/g, ' '));
  const ci = makeCi(headers);

  // "STS site ID" (DPR / GUJ&MUM) or "STPL Site ID" (Site master)
  const iId   = ci('sts site id', 'stpl site id', 'site id');
  const iName = ci('site name');
  const iCirc = ci('circle name', 'circle');
  const iDist = ci('district');
  const iPers = ci('site acquisition person name');
  const iLat  = ci('lat', 'latitude');
  const iLng  = ci('long', 'longitude');

  if (iId === -1 || iLat === -1 || iLng === -1) return;

  for (let i = 1; i < rows.length; i++) {
    const r   = rows[i];
    const id  = String(r[iId] || '').trim();
    const lat = parseFloat(r[iLat]);
    const lng = parseFloat(r[iLng]);
    if (!id || isNaN(lat) || isNaN(lng) || !lat || !lng) continue;
    // Don't overwrite — first valid entry for an ID wins
    if (map.has(id.toUpperCase())) continue;
    map.set(id.toUpperCase(), {
      stsId:  id,
      name:   String(r[iName] || '').trim(),
      circle: String(r[iCirc] || '').trim(),
      dist:   String(r[iDist] || '').trim(),
      person: String(r[iPers] || '').trim(),
      lat, lng,
      source: 'Site Lat/Long',
    });
  }
}

// ── Parse Site Lat/Long file — reads ALL sheets ────────────────────────────
function parseSiteLatLong(wb) {
  const map = new Map();
  // Preferred order: Site master has STPL IDs, DPR + GUJ&MUM have STS IDs
  const order = ['Site master', 'DPR', 'GUJ&MUM', ...wb.SheetNames];
  const seen  = new Set();
  for (const name of order) {
    if (seen.has(name) || !wb.Sheets[name]) continue;
    seen.add(name);
    parseLLSheet(wb.Sheets[name], map);
  }
  return map;
}

// ── Build merged array for stats ───────────────────────────────────────────
function buildMergedArray() {
  const merged = new Map();
  for (const [k, v] of panIdMap)  merged.set(k, v);
  for (const [k, v] of llIdMap) {
    if (merged.has(k)) {
      // Site Lat/Long has authoritative coordinates — override
      const s = { ...merged.get(k), lat: v.lat, lng: v.lng };
      merged.set(k, s);
    } else {
      merged.set(k, v);
    }
  }
  return [...merged.values()].filter(s => s.lat && s.lng);
}

// ── Update header pill ─────────────────────────────────────────────────────
function updateMasterStatus() {
  const pill   = document.getElementById('headerMasterPill');
  const statEl = document.getElementById('statMasterSites');

  statEl.textContent = masterSites.length || '–';

  if (masterSites.length) {
    pill.textContent = `✔ ${masterSites.length} Sites Loaded`;
    pill.className   = 'master-pill loaded';
  } else {
    pill.textContent = '⚠ Master Loading…';
    pill.className   = 'master-pill none';
  }
}

// ── Auto-load both master files on startup ─────────────────────────────────
async function autoLoadMasters() {
  // Serve from cache instantly while fresh copy loads in background
  const cached = localStorage.getItem('masterCache');
  if (cached) {
    try {
      const { pan, ll } = JSON.parse(cached);
      panIdMap     = new Map(pan);
      llIdMap      = new Map(ll);
      masterSites  = buildMergedArray();
      updateMasterStatus();
      loadStats();
    } catch (_) { /* bad cache, will be overwritten */ }
  }

  try {
    const [panBuf, llBuf] = await Promise.all([
      fetch('pan-india-master.xlsb').then(r  => { if (!r.ok) throw new Error('pan-india-master.xlsb not found');   return r.arrayBuffer(); }),
      fetch('site-latlong-master.xlsx').then(r => { if (!r.ok) throw new Error('site-latlong-master.xlsx not found'); return r.arrayBuffer(); }),
    ]);

    const panWb = XLSX.read(new Uint8Array(panBuf), { type: 'array' });
    const llWb  = XLSX.read(new Uint8Array(llBuf),  { type: 'array' });

    const panSheet = panWb.Sheets['Site master'];
    if (!panSheet) throw new Error('PAN India Dashboard: "Site master" sheet not found');

    panIdMap    = parsePanIndia(panSheet);
    llIdMap     = parseSiteLatLong(llWb);
    masterSites = buildMergedArray();

    // Cache serialised maps for next load
    localStorage.setItem('masterCache', JSON.stringify({
      pan: [...panIdMap.entries()],
      ll:  [...llIdMap.entries()],
    }));

    updateMasterStatus();
    loadStats();
    toast(`✔ Master ready — ${panIdMap.size} PAN India · ${llIdMap.size} Site Lat/Long`);
  } catch (err) {
    console.error('Master load error:', err);
    if (!masterSites.length) {
      const pill = document.getElementById('headerMasterPill');
      pill.textContent = '✘ Master Load Failed';
      pill.className   = 'master-pill none';
      toast('⚠ ' + err.message, 6000);
    }
  }
}

// ── Stats ──────────────────────────────────────────────────────────────────
function loadStats() {
  const reports        = getReports();
  const totalMatched   = reports.reduce((s, r) => s + r.matchedCount,   0);
  const totalUnmatched = reports.reduce((s, r) => s + r.unmatchedCount, 0);
  document.getElementById('statMasterSites').textContent = masterSites.length || '–';
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
        <div class="report-meta">By ${esc(r.uploadedBy)} &nbsp;·&nbsp; ${r.createdAt} &nbsp;·&nbsp; ${r.totalRows} rows</div>
      </div>
      <div class="report-badges">
        <span class="badge badge-green">✔ ${r.matchedCount} Verified</span>
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
// For each row in the uploaded file:
//   1. Get the person's Lat/Long.
//   2. Find the nearest master site (from PAN India + Site Lat/Long) by GPS.
//   3. If nearest site is ≤ 500 m → Work Done, Verified.
//      If nearest site is > 500 m → Not Verified.
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
  btn.disabled    = true;
  btn.textContent = 'Processing…';
  msgEl.textContent = '';
  msgEl.className   = '';

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
    const colTime    = ci('time (gmt', 'time', 'timestamp', 'date');
    const colTracker = ci('tracker_id', 'tracker', 'device');
    const colPerson  = ci('person', 'name', 'employee', 'user', 'field');

    if (colLat === -1 || colLng === -1)
      throw new Error('Lat / Long columns not found in the uploaded file');

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

    const TOLERANCE  = 50; // metres
    const resultRows = [];
    let matchedCount   = 0;
    let unmatchedCount = 0;
    let skippedCount   = 0;
    let rowNumber      = 0;

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];

      const userLat = parseFloat(r[colLat]);
      const userLng = parseFloat(r[colLng]);
      if (!userLat || !userLng || isNaN(userLat) || isNaN(userLng)) {
        skippedCount++;
        continue;
      }

      const timeOfRow = colTime !== -1 ? String(r[colTime] || '') : '';
      rowNumber++;

      // Find the nearest master site by GPS
      let nearestSite = null;
      let nearestDist = Infinity;
      const degTol = TOLERANCE / 111000;

      for (const site of masterSites) {
        if (Math.abs(userLat - site.lat) > degTol) continue;
        if (Math.abs(userLng - site.lng) > degTol) continue;
        const d = haversineMeters(userLat, userLng, site.lat, site.lng);
        if (d < nearestDist) { nearestDist = d; nearestSite = site; }
      }

      const matched = nearestSite !== null && nearestDist <= TOLERANCE;
      if (matched) matchedCount++;
      else         unmatchedCount++;

      resultRows.push({
        rowNumber,
        personName,
        timeOfVisit:     timeOfRow,
        userLat, userLng,
        matchedSiteId:   nearestSite ? nearestSite.stsId  : '–',
        matchedSiteName: nearestSite ? nearestSite.name   : '–',
        district:        nearestSite ? nearestSite.dist   : '–',
        circle:          nearestSite ? nearestSite.circle : '–',
        masterSource:    nearestSite ? nearestSite.source : '–',
        masterLat:       nearestSite ? nearestSite.lat    : null,
        masterLng:       nearestSite ? nearestSite.lng    : null,
        distanceMeters:  nearestSite ? Math.round(nearestDist) : null,
        matched,
        status: matched ? 'Work Done - Verified' : 'Not Matched',
      });
    }

    if (!rowNumber && skippedCount > 0)
      throw new Error('No valid GPS rows found — all rows had 0,0 or blank coordinates');
    if (!rowNumber)
      throw new Error('No data rows found in the uploaded file');

    const report = {
      id:            Date.now().toString(),
      fileName:      file.name,
      uploadedBy:    personName,
      createdAt:     new Date().toLocaleDateString('en-IN', {
        day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit',
      }),
      matchedCount,
      unmatchedCount,
      totalRows:     rowNumber,
      rows:          resultRows,
    };

    saveReport(report);
    loadStats();

    let summary = `Done — ${rowNumber} rows · ${matchedCount} Verified · ${unmatchedCount} Not Verified`;
    if (skippedCount) summary += ` · ${skippedCount} rows skipped (no GPS)`;
    msgEl.textContent = summary;
    msgEl.className   = 'success';
    viewReport(report.id);
  } catch (err) {
    document.getElementById('uploadMsg').textContent = 'Error: ' + err.message;
    document.getElementById('uploadMsg').className   = 'error';
  }

  btn.disabled    = false;
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
    const statusCell = r.status === 'Work Done - Verified'
      ? `<span class="status-pill pill-green">✔ Work Done, Verified</span>`
      : r.status === 'Site Not in Master'
        ? `<span class="status-pill pill-orange">⚠ Site Not in Master</span>`
        : `<span class="status-pill pill-red">✘ Not Verified</span>`;
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
        <span class="coord-sep">↕ ${esc(r.masterSource || '–')}</span>
        <span class="coord-master">${masterCoords}</span>
      </td>
      <td>${distCell}</td>
      <td style="font-size:.82rem">${esc(r.timeOfVisit) || '–'}</td>
      <td>${statusCell}</td>
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

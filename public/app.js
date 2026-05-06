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
  const pill = document.getElementById('headerMasterPill');
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
  renderRecentReports(getReports().slice(0, 5));
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
        <span class="badge badge-green">✔ ${r.matchedCount} Sites Verified</span>
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
// 1. Read ALL GPS pings from the uploaded file into memory.
// 2. For each master site, find the single nearest ping.
// 3. If that nearest ping is ≤ 50 m → Work Done, Verified.
//    Otherwise → Not Verified.
// Result: one row per master site (not one row per GPS ping).
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

    const colTime    = ci('time (gmt', 'time', 'timestamp', 'date');
    const colTracker = ci('tracker_id', 'tracker', 'device');
    const colPerson  = ci('person', 'name', 'employee', 'user', 'field');

    // ── Auto-detect how lat/lng are stored ────────────────────────────────
    // Strategy 1: separate columns by header name
    let colLat = ci('lat', 'latitude');
    let colLng = ci('lng', 'long', 'longitude');

    // Strategy 2: single combined column — scan every column's values
    // looking for "number, number" or "number number" patterns
    let colCombined = -1;
    if (colLat === -1 || colLng === -1) {
      outer: for (let c = 0; c < headers.length; c++) {
        for (let r = 1; r < rows.length; r++) {
          const raw   = String(rows[r][c] || '').trim();
          const parts = raw.split(/[\s,\/]+/).map(Number).filter(n => !isNaN(n));
          if (parts.length >= 2) { colCombined = c; break outer; }
        }
      }
    }

    // Strategy 3: find two numeric columns whose values fall in India lat/lng range
    if (colLat === -1 && colLng === -1 && colCombined === -1) {
      const numCols = [];
      for (let c = 0; c < headers.length; c++) {
        const sample = rows.slice(1, 6).map(r => parseFloat(r[c])).filter(v => !isNaN(v));
        if (sample.length) numCols.push({ c, avg: sample.reduce((a,b) => a+b,0)/sample.length });
      }
      const latCol = numCols.find(({ avg }) => avg >= 6  && avg <= 38);
      const lngCol = numCols.find(({ avg }) => avg >= 60 && avg <= 100);
      if (latCol) colLat = latCol.c;
      if (lngCol) colLng = lngCol.c;
    }

    if (colLat === -1 && colLng === -1 && colCombined === -1)
      throw new Error('Could not find Lat/Long data. Make sure the file has GPS coordinates.');

    // Extract lat/lng from a single row
    function getLatLng(r) {
      if (colLat !== -1 && colLng !== -1) {
        return { lat: parseFloat(r[colLat]), lng: parseFloat(r[colLng]) };
      }
      const parts = String(r[colCombined] || '').trim().split(/[\s,\/]+/).map(Number);
      return { lat: parts[0], lng: parts[1] };
    }

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

    // ── Step 1: Collect ALL valid GPS pings from the uploaded file ─────────
    const pings = [];
    let skippedCount = 0;
    for (let i = 1; i < rows.length; i++) {
      const r          = rows[i];
      const { lat, lng } = getLatLng(r);
      if (!lat || !lng || isNaN(lat) || isNaN(lng)) { skippedCount++; continue; }
      pings.push({
        lat, lng,
        time: colTime !== -1 ? String(r[colTime] || '') : '',
      });
    }

    if (!pings.length) throw new Error('No valid GPS coordinates found in the uploaded file');

    // ── Step 2: For each master site, find the single nearest ping ─────────
    const TOLERANCE  = 50; // metres
    const resultRows = [];
    let matchedCount   = 0;

    for (const site of masterSites) {
      let nearestDist = Infinity;
      let nearestTime = '';
      let nearestLat  = null;
      let nearestLng  = null;
      const degTol = TOLERANCE / 111000;

      for (const ping of pings) {
        // Quick bounding-box pre-filter before expensive haversine
        if (Math.abs(ping.lat - site.lat) > degTol) continue;
        if (Math.abs(ping.lng - site.lng) > degTol) continue;
        const d = haversineMeters(ping.lat, ping.lng, site.lat, site.lng);
        if (d < nearestDist) {
          nearestDist = d;
          nearestTime = ping.time;
          nearestLat  = ping.lat;
          nearestLng  = ping.lng;
        }
      }

      // Skip sites the person never came within 50 m of
      if (nearestDist > TOLERANCE) continue;

      matchedCount++;
      resultRows.push({
        rowNumber: matchedCount,
        personName,
        timeOfVisit:     nearestTime,
        userLat:         nearestLat,
        userLng:         nearestLng,
        matchedSiteId:   site.stsId,
        matchedSiteName: site.name,
        district:        site.dist,
        circle:          site.circle,
        masterSource:    site.source,
        masterLat:       site.lat,
        masterLng:       site.lng,
        distanceMeters:  Math.round(nearestDist),
        matched:         true,
        status:          'Work Done - Verified',
      });
    }

    const report = {
      id:            Date.now().toString(),
      fileName:      file.name,
      uploadedBy:    personName,
      createdAt:     new Date().toLocaleDateString('en-IN', {
        day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit',
      }),
      matchedCount,
      totalRows:     matchedCount,
      rows:          resultRows,
    };

    saveReport(report);
    loadStats();

    let summary = `Done — ${pings.length} GPS pings processed · ${matchedCount} sites Verified (within 50 m)`;
    if (skippedCount) summary += ` · ${skippedCount} pings skipped (no GPS)`;
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
    const statusCell = `<span class="status-pill pill-green">✔ Work Done, Verified</span>`;
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

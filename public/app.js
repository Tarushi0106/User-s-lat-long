/* ── All logic runs in the browser. No backend server required. ── */

let masterSites = [];
let currentRows = [];

// ── Boot ───────────────────────────────────────────────────────────────────
document.getElementById('headerDate').textContent = new Date().toLocaleDateString('en-IN', {
  weekday: 'short', day: '2-digit', month: 'short', year: 'numeric'
});
document.getElementById('footerYear').textContent = new Date().getFullYear();

// ── Toast ──────────────────────────────────────────────────────────────────
function toast(msg, ms = 3000) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), ms);
}

// ── Haversine ──────────────────────────────────────────────────────────────
function haversineMeters(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// ── Load master data from bundled Excel ───────────────────────────────────
document.getElementById('seedBtn').addEventListener('click', syncMaster);

async function syncMaster() {
  const btn = document.getElementById('seedBtn');
  btn.disabled = true;
  btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="animation:spin 1s linear infinite"><polyline points="1 4 1 10 7 10"/><polyline points="23 20 23 14 17 14"/><path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4l-4.64 4.36A9 9 0 0 1 3.51 15"/></svg> Syncing…`;
  try {
    const res = await fetch('master.xlsx');
    const buf = await res.arrayBuffer();
    const wb  = XLSX.read(new Uint8Array(buf), { type: 'array' });
    const ws  = wb.Sheets['Site master'];
    if (!ws) throw new Error('Sheet "Site master" not found');

    masterSites = parseMasterSheet(ws);
    localStorage.setItem('masterSites', JSON.stringify(masterSites));
    toast(`✔ ${masterSites.length} sites loaded`);
    loadStats();
  } catch (e) {
    toast('Error loading master: ' + e.message);
  }
  btn.disabled = false;
  btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="1 4 1 10 7 10"/><polyline points="23 20 23 14 17 14"/><path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4l-4.64 4.36A9 9 0 0 1 3.51 15"/></svg> Sync Master Data`;
}

function parseMasterSheet(ws) {
  const rows    = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const headers = rows[0].map(h => String(h || '').trim().toLowerCase());
  const ci = (...names) => {
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
  const iLat  = ci('latitude');
  const iLng  = ci('longitude');
  const iId   = ci('stpl site id', 'site id');
  const iName = ci('site name');
  const iCirc = ci('circle name', 'circle');
  const iDist = ci('district');
  const iPers = ci('site acquisition person name');

  const sites = [];
  for (let i = 1; i < rows.length; i++) {
    const r   = rows[i];
    const lat = parseFloat(r[iLat]);
    const lng = parseFloat(r[iLng]);
    if (!lat || !lng || isNaN(lat) || isNaN(lng)) continue;
    sites.push({
      stsId:  String(r[iId]   || '').trim(),
      name:   String(r[iName] || '').trim(),
      circle: String(r[iCirc] || '').trim(),
      dist:   String(r[iDist] || '').trim(),
      person: String(r[iPers] || '').trim(),
      lat, lng,
    });
  }
  return sites;
}

// ── Stats from localStorage ────────────────────────────────────────────────
function loadStats() {
  const reports = getReports();
  const totalMatched   = reports.reduce((s, r) => s + r.matchedCount,   0);
  const totalUnmatched = reports.reduce((s, r) => s + r.unmatchedCount, 0);
  document.getElementById('statUploads').textContent   = reports.length;
  document.getElementById('statRows').textContent      = totalMatched + totalUnmatched;
  document.getElementById('statMatched').textContent   = totalMatched;
  document.getElementById('statUnmatched').textContent = totalUnmatched;
  renderRecentReports(reports.slice(0, 5));
}

// ── Reports (localStorage) ─────────────────────────────────────────────────
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
    el.innerHTML = '<div class="empty-msg">No uploads yet. Click "Sync Master Data" first, then upload a tracker file.</div>';
    return;
  }
  el.innerHTML = reports.map(r => `
    <div class="report-list-item">
      <div onclick="viewReport('${r.id}')" style="flex:1;cursor:pointer;">
        <div class="report-name">${esc(r.fileName)}</div>
        <div class="report-meta">By ${esc(r.uploadedBy)} &nbsp;·&nbsp; ${r.createdAt}</div>
      </div>
      <div class="report-badges">
        <span class="badge badge-green">✔ ${r.matchedCount} Work Done, Verified</span>
        <span class="badge badge-red">✘ ${r.unmatchedCount} Work Not Done, Not Verified</span>
        <button class="btn-delete" onclick="deleteReport(event,'${r.id}')" title="Delete">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/></svg>
        </button>
      </div>
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

// ── File input ─────────────────────────────────────────────────────────────
document.getElementById('fileInput').addEventListener('change', function () {
  document.getElementById('fileName').textContent = this.files[0]?.name || 'Choose file…';
});

// ── Upload & Match ─────────────────────────────────────────────────────────
document.getElementById('uploadForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const file = document.getElementById('fileInput').files[0];
  if (!file) return;

  // Load master from cache if not already loaded
  if (!masterSites.length) {
    const cached = localStorage.getItem('masterSites');
    if (cached) masterSites = JSON.parse(cached);
  }
  if (!masterSites.length) {
    document.getElementById('uploadMsg').textContent = 'Please click "Sync Master Data" first.';
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
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    if (rows.length < 2) throw new Error('File has no data rows');

    const headers = rows[0].map(h => String(h || '').trim().toLowerCase());
    const ci = (...names) => {
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

    const colLat     = ci('lat', 'latitude');
    const colLng     = ci('lng', 'long', 'longitude');
    const colTracker = ci('tracker_id', 'tracker', 'device');
    const colTime    = ci('time (gmt', 'time', 'timestamp');

    if (colLat === -1 || colLng === -1) throw new Error('Lat/Lng columns not found');

    // Extract person name
    let personName = document.getElementById('uploadedBy').value.trim() || 'Unknown';
    for (let i = 1; i < rows.length; i++) {
      if (colTracker !== -1 && rows[i][colTracker]) {
        const raw = String(rows[i][colTracker]).trim();
        personName = raw.includes('@') ? raw.split('@')[1] : raw;
        break;
      }
    }

    // Collect valid pings
    const pings = [];
    for (let i = 1; i < rows.length; i++) {
      const r   = rows[i];
      const lat = parseFloat(r[colLat]);
      const lng = parseFloat(r[colLng]);
      if (!lat || !lng || isNaN(lat) || isNaN(lng)) continue;
      pings.push({ lat, lng, time: colTime !== -1 ? String(r[colTime] || '') : '' });
    }

    if (!pings.length) throw new Error('No valid GPS coordinates (all were 0,0)');

    // Match pings against all master sites
    const TOLERANCE = 500; // metres
    const resultRows = [];
    let matchedCount = 0;

    for (const site of masterSites) {
      let bestDist = Infinity;
      let bestTime = '';
      for (const ping of pings) {
        const degTol = TOLERANCE / 111000;
        if (Math.abs(ping.lat - site.lat) > degTol) continue;
        if (Math.abs(ping.lng - site.lng) > degTol) continue;
        const d = haversineMeters(ping.lat, ping.lng, site.lat, site.lng);
        if (d < bestDist) { bestDist = d; bestTime = ping.time; }
      }
      if (bestDist > TOLERANCE) continue;
      matchedCount++;
      resultRows.push({
        rowNumber:       matchedCount,
        personName,
        timeOfVisit:     bestTime,
        matchedSiteId:   site.stsId,
        matchedSiteName: site.name,
        district:        site.dist,
        circle:          site.circle,
        distanceMeters:  Math.round(bestDist),
        matched:         true,
        status:          'Work Done - Verified',
      });
    }

    const report = {
      id:            Date.now().toString(),
      fileName:      file.name,
      uploadedBy:    personName,
      createdAt:     new Date().toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit' }),
      matchedCount,
      unmatchedCount: 0,
      totalRows:     matchedCount,
      rows:          resultRows,
    };

    saveReport(report);
    loadStats();

    const pct = pings.length > 0 ? Math.round(matchedCount / masterSites.length * 100) : 0;
    msgEl.textContent = `Done — ${matchedCount} sites verified out of ${masterSites.length} (${pings.length} valid GPS pings processed)`;
    msgEl.className   = 'success';

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
  const rows = currentRows.filter(r => {
    const hay = `${r.personName} ${r.matchedSiteId} ${r.matchedSiteName} ${r.district}`.toLowerCase();
    return (!q || hay.includes(q)) && (!st || r.status === st);
  });
  document.getElementById('resultCount').textContent = rows.length + ' of ' + currentRows.length + ' rows';
  renderTable(rows);
}

function renderTable(rows) {
  document.getElementById('resultCount').textContent = rows.length + ' rows';
  const tbody = document.getElementById('reportBody');
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:28px;color:var(--muted)">No sites verified.</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map(r => `
    <tr>
      <td style="color:var(--muted)">${r.rowNumber}</td>
      <td><strong>${esc(r.personName)}</strong></td>
      <td style="font-size:.8rem;color:var(--muted)">${esc(r.matchedSiteId) || '–'}</td>
      <td>${esc(r.matchedSiteName) || '–'}</td>
      <td style="font-size:.82rem">${esc(r.district) || '–'}</td>
      <td>${r.distanceMeters != null ? `<span class="dist-tag">${r.distanceMeters} m</span>` : '–'}</td>
      <td style="font-size:.82rem">${esc(r.timeOfVisit) || '–'}</td>
      <td><span class="status-pill pill-green">✔ Work Done, Verified</span></td>
    </tr>
  `).join('');
}

// ── Helpers ────────────────────────────────────────────────────────────────
function esc(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

const style = document.createElement('style');
style.textContent = '@keyframes spin { to { transform: rotate(360deg); } }';
document.head.appendChild(style);

// ── Init ───────────────────────────────────────────────────────────────────
const cached = localStorage.getItem('masterSites');
if (cached) masterSites = JSON.parse(cached);
loadStats();

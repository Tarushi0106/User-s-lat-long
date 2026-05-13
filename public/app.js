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
    const buf = await file.arrayBuffer();
    const wb  = XLSX.read(new Uint8Array(buf), { type: 'array' });

    // Pick the sheet that actually contains India GPS coordinates.
    // Files like productivity reports have a summary/pivot sheet first;
    // the real data (with lat/long) is in a later sheet.
    const _gpsRe = /(\d{1,3}\.\d+)\s*,\s*(\d{1,3}\.\d+)/;
    function sheetHasCoords(s) {
      const sample = XLSX.utils.sheet_to_json(s, { header: 1, defval: '' }).slice(0, 60);
      return sample.some(row => row.some(cell => {
        const m = _gpsRe.exec(String(cell || ''));
        if (!m) return false;
        const lat = parseFloat(m[1]), lng = parseFloat(m[2]);
        return lat >= 6 && lat <= 38 && lng >= 60 && lng <= 100;
      }));
    }
    let ws = wb.Sheets[wb.SheetNames[0]];
    for (const name of wb.SheetNames) {
      if (sheetHasCoords(wb.Sheets[name])) { ws = wb.Sheets[name]; break; }
    }

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    if (rows.length < 2) throw new Error('File has no data rows');

    // Auto-detect the real header row — skip title / merged / blank rows at top.
    // A real header row has ≥ 3 non-empty cells AND ≥ 2 of them match known column keywords.
    const HDR_KEYWORDS = ['lat', 'long', 'lng', 'gps', 'coordinate', 'location',
                          'name', 'employee', 'person', 'staff', 'date', 'time',
                          'timestamp', 'id', 'site', 'status', 'tracker', 'device'];
    let headerRowIdx = 0;
    for (let r = 0; r < Math.min(rows.length, 8); r++) {
      const cells    = rows[r].map(c => String(c || '').trim().toLowerCase());
      const nonEmpty = cells.filter(Boolean).length;
      const kwHits   = cells.filter(c => HDR_KEYWORDS.some(k => c.includes(k))).length;
      if (nonEmpty >= 3 && kwHits >= 2) { headerRowIdx = r; break; }
    }

    const headers = rows[headerRowIdx].map(h => String(h || '').trim().toLowerCase());
    const ci = makeCi(headers);

    const colTime    = ci('time (gmt', 'time', 'timestamp', 'date');
    const colTracker = ci('tracker_id', 'tracker', 'device');
    const colPerson  = ci('full name', 'employee name', 'staff name', 'person name',
                          'field person', 'person', 'name', 'employee', 'user', 'field', 'engineer');

    // ── Extract all valid India lat/lng pairs from any cell value ─────────
    // Handles comma-separated:  "28.43, 77.31"  |  "28.43,77.31"
    //         space-separated:  "19.082862 72.851958"
    //         multiple pairs:   "28.43,77.31 & 28.50,77.09"
    //         mixed text:       "28.63,77.38-SITE-ID, 28.62,77.29"
    function extractCoords(cellValue) {
      const str    = String(cellValue == null ? '' : cellValue);
      const coords = [];
      // Pass 1: comma-separated (permissive decimal)
      const re1 = /(-?\d{1,3}(?:\.\d+)?)\s*,\s*(-?\d{1,3}(?:\.\d+)?)/g;
      // Pass 2: space-separated (require ≥4 decimal places to avoid false positives)
      const re2 = /(-?\d{1,3}\.\d{4,})\s+(-?\d{2,3}\.\d{4,})/g;
      for (const re of [re1, re2]) {
        let m;
        while ((m = re.exec(str)) !== null) {
          const lat = parseFloat(m[1]), lng = parseFloat(m[2]);
          if (lat >= 6 && lat <= 38 && lng >= 60 && lng <= 100)
            coords.push({ lat, lng });
        }
      }
      return coords;
    }

    // ── Detect which column(s) hold coordinates ───────────────────────────
    // Strategy 1: named separate lat / lng columns (must be distinct and numeric)
    let colLat = ci('latitude', 'lat');
    let colLng = ci('longitude', 'lon', 'long', 'lng');
    let colCombined = -1;

    if (colLat !== -1 && colLng !== -1 && colLat !== colLng) {
      const samp = rows.slice(headerRowIdx + 1, Math.min(rows.length, headerRowIdx + 11));
      const tLat = parseFloat(String(samp.map(r => r[colLat]).find(v => String(v || '').trim()) || ''));
      const tLng = parseFloat(String(samp.map(r => r[colLng]).find(v => String(v || '').trim()) || ''));
      if (isNaN(tLat) || isNaN(tLng) || tLat < 6 || tLat > 38 || tLng < 60 || tLng > 100) {
        colLat = -1; colLng = -1;
      }
    } else {
      colLat = -1; colLng = -1; // same column or both missing → fall through to combined scan
    }

    // Strategy 2: scan every column and pick whichever has the most valid coord pairs.
    // This handles combined "lat,lng" cells, non-standard headers, merged cells, etc.
    if (colLat === -1 || colLng === -1) {
      const numCols   = rows[headerRowIdx] ? rows[headerRowIdx].length : 0;
      const scanStart = headerRowIdx + 1;
      const scanLimit = Math.min(rows.length, scanStart + 200);
      let bestCount = 0;
      for (let c = 0; c < numCols; c++) {
        let count = 0;
        for (let i = scanStart; i < scanLimit; i++) {
          if (rows[i] && extractCoords(rows[i][c]).length > 0) count++;
        }
        if (count > bestCount) { bestCount = count; colCombined = c; }
      }
    }

    if (colLat === -1 && colLng === -1 && colCombined === -1) {
      const preview = headers.slice(0, 10).filter(Boolean).join(' | ');
      throw new Error(`Could not find GPS coordinates. Columns detected: ${preview || '(none)'}`);
    }

    // ── Detect file mode ──────────────────────────────────────────────────
    // Productivity report: one summary row per employee (show ALL rows + status).
    // GPS tracker: many pings per person (show only verified site visits).
    const manualName   = document.getElementById('uploadedBy').value.trim();
    const allDataRows  = rows.slice(headerRowIdx + 1).filter(r => r.some(c => String(c || '').trim()));
    const nameSet      = new Set();
    if (colPerson !== -1)
      allDataRows.forEach(r => { const n = String(r[colPerson] || '').trim(); if (n) nameSet.add(n); });
    // Productivity mode: >1 unique person AND ≥40 % of rows are unique persons
    const isProductivityReport = nameSet.size > 1 && nameSet.size / allDataRows.length >= 0.4;

    const TOLERANCE  = 50; // metres
    const resultRows = [];
    let matchedCount = 0;
    let reportLabel  = '';
    let summary      = '';

    if (isProductivityReport) {
      // ── PRODUCTIVITY REPORT: one row per employee ─────────────────────────
      // Show every employee — Verified | Not at Master Site | Leave/WFH/etc.
      const colRemark = ci('validated remark', 'remark', 'validated', 'attendance', 'status');

      for (let i = headerRowIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r.some(c => String(c || '').trim())) continue;

        const personName = (colPerson !== -1 ? String(r[colPerson] || '').trim() : '')
          || manualName || 'Unknown';
        const fileStatus = colRemark !== -1 ? String(r[colRemark] || '').trim() : '';

        // Get this row's GPS coordinates
        let rowLat = null, rowLng = null;
        if (colLat !== -1 && colLng !== -1) {
          const la = parseFloat(r[colLat]), lo = parseFloat(r[colLng]);
          if (!isNaN(la) && !isNaN(lo) && la && lo) { rowLat = la; rowLng = lo; }
        } else if (colCombined !== -1) {
          const coords = extractCoords(r[colCombined]);
          if (coords.length) { rowLat = coords[0].lat; rowLng = coords[0].lng; }
        }

        // Find nearest master site (no 50 m cap — always show nearest for context)
        let nearestSite = null, nearestDist = Infinity;
        if (rowLat !== null) {
          for (const site of masterSites) {
            const d = haversineMeters(rowLat, rowLng, site.lat, site.lng);
            if (d < nearestDist) { nearestDist = d; nearestSite = site; }
          }
        }

        const verified = nearestDist <= TOLERANCE;
        if (verified) matchedCount++;
        const status = rowLat === null
          ? (fileStatus || 'No GPS')
          : (verified ? 'Work Done - Verified' : 'Not at Master Site');

        resultRows.push({
          rowNumber:       resultRows.length + 1,
          personName,
          timeOfVisit:     '',
          userLat:         rowLat,
          userLng:         rowLng,
          matchedSiteId:   nearestSite?.stsId   || '',
          matchedSiteName: nearestSite?.name    || '',
          district:        nearestSite?.dist    || '',
          circle:          nearestSite?.circle  || '',
          masterSource:    nearestSite?.source  || '',
          masterLat:       nearestSite?.lat     ?? null,
          masterLng:       nearestSite?.lng     ?? null,
          distanceMeters:  nearestDist !== Infinity ? Math.round(nearestDist) : null,
          matched:         verified,
          status,
        });
      }

      reportLabel = `${nameSet.size} persons`;
      summary     = `Done — ${resultRows.length} employees · ${matchedCount} Verified · ${resultRows.length - matchedCount} other`;

    } else {
      // ── GPS TRACKER: many pings per person → show only verified site visits ─
      const pingsByPerson = new Map();
      let skippedCount = 0;

      for (let i = headerRowIdx + 1; i < rows.length; i++) {
        const r    = rows[i];
        const time = colTime !== -1 ? String(r[colTime] || '') : '';

        let rowPerson = manualName;
        if (!rowPerson) {
          if (colPerson !== -1 && String(r[colPerson] || '').trim())
            rowPerson = String(r[colPerson]).trim();
          else if (colTracker !== -1 && String(r[colTracker] || '').trim()) {
            const raw = String(r[colTracker]).trim();
            rowPerson = raw.includes('@') ? raw.split('@')[1] : raw;
          }
        }
        if (!rowPerson) rowPerson = 'Unknown';

        if (!pingsByPerson.has(rowPerson)) pingsByPerson.set(rowPerson, []);
        const bucket = pingsByPerson.get(rowPerson);

        if (colLat !== -1 && colLng !== -1) {
          const lat = parseFloat(r[colLat]), lng = parseFloat(r[colLng]);
          if (!lat || !lng || isNaN(lat) || isNaN(lng)) { skippedCount++; continue; }
          bucket.push({ lat, lng, time });
        } else {
          const coords = extractCoords(r[colCombined]);
          if (!coords.length) { skippedCount++; continue; }
          for (const { lat, lng } of coords) bucket.push({ lat, lng, time });
        }
      }

      const totalPings = [...pingsByPerson.values()].reduce((s, a) => s + a.length, 0);
      if (!totalPings) throw new Error('No valid GPS coordinates found in the uploaded file');

      for (const [personName, pings] of pingsByPerson) {
        for (const site of masterSites) {
          let nearestDist = Infinity, nearestTime = '', nearestLat = null, nearestLng = null;
          const degTol = TOLERANCE / 111000;
          for (const ping of pings) {
            if (Math.abs(ping.lat - site.lat) > degTol) continue;
            if (Math.abs(ping.lng - site.lng) > degTol) continue;
            const d = haversineMeters(ping.lat, ping.lng, site.lat, site.lng);
            if (d < nearestDist) { nearestDist = d; nearestTime = ping.time; nearestLat = ping.lat; nearestLng = ping.lng; }
          }
          if (nearestDist > TOLERANCE) continue;
          matchedCount++;
          resultRows.push({
            rowNumber:       matchedCount,
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
      }

      const personCount = pingsByPerson.size;
      reportLabel = personCount > 1 ? `${personCount} persons` : ([...pingsByPerson.keys()][0] || 'Unknown');
      summary = `Done — ${totalPings} GPS pings · ${personCount > 1 ? personCount + ' persons · ' : ''}${matchedCount} sites Verified (within 50 m)`;
      if (skippedCount) summary += ` · ${skippedCount} rows skipped`;
    }
    const report = {
      id:         Date.now().toString(),
      fileName:   file.name,
      uploadedBy: reportLabel,
      createdAt:  new Date().toLocaleDateString('en-IN', {
        day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit',
      }),
      matchedCount,
      totalRows:  resultRows.length,
      rows:       resultRows,
    };

    saveReport(report);
    loadStats();

    msgEl.innerHTML = `${summary} &nbsp;<button onclick="downloadCurrentReport()" style="margin-left:8px;padding:4px 12px;background:#15803d;color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:700;cursor:pointer;font-family:inherit;">&#8595; Download Excel</button>`;
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
    const distFmt  = r.distanceMeters != null
      ? (r.distanceMeters < 1000 ? r.distanceMeters + ' m' : (r.distanceMeters / 1000).toFixed(1) + ' km')
      : null;
    const distCell = distFmt
      ? `<span class="dist-tag${r.matched ? '' : ' dist-far'}">${distFmt}</span>`
      : '–';
    let statusCell;
    if (r.status === 'Work Done - Verified')
      statusCell = `<span class="status-pill pill-green">✔ Work Done, Verified</span>`;
    else if (r.status === 'Not at Master Site')
      statusCell = `<span class="status-pill pill-orange">✗ Not at Site</span>`;
    else
      statusCell = `<span class="status-pill" style="background:var(--gray-bg);color:var(--gray)">${esc(r.status)}</span>`;
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
      <td style="font-size:.82rem">${fmtTime(r.timeOfVisit)}</td>
      <td>${statusCell}</td>
    </tr>`;
  }).join('');
}

// ── Download current report as Excel ──────────────────────────────────────
function downloadCurrentReport() {
  if (!currentRows.length) return;
  const title  = document.getElementById('reportTitle').textContent;
  const colHdr = ['#', 'Person Name', 'Site ID', 'Site Name (Master)', 'District', 'Circle',
                  'Person GPS', 'Master GPS', 'Distance (m)', 'Time', 'Status'];
  const data = currentRows.map(r => [
    r.rowNumber,
    r.personName        || '',
    r.matchedSiteId     || '',
    r.matchedSiteName   || '',
    r.district          || '',
    r.circle            || '',
    r.userLat   != null ? `${r.userLat.toFixed(6)}, ${r.userLng.toFixed(6)}`     : '',
    r.masterLat != null ? `${r.masterLat.toFixed(6)}, ${r.masterLng.toFixed(6)}` : '',
    r.distanceMeters != null ? r.distanceMeters : '',
    fmtTime(r.timeOfVisit),
    r.status      || '',
  ]);

  const ws = XLSX.utils.aoa_to_sheet([colHdr, ...data]);

  // Column widths
  ws['!cols'] = [
    { wch: 5 }, { wch: 22 }, { wch: 22 }, { wch: 28 },
    { wch: 18 }, { wch: 14 }, { wch: 26 }, { wch: 26 },
    { wch: 13 }, { wch: 18 }, { wch: 22 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Verification');
  const safeName = title.replace(/[\\/:*?"<>|]/g, '_').slice(0, 60);
  XLSX.writeFile(wb, safeName + '.xlsx');
}

// ── Helpers ────────────────────────────────────────────────────────────────
const _MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function fmtTime(val) {
  if (!val && val !== 0) return '–';
  const s = String(val).trim();
  if (!s) return '–';
  let d;
  if (/^\d{9,13}$/.test(s)) {
    // Unix timestamp: 10 digits = seconds, 13 digits = milliseconds
    const ms = s.length <= 10 ? Number(s) * 1000 : Number(s);
    d = new Date(ms);
  } else {
    d = new Date(s);
  }
  if (isNaN(d.getTime())) return s;
  const dd  = String(d.getDate()).padStart(2, '0');
  const mon = _MONTHS[d.getMonth()];
  const hh  = String(d.getHours()).padStart(2, '0');
  const min = String(d.getMinutes()).padStart(2, '0');
  const sec = String(d.getSeconds()).padStart(2, '0');
  return `${dd} ${mon} ${hh}:${min}:${sec}`;
}
function esc(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

const style = document.createElement('style');
style.textContent = '@keyframes spin { to { transform: rotate(360deg); } }';
document.head.appendChild(style);

// ── Init ───────────────────────────────────────────────────────────────────
loadStats();
autoLoadMasters();

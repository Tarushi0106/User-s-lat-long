const API = window.location.hostname === 'localhost'
  ? ''
  : 'https://your-app.railway.app';  // replace with your Railway URL
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

// ── Stats ──────────────────────────────────────────────────────────────────
async function loadStats() {
  try {
    const data = await fetch(`${API}/api/reports/dashboard/stats`).then(r => r.json());
    const s = data.stats;
    document.getElementById('statUploads').textContent   = (s.totalUploads   ?? 0).toLocaleString();
    document.getElementById('statRows').textContent      = (s.totalRows       ?? 0).toLocaleString();
    document.getElementById('statMatched').textContent   = (s.totalMatched    ?? 0).toLocaleString();
    document.getElementById('statUnmatched').textContent = (s.totalUnmatched  ?? 0).toLocaleString();
    renderRecentReports(data.recentReports || []);
  } catch (e) { console.error(e); }
}

// ── Recent reports ─────────────────────────────────────────────────────────
function renderRecentReports(reports) {
  const el = document.getElementById('recentReports');
  if (!reports.length) {
    el.innerHTML = '<div class="empty-msg">No uploads yet.</div>';
    return;
  }
  el.innerHTML = reports.map(r => `
    <div class="report-list-item">
      <div onclick="loadReport('${r._id}')" style="flex:1;cursor:pointer;">
        <div class="report-name">${esc(r.fileName)}</div>
        <div class="report-meta">By ${esc(r.uploadedBy)} &nbsp;·&nbsp; ${fmtDate(r.createdAt)}</div>
      </div>
      <div class="report-badges">
        <span class="badge badge-green">✔ ${r.matchedCount} Work Done, Verified</span>
        <span class="badge badge-red">✘ ${r.unmatchedCount} Work Not Done, Not Verified</span>
        <button class="btn-delete" onclick="deleteReport(event,'${r._id}')" title="Delete">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/></svg>
        </button>
      </div>
    </div>
  `).join('');
}

// ── Sync master data ───────────────────────────────────────────────────────
document.getElementById('seedBtn').addEventListener('click', async () => {
  const btn = document.getElementById('seedBtn');
  btn.disabled = true;
  btn.textContent = 'Syncing…';
  try {
    const data = await fetch(`${API}/api/sites/seed`, { method: 'POST' }).then(r => r.json());
    toast(data.success ? `✔ ${data.seeded} sites loaded.` : 'Sync failed: ' + data.error);
  } catch (e) { toast('Error: ' + e.message); }
  btn.disabled = false;
  btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="1 4 1 10 7 10"/><polyline points="23 20 23 14 17 14"/><path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4l-4.64 4.36A9 9 0 0 1 3.51 15"/></svg> Sync Master Data`;
});

// ── File pick ──────────────────────────────────────────────────────────────
document.getElementById('fileInput').addEventListener('change', function () {
  document.getElementById('fileName').textContent = this.files[0]?.name || 'Choose file…';
});

// ── Upload ─────────────────────────────────────────────────────────────────
document.getElementById('uploadForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const file = document.getElementById('fileInput').files[0];
  if (!file) return;

  const btn   = document.getElementById('uploadBtn');
  const msgEl = document.getElementById('uploadMsg');
  btn.disabled = true;
  btn.textContent = 'Processing…';
  msgEl.textContent = '';
  msgEl.className = '';

  const fd = new FormData();
  fd.append('file', file);
  fd.append('uploadedBy', document.getElementById('uploadedBy').value.trim() || 'Admin');

  try {
    const data = await fetch(`${API}/api/upload`, { method: 'POST', body: fd }).then(r => r.json());
    if (data.success) {
      const pct = data.totalRows ? Math.round(data.matchedCount / data.totalRows * 100) : 0;
      msgEl.textContent = `Done — ${data.matchedCount} verified, ${data.unmatchedCount} not matched out of ${data.totalRows} rows (${pct}% match rate)`;
      msgEl.className = 'success';
      loadStats();
      loadReport(data.reportId);
    } else {
      msgEl.textContent = data.error || 'Upload failed';
      msgEl.className = 'error';
    }
  } catch (err) {
    msgEl.textContent = 'Error: ' + err.message;
    msgEl.className = 'error';
  }

  btn.disabled = false;
  btn.textContent = 'Upload & Match';
});

// ── Delete report ──────────────────────────────────────────────────────────
async function deleteReport(e, id) {
  e.stopPropagation();
  if (!confirm('Delete this report?')) return;
  try {
    await fetch(`${API}/api/reports/${id}`, { method: 'DELETE' });
    toast('Report deleted.');
    loadStats();
    document.getElementById('reportDetail').style.display = 'none';
  } catch (err) {
    toast('Delete failed: ' + err.message);
  }
}

// ── Report detail ──────────────────────────────────────────────────────────
async function loadReport(id) {
  try {
    const report = await fetch(`${API}/api/reports/${id}`).then(r => r.json());
    currentRows  = report.rows || [];
    document.getElementById('reportTitle').textContent = report.fileName + ' — ' + fmtDate(report.createdAt);
    document.getElementById('reportDetail').style.display = '';
    document.getElementById('searchInput').value  = '';
    document.getElementById('statusFilter').value = '';
    renderTable(currentRows);
    setTimeout(() => document.getElementById('reportDetail').scrollIntoView({ behavior: 'smooth' }), 80);
  } catch (e) { toast('Could not load report'); }
}

document.getElementById('closeDetail').addEventListener('click', () => {
  document.getElementById('reportDetail').style.display = 'none';
});

document.getElementById('searchInput').addEventListener('input', filterTable);
document.getElementById('statusFilter').addEventListener('change', filterTable);

function filterTable() {
  const q  = document.getElementById('searchInput').value.toLowerCase();
  const st = document.getElementById('statusFilter').value;
  const rows = currentRows.filter(r => {
    const hay = `${r.personName} ${r.matchedSiteId} ${r.matchedSiteName} ${r.district} ${r.circle} ${r.status}`.toLowerCase();
    return (!q || hay.includes(q)) && (!st || r.status === st);
  });
  document.getElementById('resultCount').textContent = rows.length + ' of ' + currentRows.length + ' rows';
  renderTable(rows);
}

function renderTable(rows) {
  document.getElementById('resultCount').textContent = rows.length + ' rows';
  const tbody = document.getElementById('reportBody');
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:28px;color:var(--muted)">No rows found.</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map(r => `
    <tr>
      <td style="color:var(--muted)">${r.rowNumber}</td>
      <td><strong>${esc(r.personName) || '–'}</strong></td>
      <td style="font-size:.8rem;color:var(--muted)">${esc(r.matchedSiteId) || '–'}</td>
      <td>${esc(r.matchedSiteName) || '<span style="color:var(--muted)">–</span>'}</td>
      <td style="font-size:.82rem;">${esc(r.district) || '–'}</td>
      <td>${r.distanceMeters != null ? `<span class="dist-tag">${r.distanceMeters} m</span>` : '–'}</td>
      <td style="font-size:.82rem;">${esc(r.timeOfVisit) || '–'}</td>
      <td><span class="status-pill ${r.matched ? 'pill-green' : 'pill-red'}">${r.matched ? '✔ Work Done, Verified' : '✘ Work Not Done, Not Verified'}</span></td>
    </tr>
  `).join('');
}

// ── Helpers ────────────────────────────────────────────────────────────────
function esc(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
function fmtDate(iso) {
  if (!iso) return '';
  return new Date(iso).toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit' });
}

loadStats();

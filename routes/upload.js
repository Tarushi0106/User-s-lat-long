const express = require('express');
const router  = express.Router();
const multer  = require('multer');
const path    = require('path');
const XLSX    = require('xlsx');
const Site    = require('../models/Site');
const Report  = require('../models/Report');

const storage = multer.diskStorage({
  destination: path.join(__dirname, '..', 'uploads'),
  filename: (req, file, cb) => cb(null, `${Date.now()}_${file.originalname}`),
});
const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    ['.xlsx', '.xls', '.csv'].includes(ext) ? cb(null, true) : cb(new Error('Only Excel/CSV files allowed'));
  },
});

function haversineMeters(lat1, lon1, lat2, lon2) {
  const R    = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a    = Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

router.post('/', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  try {
    const toleranceMeters = parseFloat(process.env.TOLERANCE_METERS) || 500;

    const wb   = XLSX.readFile(req.file.path);
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (rows.length < 2) return res.status(400).json({ error: 'Excel has no data rows' });

    const rawHeaders = rows[0];
    const headers    = rawHeaders.map(h => String(h || '').trim().toLowerCase());

    // Find column index — exact match first, then partial
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

    if (colLat === -1 || colLng === -1) {
      return res.status(400).json({
        error: 'Could not find Lat/Lng columns.',
        detectedHeaders: rawHeaders,
      });
    }

    // ── Extract person name from first tracker_id row ─────────────────────
    let personName = req.body.uploadedBy || 'Unknown';
    for (let i = 1; i < rows.length; i++) {
      if (colTracker !== -1 && rows[i][colTracker]) {
        const raw = String(rows[i][colTracker]).trim();
        personName = raw.includes('@') ? raw.split('@')[1] : raw;
        break;
      }
    }

    // ── Collect valid GPS pings (skip lat=0 / lng=0) ───────────────────────
    const pings = [];
    for (let i = 1; i < rows.length; i++) {
      const r   = rows[i];
      const lat = parseFloat(r[colLat]);
      const lng = parseFloat(r[colLng]);
      if (!lat || !lng || isNaN(lat) || isNaN(lng)) continue;
      pings.push({ lat, lng, time: colTime !== -1 ? String(r[colTime] || '') : '' });
    }

    if (!pings.length) {
      return res.status(400).json({ error: 'No valid GPS coordinates found — all rows had lat=0 / lng=0.' });
    }

    // ── Load ALL master sites ──────────────────────────────────────────────
    const allSites = await Site.find(
      {},
      'lat long stsId siteName district circle acquisitionPerson'
    ).lean();

    // ── Match: for each site, check if any ping was within toleranceMeters ─
    // Result = one row per SITE that was visited (not one row per ping)
    const resultRows = [];
    let matchedCount = 0;

    for (const site of allSites) {
      if (!site.lat || !site.long) continue;

      let bestDist = Infinity;
      let bestTime = '';

      for (const ping of pings) {
        // Fast bounding box (1 degree ≈ 111km, so toleranceMeters/111000 degrees)
        const degTol = toleranceMeters / 111000;
        if (Math.abs(ping.lat - site.lat) > degTol) continue;
        if (Math.abs(ping.lng - site.long) > degTol) continue;

        const dist = haversineMeters(ping.lat, ping.lng, site.lat, site.long);
        if (dist < bestDist) {
          bestDist = dist;
          bestTime = ping.time;
        }
      }

      const visited = bestDist <= toleranceMeters;
      if (!visited) continue; // only include sites that were actually visited

      matchedCount++;
      resultRows.push({
        rowNumber:       resultRows.length + 1,
        personName,
        uploadedLat:     null,
        uploadedLong:    null,
        timeOfVisit:     bestTime,
        matched:         true,
        matchedSite:     site._id,
        matchedSiteId:   site.stsId,
        matchedSiteName: site.siteName,
        matchedLat:      site.lat,
        matchedLong:     site.long,
        distanceMeters:  Math.round(bestDist),
        district:        site.district,
        circle:          site.circle,
        assignedTo:      site.acquisitionPerson || '',
        status:          'Work Done - Verified',
      });
    }

    const report = await Report.create({
      fileName:       req.file.originalname,
      uploadedBy:     personName,
      totalRows:      matchedCount,       // only verified sites count
      matchedCount,
      unmatchedCount: 0,                  // we only store visited sites now
      rows:           resultRows,
    });

    res.json({
      success:       true,
      reportId:      report._id,
      totalRows:     matchedCount,
      matchedCount,
      unmatchedCount: 0,
      personName,
      totalPings:    pings.length,
    });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;

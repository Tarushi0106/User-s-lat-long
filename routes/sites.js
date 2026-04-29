const express = require('express');
const router  = express.Router();
const path    = require('path');
const XLSX    = require('xlsx');
const Site    = require('../models/Site');

router.post('/seed', async (req, res) => {
  try {
    const filePath = path.join(__dirname, '..', process.env.MASTER_EXCEL);
    const wb   = XLSX.readFile(filePath);
    const ws   = wb.Sheets['Site master'];

    if (!ws) return res.status(400).json({ error: 'Sheet "Site master" not found in Excel.' });

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

    const idxLat    = ci('latitude');
    const idxLong   = ci('longitude');
    const idxStsId  = ci('stpl site id', 'site id');
    const idxName   = ci('site name');
    const idxCircle = ci('circle name', 'circle');
    const idxDist   = ci('district');
    const idxStatus = ci('tower status', 'updated status');
    const idxAddr   = ci('site address', 'address');
    const idxPerson = ci('site acquisition person name');

    if (idxLat === -1 || idxLong === -1) {
      return res.status(400).json({ error: 'Latitude/Longitude columns not found in Site master sheet.' });
    }

    const docs = [];
    for (let i = 1; i < rows.length; i++) {
      const r   = rows[i];
      const lat = parseFloat(r[idxLat]);
      const lng = parseFloat(r[idxLong]);
      if (!lat || !lng || isNaN(lat) || isNaN(lng)) continue;

      docs.push({
        stsId:             idxStsId  !== -1 ? String(r[idxStsId]  || '').trim() : '',
        siteName:          idxName   !== -1 ? String(r[idxName]   || '').trim() : '',
        circle:            idxCircle !== -1 ? String(r[idxCircle] || '').trim() : '',
        district:          idxDist   !== -1 ? String(r[idxDist]   || '').trim() : '',
        updatedStatus:     idxStatus !== -1 ? String(r[idxStatus] || '').trim() : '',
        lat,
        long:              lng,
        address:           idxAddr   !== -1 ? String(r[idxAddr]   || '').trim() : '',
        acquisitionPerson: idxPerson !== -1 ? String(r[idxPerson] || '').trim() : '',
        sheet:             'Site master',
      });
    }

    await Site.deleteMany({});
    await Site.insertMany(docs);
    res.json({ success: true, seeded: docs.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.get('/', async (req, res) => {
  try {
    const page   = parseInt(req.query.page)  || 1;
    const limit  = parseInt(req.query.limit) || 50;
    const search = req.query.search || '';
    const query  = search ? {
      $or: [
        { siteName:          { $regex: search, $options: 'i' } },
        { stsId:             { $regex: search, $options: 'i' } },
        { acquisitionPerson: { $regex: search, $options: 'i' } },
      ]
    } : {};
    const [sites, total] = await Promise.all([
      Site.find(query).skip((page - 1) * limit).limit(limit).lean(),
      Site.countDocuments(query),
    ]);
    res.json({ sites, total, page, pages: Math.ceil(total / limit) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;

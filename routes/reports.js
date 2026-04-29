const express = require('express');
const router = express.Router();
const Report = require('../models/Report');

// All reports (list)
router.get('/', async (req, res) => {
  try {
    const reports = await Report.find({}, '-rows')
      .sort({ createdAt: -1 })
      .limit(50)
      .lean();
    res.json(reports);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Single report with all rows
router.get('/:id', async (req, res) => {
  try {
    const report = await Report.findById(req.params.id).lean();
    if (!report) return res.status(404).json({ error: 'Report not found' });
    res.json(report);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Dashboard summary stats
router.get('/dashboard/stats', async (req, res) => {
  try {
    const [total, last5] = await Promise.all([
      Report.aggregate([
        {
          $group: {
            _id: null,
            totalUploads:   { $sum: 1 },
            totalRows:      { $sum: '$totalRows' },
            totalMatched:   { $sum: '$matchedCount' },
            totalUnmatched: { $sum: '$unmatchedCount' },
          },
        },
      ]),
      Report.find({}, '-rows').sort({ createdAt: -1 }).limit(5).lean(),
    ]);

    const stats = total[0] || { totalUploads: 0, totalRows: 0, totalMatched: 0, totalUnmatched: 0 };
    delete stats._id;
    res.json({ stats, recentReports: last5 });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Delete report
router.delete('/:id', async (req, res) => {
  try {
    await Report.findByIdAndDelete(req.params.id);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;

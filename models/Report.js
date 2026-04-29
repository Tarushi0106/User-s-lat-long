const mongoose = require('mongoose');

const matchedRowSchema = new mongoose.Schema({
  rowNumber: Number,
  uploadedLat: Number,
  uploadedLong: Number,
  personName: String,
  siteId: String,
  date: String,
  rawData: mongoose.Schema.Types.Mixed,
  // match result
  matched: Boolean,
  matchedSite: {
    type: mongoose.Schema.Types.ObjectId,
    ref: 'Site',
    default: null,
  },
  matchedSiteId: String,
  matchedSiteName: String,
  matchedLat: Number,
  matchedLong: Number,
  distanceMeters: Number,
  status: {
    type: String,
    enum: ['Work Done - Verified', 'Not Matched'],
    default: 'Not Matched',
  },
});

const reportSchema = new mongoose.Schema({
  fileName: String,
  uploadedBy: String,
  totalRows: Number,
  matchedCount: Number,
  unmatchedCount: Number,
  rows: [matchedRowSchema],
}, { timestamps: true });

module.exports = mongoose.model('Report', reportSchema);

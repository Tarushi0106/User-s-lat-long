const mongoose = require('mongoose');

const siteSchema = new mongoose.Schema({
  stsId:            String,
  siteName:         String,
  circle:           String,
  district:         String,
  updatedStatus:    String,
  lat:              Number,
  long:             Number,
  acquiredLat:      Number,
  acquiredLong:     Number,
  address:          String,
  acquisitionPerson: String,  // "Site acquisition person name" from Site master sheet
  sheet:            String,
}, { timestamps: true });

siteSchema.index({ acquisitionPerson: 1 });
siteSchema.index({ lat: 1, long: 1 });

module.exports = mongoose.model('Site', siteSchema);

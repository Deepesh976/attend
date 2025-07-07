// models/Activity.js
const mongoose = require('mongoose');

const activitySchema = new mongoose.Schema({
  empId: { type: String, required: true },
  empName: { type: String, required: true },
  date: { type: Date, required: true },
  shift: { type: String, default: '' },
  timeInActual: { type: String, default: '00:00:00' },
  timeOutActual: { type: String, default: '00:00:00' },
  lateBy: { type: String, default: '00:00:00' },
  earlyBy: { type: String, default: '00:00:00' },
  ot: { type: String, default: '00:00:00' },
  duration: { type: String, default: '00:00:00' },
  totalDuration: { type: String, default: '00:00:00' },
  status: { type: String, default: 'A' },
  total_present: { type: Number, default: 0 },
  total_absent: { type: Number, default: 0 },
  total_leave: { type: Number, default: 0 },
  total_wo: { type: Number, default: 0 },
  total_ho: { type: Number, default: 0 },
  total_regular_ot: { type: String, default: '00:00:00' }
}, { timestamps: true });

module.exports = mongoose.model('Activity', activitySchema);
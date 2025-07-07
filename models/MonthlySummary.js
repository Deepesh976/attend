const mongoose = require('mongoose');

const monthlySummarySchema = new mongoose.Schema({
  empId: {
    type: String,
    required: true,
    index: true
  },
  empName: {
    type: String,
    required: true
  },
  year: {
    type: Number,
    required: true,
    index: true
  },
  month: {
    type: Number,
    required: true,
    index: true
  },
  monthName: {
    type: String,
    required: true
  },
  totalPresent: {
    type: Number,
    default: 0
  },
  totalAbsent: {
    type: Number,
    default: 0
  },
  totalLeaveTaken: {
    type: Number,
    default: 0
  },
  totalWeeklyOffPresent: {
    type: Number,
    default: 0
  },
  totalDuration: {
    type: String,
    default: '00:00'
  },
  totalTDuration: {
    type: String,
    default: '00:00'
  },
  totalOverTime: {
    type: String,
    default: '00:00'
  },
  totalWOCount: {
    type: Number,
    default: 0
  },
  totalHOCount: {
    type: Number,
    default: 0
  },
  totalLateBy: {
    type: String,
    default: '00:00'
  },
  totalEarlyBy: {
    type: String,
    default: '00:00'
  },
  totalRegularOT: {
    type: String,
    default: '00:00'
  },
  createdAt: {
    type: Date,
    default: Date.now
  },
  updatedAt: {
    type: Date,
    default: Date.now
  }
});

// Create compound index for efficient querying
monthlySummarySchema.index({ empId: 1, year: 1, month: 1 }, { unique: true });

// Update the updatedAt field before saving
monthlySummarySchema.pre('save', function(next) {
  this.updatedAt = new Date();
  next();
});

module.exports = mongoose.model('MonthlySummary', monthlySummarySchema); 
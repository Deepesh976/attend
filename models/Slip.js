const mongoose = require('mongoose');

const slipSchema = new mongoose.Schema({
  empId: { type: String, required: true, trim: true },
  empName: { type: String, required: true, trim: true },
  empUnit: { type: String, trim: true },
  designation: { type: String, trim: true },
  department: { type: String, trim: true },
  dateOfJoining: { type: Date },
  uanNo: { type: String, trim: true },
  esiNo: { type: String, trim: true },
  bankAccountNo: { type: String, trim: true },
  totalDays: { type: Number, default: 0 },
  daysWorked: { type: Number, default: 0 },
  lop: { type: Number, default: 0 },
  annualLeaves: { type: Number, default: 0 },
  plMlBl: { type: Number, default: 0 },
  earnings: {
    basic: { type: Number, default: 0 },
    hra: { type: Number, default: 0 },
    conveyance: { type: Number, default: 0 },
    transportAllowances: { type: Number, default: 0 },
    otherAllowances: { type: Number, default: 0 },
    incentives: { type: Number, default: 0 }
  },
  deductions: {
    esi: { type: Number, default: 0 },
    pf: { type: Number, default: 0 },
    tax: { type: Number, default: 0 },
    gpap: { type: Number, default: 0 },
    otherDeductions: { type: Number, default: 0 },
    lop: { type: Number, default: 0 }
  },
  grossEarnings: { type: Number, default: 0 },
  totalDeductions: { type: Number, default: 0 },
  netSalary: { type: Number, default: 0 },
  month: { type: Number, required: true },
  year: { type: Number, required: true }
}, {
  timestamps: true
});

// Calculate gross earnings, total deductions, and net salary before saving
slipSchema.pre('save', function(next) {
  this.grossEarnings = Object.values(this.earnings).reduce((sum, val) => sum + (val || 0), 0);
  this.totalDeductions = Object.values(this.deductions).reduce((sum, val) => sum + (val || 0), 0);
  this.netSalary = this.grossEarnings - this.totalDeductions;
  next();
});

module.exports = mongoose.model('Slip', slipSchema);

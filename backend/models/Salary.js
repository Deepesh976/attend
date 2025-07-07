const mongoose = require('mongoose');

const salarySchema = new mongoose.Schema({
  year: { type: Number, required: true },
  month: { type: String, required: true },
  monthNumber: { type: Number, required: true },

  empId: { type: String, required: true },
  empName: String,
  department: String,
  designation: String,
  dob: Date,
  doj: Date,

  actualCTCWithoutLOP: { type: Number, default: 0 },
  lopCTC: { type: Number, default: 0 },
  totalDays: { type: Number, default: 0 },
  daysWorked: { type: Number, default: 0 },
  al: { type: Number, default: 0 },
  pl: { type: Number, default: 0 },
  blOrMl: { type: Number, default: 0 },
  lop: { type: Number, default: 0 },
  daysPaid: { type: Number, default: 0 },

  consileSalary: { type: Number, default: 0 },
  basic: { type: Number, default: 0 },
  hra: { type: Number, default: 0 },
  cca: { type: Number, default: 0 },
  transportAllowance: { type: Number, default: 0 },
  otherAllowance1: { type: Number, default: 0 },

  lop2: { type: Number, default: 0 },
  basic3: { type: Number, default: 0 },
  hra4: { type: Number, default: 0 },
  cca5: { type: Number, default: 0 },
  transportAllowance6: { type: Number, default: 0 },
  otherAllowance17: { type: Number, default: 0 },

  grossPay: { type: Number, default: 0 },
  plb: { type: Number, default: 0 },
  pf: { type: Number, default: 0 },
  esi: { type: Number, default: 0 },
  pt: { type: Number, default: 0 },
  tds: { type: Number, default: 0 },
  gpap: { type: Number, default: 0 },
  otherDeductions: { type: Number, default: 0 },
  netPay: { type: Number, default: 0 },

  pfEmployerShare: { type: Number, default: 0 },
  esiEmployerShare: { type: Number, default: 0 },
  bonus: { type: Number, default: 0 }
}, {
  timestamps: true
});

salarySchema.index({ empId: 1, year: 1, month: 1 }, { unique: true });

module.exports = mongoose.model('Salary', salarySchema);
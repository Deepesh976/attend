const mongoose = require('mongoose');

const employeeSchema = new mongoose.Schema({
  empStatus: {
    type: String,
    enum: ['W', 'L'],
    trim: true,
    default: 'W'
  },
  empUnit: { type: String, required: true, trim: true },
  empId: { type: String, required: true, trim: true },
  empName: { type: String, required: true, trim: true },
  dob: { type: Date },
  bloodGroup: { type: String, trim: true },
  doj: { type: Date },
  gender: {
    type: String,
    enum: ['Male', 'Female', 'Other'],
    trim: true
  },
  qualification: { type: String, trim: true },
  experience: { type: String, trim: true },
  personalEmail: {
    type: String,
    trim: true,
    lowercase: true,
    match: [/.+\@.+\..+/, 'Invalid personal email address']
  },
  contactNo: {
    type: String,
    trim: true,
    match: [/^\d{10}$/, 'Invalid contact number']
  },
  department: { type: String, trim: true },
  designation: { type: String, trim: true },
  officialEmail: {
    type: String,
    trim: true,
    lowercase: true,
    match: [/.+\@.+\..+/, 'Invalid official email address']
  },
  panNo: { type: String, trim: true },
  aadharNo: { type: String, trim: true },
  pfNo: { type: String, trim: true },
  uanNo: { type: String, trim: true },
  esiNo: { type: String, trim: true },
  postalAddress: { type: String, trim: true },
  permanentAddress: { type: String, trim: true },
  bankAccount: { type: String, trim: true },
  bankName: { type: String, trim: true },
  ifsc: { type: String, trim: true },
  bankBranch: { type: String, trim: true },
  fatherName: { type: String, trim: true },
  motherName: { type: String, trim: true },
  spouse: { type: String, trim: true },
  nomineeName: { type: String, trim: true },
  emergencyContact: {
    type: String,
    trim: true,
    match: [/^\d{10}$/, 'Invalid emergency contact number']
  },
  exitDate: { type: Date },
  settlementAmount: { type: Number, default: 0 },
  remarks: { type: String, trim: true },
  hiredCtc: { type: Number, default: 0 },
  joiningCtc: { type: Number, default: 0 },
  ctc2025: { type: Number, default: 0 },
  yearsWorked: { type: Number, default: 0 }
}, {
  timestamps: true
});

module.exports = mongoose.model('Employee', employeeSchema);

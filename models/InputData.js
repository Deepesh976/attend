const mongoose = require('mongoose');

const inputDataSchema = new mongoose.Schema({
  EmpID: {
    type: String,
    required: true,
    trim: true
  },
  EmpName: {
    type: String,
    required: true,
    trim: true
  },
  ActualCTCWithoutLossOfPay: {
    type: Number,
    required: false,
    default: 0
  },
  CONSILESALARY: {
    type: Number,
    required: false,
    default: 0
  },
  Basic: {
    type: Number,
    required: false,
    default: 0
  },
  HRA: {
    type: Number,
    required: false,
    default: 0
  },
  CCA: {
    type: Number,
    required: false,
    default: 0
  },
  TRP_ALW: {
    type: Number,
    required: false,
    default: 0
  },
  O_ALW1: {
    type: Number,
    required: false,
    default: 0
  }
}, {
  timestamps: true
});

module.exports = mongoose.model('InputData', inputDataSchema);

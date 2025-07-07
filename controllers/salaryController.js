const Salary = require('../models/Salary');
const Employee = require('../models/Employee');
const InputData = require('../models/InputData');
const MonthlySummary = require('../models/MonthlySummary');
const XLSX = require('xlsx');

const safe = (val) => isNaN(val) ? 0 : Number(val);

const computeDerivedFields = (data) => {
  const totalDays = safe(data.totalDays);
  const daysWorked = safe(data.daysWorked);
  const lop = totalDays - daysWorked;
  const daysPaid = totalDays - lop;

  const consileSalary = safe(data.consileSalary);
  const basic = safe(data.basic);
  const hra = safe(data.hra);
  const cca = safe(data.cca);
  const trp = safe(data.transportAllowance);
  const oalw = safe(data.otherAllowance1);
  const plb = safe(data.plb);
  const tds = safe(data.tds);

  const lop2 = totalDays ? (consileSalary / totalDays) * lop : 0;
  const basic3 = totalDays ? (basic / totalDays) * daysPaid : 0;
  const hra4 = totalDays ? (hra / totalDays) * daysPaid : 0;
  const cca5 = totalDays ? (cca / totalDays) * daysPaid : 0;
  const trp6 = totalDays ? (trp / totalDays) * daysPaid : 0;
  const oalw17 = totalDays ? (oalw / totalDays) * daysPaid : 0;

  const grossPay = basic3 + hra4 + cca5 + trp6 + oalw17;
  const esi = consileSalary > 21000 ? 0 : ((grossPay + plb) * 0.0075);

  let pt = 0;
  if (grossPay > 20000) pt = 200;
  else if (grossPay > 15000) pt = 150;
  else pt = 100;

  const gpap = (consileSalary * 12 * 5 * 0.12) / 100 / 12;
  const otherDeductions = 500;
  const pf = basic3 * 0.12;
  const netPay = (grossPay + plb) - (pf + esi + pt + tds + gpap + otherDeductions);
  const pfEmployerShare = basic3 * 0.12;
  const esiEmployerShare = consileSalary > 21000 ? 0 : ((grossPay + plb) * 0.0325);
  const bonus = grossPay * 0.0833;

  const lopCTC = grossPay + pfEmployerShare + esiEmployerShare + bonus;

  return {
    lop,
    daysPaid,
    lop2,
    basic3,
    hra4,
    cca5,
    transportAllowance6: trp6,
    otherAllowance17: oalw17,
    grossPay,
    pf,
    esi,
    pt,
    gpap,
    otherDeductions,
    netPay,
    pfEmployerShare,
    esiEmployerShare,
    bonus,
    lopCTC
  };
};

const generateSalaryFromEmployee = async (req, res) => {
  try {
    await Salary.deleteMany();

    const employees = await Employee.find();
    const inputData = await InputData.find();
    const summaries = await MonthlySummary.find();

    const inputMap = {};
    inputData.forEach(i => {
      inputMap[(i.EmpID || '').trim().toLowerCase()] = i;
    });

    const allDocs = [];

    for (const emp of employees) {
      const empId = (emp.empId || '').trim();
      const empSummaries = summaries
        .filter(s => (s.empId || '').toLowerCase() === empId.toLowerCase())
        .sort((a, b) => a.year - b.year || a.month - b.month);

      if (!empSummaries.length) continue;

      const input = inputMap[empId.toLowerCase()] || {};
      let carriedAL = 0;

      for (const summary of empSummaries) {
        const month = summary.month;
        const year = summary.year;
        const monthName = summary.monthName || new Date(year, month - 1).toLocaleString('default', { month: 'long' });

        let grantedAL = 1;
        if (month % 4 === 0) grantedAL += 1;
        const totalAL = carriedAL + grantedAL;

        const usedAL = safe(summary.annualLeave || 0);
        const remainingAL = Math.max(totalAL - usedAL, 0);

        const base = {
          empId,
          empName: emp.empName,
          department: emp.department,
          designation: emp.designation,
          dob: emp.dob,
          doj: emp.doj,
          year,
          month: monthName,
          monthNumber: month,
          totalDays: safe(summary.totalDays || 30),
          daysWorked: safe(summary.totalPresent || 0),
          consileSalary: safe(input.CONSILESALARY),
          basic: safe(input.Basic),
          hra: safe(input.HRA),
          cca: safe(input.CCA),
          transportAllowance: safe(input.TRP_ALW),
          otherAllowance1: safe(input.O_ALW1),
          plb: safe(input.PLB),
          tds: safe(input.TDS),
          actualCTCWithoutLOP: safe(input.ActualCTCWithoutLossOfPay || 0),
          annualLeaves: totalAL,
          usedAL,
          carriedAL: remainingAL
        };

        const computed = computeDerivedFields(base);
        allDocs.push({ ...base, ...computed });

        carriedAL = remainingAL;
      }
    }

    if (allDocs.length) {
      await Salary.bulkWrite(
        allDocs.map(doc => ({
          updateOne: {
            filter: { empId: doc.empId, year: doc.year, month: doc.month },
            update: { $set: doc },
            upsert: true
          }
        }))
      );
    }

    res.json({ message: `${allDocs.length} salary records generated with AL carry-forward.` });
  } catch (err) {
    res.status(500).json({ message: 'Failed to generate salary', error: err.message });
  }
};

module.exports = {
  createSalary: async (req, res) => {
    try {
      const data = req.body;
      const computed = computeDerivedFields(data);
      const salary = new Salary({ ...data, ...computed });
      await salary.save();
      res.status(201).json({ message: 'Salary record created', salary });
    } catch (error) {
      res.status(400).json({ error: error.message });
    }
  },

  updateSalary: async (req, res) => {
    try {
      const data = req.body;
      const computed = computeDerivedFields(data);
      const updated = await Salary.findByIdAndUpdate(req.params.id, { ...data, ...computed }, { new: true });
      if (!updated) return res.status(404).json({ error: 'Salary record not found' });
      res.json({ message: 'Salary updated', salary: updated });
    } catch (error) {
      res.status(400).json({ error: error.message });
    }
  },

  getAllSalaries: async (req, res) => {
    try {
      const salaries = await Salary.find();
      const formatted = salaries.map(s => ({
        ...s._doc,
        dob: s.dob ? s.dob.toISOString().split('T')[0] : '',
        doj: s.doj ? s.doj.toISOString().split('T')[0] : ''
      }));
      res.json(formatted);
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  },

  getSalaryById: async (req, res) => {
    try {
      const salary = await Salary.findById(req.params.id);
      if (!salary) return res.status(404).json({ error: 'Salary not found' });
      res.json({
        ...salary._doc,
        dob: salary.dob ? salary.dob.toISOString().split('T')[0] : '',
        doj: salary.doj ? salary.doj.toISOString().split('T')[0] : ''
      });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  },

  deleteSalary: async (req, res) => {
    try {
      const deleted = await Salary.findByIdAndDelete(req.params.id);
      if (!deleted) return res.status(404).json({ error: 'Salary not found' });
      res.json({ message: 'Salary deleted' });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  },

  uploadSalaryExcel: async (req, res) => {
    try {
      if (!req.file || !req.file.buffer) return res.status(400).json({ message: 'No file uploaded.' });

      const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

      const docs = rows.map(row => {
        const base = {
          empId: row['EmpID']?.toString().trim(),
          empName: row['EmpName'],
          department: row['DEPT'],
          designation: row['DESIGNATION'],
          dob: row['DOB'] ? new Date(row['DOB']) : null,
          doj: row['DOJ'] ? new Date(row['DOJ']) : null,
          year: Number(row['Year']),
          month: row['Month'],
          totalDays: safe(row['Total Days']),
          daysWorked: safe(row['Days Worked']),
          consileSalary: safe(row['CONSILE SALARY']),
          basic: safe(row['BASIC']),
          hra: safe(row['HRA']),
          cca: safe(row['CCA']),
          transportAllowance: safe(row['TRP_ALW']),
          otherAllowance1: safe(row['O_ALW1']),
          plb: safe(row['PLB']),
          tds: safe(row['TDS']),
          actualCTCWithoutLOP: safe(row['ACTUAL CTC WITHOUT LOSS OF PAY'] || 0)
        };
        return { ...base, ...computeDerivedFields(base) };
      });

      if (docs.length) {
        await Salary.bulkWrite(
          docs.map(doc => ({
            updateOne: {
              filter: { empId: doc.empId, year: doc.year, month: doc.month },
              update: { $set: doc },
              upsert: true
            }
          }))
        );
      }

      res.json({ message: `${docs.length} salary records uploaded.` });
    } catch (error) {
      res.status(500).json({ message: 'Upload failed', error: error.message });
    }
  },

  generateSalaryFromEmployee
};

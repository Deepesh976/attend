const Slip = require('../models/Slip');
const Employee = require('../models/Employee');
const Salary = require('../models/Salary');
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

// Helper: Flatten object for template replacement
const flattenObject = (obj, parent = '', res = {}) => {
  for (let key in obj) {
    const propName = parent ? `${parent}.${key}` : key;
    if (typeof obj[key] === 'object' && obj[key] !== null) {
      flattenObject(obj[key], propName, res);
    } else {
      res[propName] = obj[key];
    }
  }
  return res;
};

// Create Salary Slip
exports.createSlip = async (req, res) => {
  try {
    const {
      empId, empName, empUnit, designation, department, dateOfJoining, uanNo, esiNo, bankAccountNo,
      totalDays, daysWorked, lop, annualLeaves, plMlBl,
      earnings, deductions
    } = req.body;

    const currentDate = new Date();
    const month = req.body.month || currentDate.getMonth() + 1;
    const year = req.body.year || currentDate.getFullYear();

    let finalEmpUnit = empUnit;
    let finalBankAccountNo = bankAccountNo;

    // 🔍 Fetch additional details from Employee if not provided
    if (empId) {
      const employee = await Employee.findOne({ empId: empId.toLowerCase() });
      if (employee) {
        if (!finalEmpUnit) finalEmpUnit = employee.empUnit || '';
        if (!finalBankAccountNo) finalBankAccountNo = employee.bankAccount || '';
        if (!designation && employee.designation) req.body.designation = employee.designation;
        if (!department && employee.department) req.body.department = employee.department;
        if (!dateOfJoining && employee.doj) req.body.dateOfJoining = employee.doj;
        if (!uanNo && employee.uanNo) req.body.uanNo = employee.uanNo;
        if (!esiNo && employee.esiNo) req.body.esiNo = employee.esiNo;
      }
    }

    // ✅ Log values to verify
    console.log('Saving Slip with Annual Leaves:', annualLeaves);
    console.log('Saving Slip with PL/ML/BL:', plMlBl);

    // 💾 Create and save slip
    const slip = new Slip({
      empId,
      empName,
      empUnit: finalEmpUnit,
      designation: req.body.designation,
      department: req.body.department,
      dateOfJoining: req.body.dateOfJoining,
      uanNo: req.body.uanNo,
      esiNo: req.body.esiNo,
      bankAccountNo: finalBankAccountNo,
      totalDays,
      daysWorked,
      lop,
      annualLeaves,
      plMlBl,
      earnings,
      deductions,
      month,
      year
    });

    await slip.save();

    // 📄 Generate PDF
    const pdfBuffer = await generateSlipPDF(slip);

    const uploadsDir = path.join(__dirname, '../uploads');
    if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });

    const fileName = `salary_slip_${empName.replace(/\s+/g, '_')}_${month}_${year}.pdf`;
    const filePath = path.join(uploadsDir, fileName);
    fs.writeFileSync(filePath, pdfBuffer);

    res.status(201).json({
      message: 'Slip created successfully',
      slip,
      pdfUrl: `/uploads/${fileName}`
    });
  } catch (error) {
    console.error('Error creating slip:', error);
    res.status(400).json({ error: error.message });
  }
};

// Generate PDF
const generateSlipPDF = async (slip) => {
  try {
    const templatePath = path.join(__dirname, '../templates/salary-slip-template.html');
    let htmlTemplate = fs.readFileSync(templatePath, 'utf8');

    const logoPath = path.join(__dirname, '../templates/company_logo.png');
    let logoBase64 = '';
    if (fs.existsSync(logoPath)) {
      const logoBuffer = fs.readFileSync(logoPath);
      logoBase64 = `data:image/png;base64,${logoBuffer.toString('base64')}`;
    }

    const flatSlip = flattenObject({
      logoPath: logoBase64,
      month: getMonthName(slip.month).toUpperCase(),
      year: slip.year,
      empName: slip.empName || '',
      designation: slip.designation || '',
      department: slip.department || '',
      dateOfJoining: slip.dateOfJoining ? new Date(slip.dateOfJoining).toLocaleDateString('en-GB') : '',
      uanNo: slip.uanNo || 'N/A',
      esiNo: slip.esiNo || 'N/A',
      bankAccountNo: slip.bankAccountNo || '',
      totalDays: slip.totalDays || 0,
      daysWorked: slip.daysWorked || 0,
      lop: slip.lop || 0,
      annualLeaves: slip.annualLeaves || 0,
      plMlBl: slip.plMlBl || 0,
      grossEarnings: slip.grossEarnings || 0,
      totalDeductions: slip.totalDeductions || 0,
      netSalary: slip.netSalary || 0,
      netSalaryInWords: numberToWords(slip.netSalary || 0),
      earnings: slip.earnings || {},
      deductions: slip.deductions || {}
    });

    // Replace placeholders
    for (const [key, value] of Object.entries(flatSlip)) {
      const regex = new RegExp(`{{\\s*${key}\\s*}}`, 'g');
      htmlTemplate = htmlTemplate.replace(regex, value);
    }

    const browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();
    await page.setContent(htmlTemplate, { waitUntil: 'networkidle0' });

    const pdfBuffer = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: '20px', right: '20px', bottom: '20px', left: '20px' }
    });

    await browser.close();
    return pdfBuffer;
  } catch (error) {
    console.error('Error generating PDF:', error);
    throw new Error('Failed to generate PDF: ' + error.message);
  }
};

// Month Name Helper
const getMonthName = (monthNumber) => {
  const months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  return months[monthNumber - 1];
};

// Number to Words Helper
const numberToWords = (num) => {
  if (isNaN(num)) return '';
  const integerPart = Math.floor(Math.abs(num));
  const decimalPart = Math.round((Math.abs(num) - integerPart) * 100);
  let words = (num < 0 ? 'Negative ' : '') + convertToWords(integerPart);
  if (decimalPart > 0) {
    words += ' and ' + convertToWords(decimalPart) + ' Paise';
  }
  return words;
};

const convertToWords = (num) => {
  const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
  const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  const convertHundreds = (n) => {
    let result = '';
    if (n >= 100) {
      result += ones[Math.floor(n / 100)] + ' Hundred ';
      n %= 100;
    }
    if (n >= 10 && n <= 19) {
      result += teens[n - 10] + ' ';
    } else if (n >= 20) {
      result += tens[Math.floor(n / 10)] + ' ';
      n %= 10;
    }
    if (n > 0 && n < 10) {
      result += ones[n] + ' ';
    }
    return result.trim();
  };

  let result = '';
  if (num >= 10000000) {
    result += convertHundreds(Math.floor(num / 10000000)) + ' Crore ';
    num %= 10000000;
  }
  if (num >= 100000) {
    result += convertHundreds(Math.floor(num / 100000)) + ' Lakh ';
    num %= 100000;
  }
  if (num >= 1000) {
    result += convertHundreds(Math.floor(num / 1000)) + ' Thousand ';
    num %= 1000;
  }
  if (num > 0) {
    result += convertHundreds(num);
  }

  return result.trim();
};

// Get all slips
exports.getAllSlips = async (req, res) => {
  try {
    const slips = await Slip.find().sort({ createdAt: -1 });
    res.json(slips);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
};

// Download slip PDF
exports.downloadSlipPDF = async (req, res) => {
  try {
    const slip = await Slip.findById(req.params.id);
    if (!slip) return res.status(404).json({ error: 'Slip not found' });

    const pdfBuffer = await generateSlipPDF(slip);
    const fileName = `salary_slip_${slip.empName.replace(/\s+/g, '_')}_${slip.month}_${slip.year}.pdf`;

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.send(pdfBuffer);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
};

// View slip PDF inline
exports.viewSlipPDF = async (req, res) => {
  try {
    const slip = await Slip.findById(req.params.id);
    if (!slip) return res.status(404).json({ error: 'Slip not found' });

    const pdfBuffer = await generateSlipPDF(slip);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'inline');
    res.send(pdfBuffer);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
};

// Get slip by ID
exports.getSlipById = async (req, res) => {
  try {
    const slip = await Slip.findById(req.params.id);
    if (!slip) return res.status(404).json({ error: 'Slip not found' });
    res.json(slip);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
};

// Delete slip
exports.deleteSlip = async (req, res) => {
  try {
    const deleted = await Slip.findByIdAndDelete(req.params.id);
    if (!deleted) return res.status(404).json({ error: 'Slip not found' });
    res.json({ message: 'Slip deleted successfully' });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
};

// Get salary details for form autofill
exports.getEmployeeSalaryDetails = async (req, res) => {
  try {
    const empId = req.params.empId;
    const employee = await Employee.findOne({ empId });
    if (!employee) return res.status(404).json({ error: 'Employee not found' });

    const salary = await Salary.findOne({ empId }).sort({ year: -1, month: -1 });
    if (!salary) return res.status(404).json({ error: 'Salary details not found' });

    const earnings = {
      basic: salary.basic || salary.BASIC || 0,
      hra: salary.hra || salary.HRA || 0,
      conveyance: salary.cca || salary.CCA || 0,
      transportAllowances: salary.trp_alw || salary.transportAllowance || 0,
      otherAllowances: salary.o_alw1 || salary.otherAllowance1 || 0,
      incentives: salary.plb || salary.PLB || 0
    };

    const deductions = {
      esi: salary.esi || salary.ESI || 0,
      pf: salary.pf || salary.PF || 0,
      tax: salary.pt || salary.PT || salary.tds || salary.TDS || 0,
      gpap: salary.gpap || salary.GPAP || 0,
      otherDeductions: salary.otherDeductions || salary.OTH_DEDS || 0,
      lop: salary.lop2 || salary.LOP2 || 0
    };

    // ✅ Include explicitly
    res.json({
      employee,
      salary,
      earnings,
      deductions,
      annualLeaves: salary.al || 0,
      plMlBl: (salary.pl || 0) + (salary.blOrMl || 0)
    });

  } catch (error) {
    res.status(500).json({ error: error.message });
  }
};

const Employee = require('../models/Employee');

// ✅ Upload multiple employees
const uploadEmployees = async (req, res) => {
  try {
    const { employees } = req.body;

    if (!Array.isArray(employees) || employees.length === 0) {
      return res.status(400).json({ message: 'No employee data provided' });
    }

    const validEmployees = [];
    const skippedRows = [];

    for (const emp of employees) {
      // 🔁 Normalize empId to lowercase
      const empId = String(emp.empId || '').trim().toLowerCase();
      const empName = String(emp.empName || '').trim();
      const empUnit = String(emp.empUnit || '').trim();

      if (!empId || !empName || !empUnit) {
        skippedRows.push(emp);
        continue;
      }

      emp.empId = empId;
      emp.empName = empName;
      emp.empUnit = empUnit;

      emp.contactNo = emp.contactNo ? String(emp.contactNo).replace(/\D/g, '').slice(-10) : '';
      emp.emergencyContact = emp.emergencyContact ? String(emp.emergencyContact).replace(/\D/g, '').slice(-10) : '';
      emp.bankAccount = emp.bankAccount ? String(emp.bankAccount).replace(/[^\d]/g, '') : '';

      emp.gender = ['m', 'male'].includes(String(emp.gender).toLowerCase()) ? 'Male'
                 : ['f', 'female'].includes(String(emp.gender).toLowerCase()) ? 'Female'
                 : 'Other';

      // 🔁 Convert blank numeric fields to numbers or leave blank
      const numberFields = ['settlementAmount', 'hiredCtc', 'joiningCtc', 'ctc2025', 'yearsWorked'];
      numberFields.forEach(field => {
        emp[field] = emp[field] === '' ? '' : Number(emp[field]) || '';
      });

      validEmployees.push(emp);
    }

    const finalEmployees = validEmployees.filter(emp => emp.empId && emp.empId !== 'null');

    if (finalEmployees.length === 0) {
      return res.status(400).json({ message: 'No valid rows to insert', skippedCount: skippedRows.length });
    }

    const result = await Employee.insertMany(finalEmployees, { ordered: false });

    return res.status(200).json({
      message: 'Employees uploaded successfully',
      insertedCount: result.length,
      skippedCount: skippedRows.length
    });
  } catch (err) {
    console.error('❌ Upload Error:', err);
    return res.status(500).json({ message: 'Server error while uploading' });
  }
};

// ✅ Get all employees
const getAllEmployees = async (req, res) => {
  try {
    const employees = await Employee.find().sort({ createdAt: -1 });
    res.status(200).json(employees);
  } catch (err) {
    res.status(500).json({ message: 'Server error while fetching employees' });
  }
};

// ✅ Delete employees by empId
const deleteEmployees = async (req, res) => {
  try {
    const { ids } = req.body;
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ message: 'No employee IDs provided' });
    }

    // Normalize input IDs to lowercase
    const normalizedIds = ids.map(id => String(id).trim().toLowerCase());

    const result = await Employee.deleteMany({ empId: { $in: normalizedIds } });
    res.status(200).json({ message: `${result.deletedCount} employees deleted` });
  } catch (err) {
    res.status(500).json({ message: 'Server error while deleting employees' });
  }
};

// ✅ Get employee by MongoDB ID
const getEmployeeById = async (req, res) => {
  try {
    const employee = await Employee.findById(req.params.id);
    if (!employee) {
      return res.status(404).json({ message: 'Employee not found' });
    }
    res.status(200).json(employee);
  } catch (err) {
    res.status(500).json({ message: 'Server error while fetching employee' });
  }
};

// ✅ Update employee by MongoDB ID
const updateEmployeeById = async (req, res) => {
  try {
    if (req.body.empId) {
      req.body.empId = String(req.body.empId).trim().toLowerCase(); // Normalize during update
    }

    const updated = await Employee.findByIdAndUpdate(req.params.id, req.body, { new: true });
    if (!updated) {
      return res.status(404).json({ message: 'Employee not found' });
    }
    res.status(200).json({ message: 'Employee updated', employee: updated });
  } catch (err) {
    res.status(500).json({ message: 'Server error while updating employee' });
  }
};

// ✅ Check if empId already exists
const checkEmpIdExists = async (req, res) => {
  try {
    const { empId } = req.query;
    if (!empId) {
      return res.status(400).json({ exists: false, message: 'empId is required' });
    }

    const normalizedId = String(empId).trim().toLowerCase();
    const employee = await Employee.findOne({ empId: normalizedId });
    res.status(200).json({ exists: !!employee });
  } catch (err) {
    res.status(500).json({ message: 'Error checking empId' });
  }
};

// ✅ Export all functions
module.exports = {
  uploadEmployees,
  getAllEmployees,
  deleteEmployees,
  getEmployeeById,
  updateEmployeeById,
  checkEmpIdExists
};

const express = require('express');
const router = express.Router();
const {
  uploadEmployees,
  getAllEmployees,
  deleteEmployees,
  getEmployeeById,
  updateEmployeeById,
  checkEmpIdExists
} = require('../controllers/employeeController');

// Optional: Validate controller functions (dev-time check)
if (
  typeof uploadEmployees !== 'function' ||
  typeof getAllEmployees !== 'function' ||
  typeof deleteEmployees !== 'function' ||
  typeof getEmployeeById !== 'function' ||
  typeof updateEmployeeById !== 'function' ||
  typeof checkEmpIdExists !== 'function'
) {
  throw new Error('❌ One or more employeeController exports are missing or not functions');
}

// ✅ Routes
router.post('/upload', uploadEmployees);              // Excel → MongoDB
router.get('/', getAllEmployees);                     // Get all employees
router.post('/delete', deleteEmployees);              // Bulk delete by empId
router.get('/check-empid', checkEmpIdExists);         // ✅ Must be before :id route
router.get('/:id', getEmployeeById);                  // Get single by MongoDB _id
router.put('/:id', updateEmployeeById);               // Update single by MongoDB _id

// Delete all employees (for dev/reset purposes)
router.delete('/all', async (req, res) => {
  try {
    const Employee = require('../models/Employee');
    const result = await Employee.deleteMany({});
    res.status(200).json({ message: `${result.deletedCount} employees removed` });
  } catch (err) {
    console.error('❌ Delete All Error:', err.message);
    res.status(500).json({ message: 'Failed to delete all employees' });
  }
});

module.exports = router;

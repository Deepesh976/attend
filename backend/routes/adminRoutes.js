const express = require('express');
const router = express.Router();
const {
  createAdmin,
  getAllAdmins,
  deleteAdmin
} = require('../controllers/superadminController'); // Make sure deleteAdmin is exported there

const { protect, authorizeRoles } = require('../middleware/authMiddleware');

// ✅ Debug Logs (optional)
console.log('✅ createAdmin:', typeof createAdmin);
console.log('✅ getAllAdmins:', typeof getAllAdmins);
console.log('✅ deleteAdmin:', typeof deleteAdmin);

// ✅ Create Admin (POST)
router.post('/create', protect, authorizeRoles('superadmin'), createAdmin);

// ✅ Fetch All Admins (GET)
router.get('/all', protect, authorizeRoles('superadmin'), getAllAdmins);

// ✅ Delete Admin by ID (DELETE)
router.delete('/:id', protect, authorizeRoles('superadmin'), deleteAdmin);

module.exports = router;

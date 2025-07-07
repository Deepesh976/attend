const express = require('express');
const router = express.Router();
const multer = require('multer');
const salaryController = require('../controllers/salaryController');

// Multer setup for file uploads
const upload = multer({ storage: multer.memoryStorage() });

// SALARY ROUTES

router.post('/upload-excel', upload.single('file'), salaryController.uploadSalaryExcel);
router.post('/generate-from-employee', salaryController.generateSalaryFromEmployee);
router.post('/', salaryController.createSalary);
router.get('/', salaryController.getAllSalaries);
router.get('/:id', salaryController.getSalaryById);
router.put('/:id', salaryController.updateSalary);
router.delete('/:id', salaryController.deleteSalary);

module.exports = router;
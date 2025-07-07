const express = require('express');
const router = express.Router();
const slipController = require('../controllers/slipController');

router.post('/', slipController.createSlip);
router.get('/', slipController.getAllSlips);
router.get('/:id', slipController.getSlipById);
router.delete('/:id', slipController.deleteSlip);

// PDF operations
router.get('/download/:id', slipController.downloadSlipPDF);
router.get('/view/:id', slipController.viewSlipPDF);

// New Route to fetch employee & salary details
router.get('/details/:empId', slipController.getEmployeeSalaryDetails);

module.exports = router;

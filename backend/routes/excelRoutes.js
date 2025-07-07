const express = require('express');
const router = express.Router();

const {
  uploadInputData,
  getAllInputData,
  deleteAllInputData,
  clearAllData,
  updateInputDataById,
  deleteInputDataById,
  deleteManyInputData // ✅ FIXED import
} = require('../controllers/inputDataController');

router.post('/upload', uploadInputData);
router.get('/', getAllInputData);
router.put('/:id', updateInputDataById);     // ✅ Edit by ID
router.delete('/:id', deleteInputDataById);  // ✅ Delete by ID
router.delete('/clear', clearAllData);
router.post('/delete-many', deleteManyInputData); // ✅ Bulk delete

module.exports = router;

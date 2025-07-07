const express = require('express');
const router = express.Router();
const {
  uploadInputData,
  getAllInputData,
  deleteAllInputData,
  clearAllData,
  updateInputDataById,
  deleteInputDataById,
  deleteManyInputData // ✅ make sure it's imported
} = require('../controllers/inputDataController');

// Your existing routes
router.post('/upload', uploadInputData);
router.get('/', getAllInputData);
router.delete('/clear', clearAllData);
router.delete('/all', deleteAllInputData);
router.put('/:id', updateInputDataById);
router.delete('/:id', deleteInputDataById);

// ✅ Add this route
router.post('/delete-many', deleteManyInputData);

module.exports = router;

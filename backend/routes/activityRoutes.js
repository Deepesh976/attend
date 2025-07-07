const express = require('express');
const router = express.Router();
const multer = require('multer');
const {
  uploadActivityData,
  getAllActivities,
  deleteAllActivities,
  uploadActivityExcel
} = require('../controllers/activityController');

// Multer configuration for file uploads
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 5 * 1024 * 1024 // 5MB file size limit
  },
  fileFilter: (req, file, cb) => {
    console.log('📁 File upload attempt:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      fieldname: file.fieldname,
      size: file.size
    });
    
    // Check file extension
    const allowedExtensions = /\.(xlsx|xls)$/i;
    const hasValidExtension = allowedExtensions.test(file.originalname);
    
    // Check MIME type - be more permissive for debugging
    const allowedMimeTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
      'application/vnd.ms-excel', // .xls
      'application/excel',
      'application/x-excel',
      'application/x-msexcel',
      'application/octet-stream' // Sometimes Excel files are sent with this MIME type
    ];
    const hasValidMimeType = allowedMimeTypes.includes(file.mimetype);
    
    // For debugging: Accept if either extension OR mimetype is valid
    if (hasValidExtension || hasValidMimeType) {
      console.log('✅ File accepted:', file.originalname);
      return cb(null, true);
    }
    
    console.log('❌ File rejected:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      hasValidExtension,
      hasValidMimeType,
      allowedMimeTypes
    });
    
    cb(new Error(`Only Excel files (.xlsx, .xls) are allowed! Received: ${file.mimetype}`));
  }
});

// Route to upload activity data in JSON format
router.post('/upload', uploadActivityData);

// Route to get all activities with optional filters (empId, startDate, endDate, page, limit)
router.get('/', getAllActivities);

// Route to delete activities based on filters (empIds, startDate, endDate)
router.delete('/', deleteAllActivities);

// Route to upload activity data via Excel file
router.post('/upload-excel', upload.single('file'), uploadActivityExcel);

module.exports = router;

const express = require('express');
const router = express.Router();
const MonthlySummary = require('../models/MonthlySummary');
const { getMonthlySummary } = require('../controllers/activityController'); // ✅ REQUIRED IMPORT

// Get monthly summary by employee ID and optional year/month
router.get('/employee/:empId', async (req, res) => {
  try {
    const { empId } = req.params;
    const { year, month } = req.query;
    
    let filter = { empId };
    
    if (year) {
      filter.year = parseInt(year);
    }
    
    if (month) {
      filter.month = parseInt(month);
    }
    
    const summaries = await MonthlySummary.find(filter)
      .sort({ year: -1, month: -1 });
    
    res.status(200).json({
      success: true,
      count: summaries.length,
      data: summaries
    });
  } catch (error) {
    console.error('❌ Get Monthly Summary Error:', error);
    res.status(500).json({
      success: false,
      message: 'Server error while fetching monthly summary',
      error: error.message
    });
  }
});

// ✅ Dynamic summary generation from activity data
router.get('/monthly-summary', getMonthlySummary);

// Get all monthly summaries with optional filters
router.get('/', async (req, res) => {
  try {
    const { empId, year, month, limit = 50, page = 1 } = req.query;
    
    let filter = {};
    
    if (empId) {
      filter.empId = { $regex: empId, $options: 'i' };
    }
    
    if (year) {
      filter.year = parseInt(year);
    }
    
    if (month) {
      filter.month = parseInt(month);
    }
    
    const skip = (parseInt(page) - 1) * parseInt(limit);
    
    const summaries = await MonthlySummary.find(filter)
      .sort({ year: -1, month: -1, empId: 1 })
      .skip(skip)
      .limit(parseInt(limit));
    
    const totalCount = await MonthlySummary.countDocuments(filter);
    
    res.status(200).json({
      success: true,
      count: summaries.length,
      totalCount,
      currentPage: parseInt(page),
      totalPages: Math.ceil(totalCount / parseInt(limit)),
      data: summaries
    });
  } catch (error) {
    console.error('❌ Get All Monthly Summaries Error:', error);
    res.status(500).json({
      success: false,
      message: 'Server error while fetching monthly summaries',
      error: error.message
    });
  }
});

// Delete monthly summary by ID
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    
    const deletedSummary = await MonthlySummary.findByIdAndDelete(id);
    
    if (!deletedSummary) {
      return res.status(404).json({
        success: false,
        message: 'Monthly summary not found'
      });
    }
    
    res.status(200).json({
      success: true,
      message: 'Monthly summary deleted successfully',
      data: deletedSummary
    });
  } catch (error) {
    console.error('❌ Delete Monthly Summary Error:', error);
    res.status(500).json({
      success: false,
      message: 'Server error while deleting monthly summary',
      error: error.message
    });
  }
});

// Delete all monthly summaries for an employee
router.delete('/employee/:empId', async (req, res) => {
  try {
    const { empId } = req.params;
    
    const result = await MonthlySummary.deleteMany({ empId });
    
    res.status(200).json({
      success: true,
      message: `${result.deletedCount} monthly summaries deleted for employee ${empId}`,
      deletedCount: result.deletedCount
    });
  } catch (error) {
    console.error('❌ Delete Employee Monthly Summaries Error:', error);
    res.status(500).json({
      success: false,
      message: 'Server error while deleting employee monthly summaries',
      error: error.message
    });
  }
});

// Get summary statistics
router.get('/stats', async (req, res) => {
  try {
    const totalSummaries = await MonthlySummary.countDocuments();
    const uniqueEmployees = await MonthlySummary.distinct('empId');
    const years = await MonthlySummary.distinct('year');
    
    res.status(200).json({
      success: true,
      data: {
        totalSummaries,
        uniqueEmployees: uniqueEmployees.length,
        employees: uniqueEmployees,
        years: years.sort((a, b) => b - a)
      }
    });
  } catch (error) {
    console.error('❌ Get Monthly Summary Stats Error:', error);
    res.status(500).json({
      success: false,
      message: 'Server error while fetching monthly summary statistics',
      error: error.message
    });
  }
});

module.exports = router;

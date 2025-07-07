const InputData = require('../models/InputData');

// ✅ Upload and save unique entries from Excel or frontend
const uploadInputData = async (req, res) => {
  try {
    const inputArray = req.body;

    if (!Array.isArray(inputArray) || inputArray.length === 0) {
      return res.status(400).json({ message: 'Invalid or empty data array.' });
    }

    // ✅ Clean and normalize incoming data
    const cleanedData = inputArray.map(item => ({
      EmpID: item.EmpID?.toString().trim(), // ✅ Keep as-is to match Employee.empId
      EmpName: item.EmpName?.toString().trim() || '',
      ActualCTCWithoutLossOfPay: Math.round(Number(item.ActualCTCWithoutLossOfPay || 0)),
      CONSILESALARY: Math.round(Number(item.CONSILESALARY || 0)),
      Basic: Math.round(Number(item.Basic || 0)),
      HRA: Math.round(Number(item.HRA || 0)),
      CCA: Math.round(Number(item.CCA || 0)),
      TRP_ALW: Math.round(Number(item.TRP_ALW || 0)),
      O_ALW1: Math.round(Number(item.O_ALW1 || 0))
    }));

    // ✅ Get existing EmpIDs to avoid duplicates
    const existing = await InputData.find({}, 'EmpID');
    const existingEmpIDs = new Set(existing.map(e => e.EmpID?.toString().trim()));

    // ✅ Only keep unique new rows
    const newUniqueData = cleanedData.filter(row => row.EmpID && !existingEmpIDs.has(row.EmpID));

    // ✅ Insert new unique rows
    if (newUniqueData.length > 0) {
      await InputData.insertMany(newUniqueData);
    }

    // ✅ Return the full updated list
    const updatedAll = await InputData.find();
    res.status(200).json(updatedAll);
  } catch (error) {
    console.error('❌ Upload Error:', error.message);
    res.status(500).json({ message: 'Error saving input data', error: error.message });
  }
};

// ✅ Get all uploaded entries
const getAllInputData = async (req, res) => {
  try {
    const data = await InputData.find();
    res.status(200).json(data);
  } catch (error) {
    res.status(500).json({ message: 'Failed to fetch data', error: error.message });
  }
};

// ✅ Delete all data (for Reset or Remove Upload)
const deleteAllInputData = async (req, res) => {
  try {
    await InputData.deleteMany({});
    res.status(200).json({ message: 'All input data removed' });
  } catch (error) {
    res.status(500).json({ message: 'Failed to delete data', error: error.message });
  }
};

// ✅ Alternative name or reuse
const clearAllData = async (req, res) => {
  try {
    await InputData.deleteMany({});
    res.status(200).json({ message: 'All input data cleared' });
  } catch (err) {
    res.status(500).json({ message: 'Failed to clear data', error: err.message });
  }
};

// ✅ Update a single input data row by _id
const updateInputDataById = async (req, res) => {
  try {
    const updated = await InputData.findByIdAndUpdate(
      req.params.id,
      req.body,
      { new: true }
    );
    if (!updated) {
      return res.status(404).json({ message: 'Data not found' });
    }
    res.status(200).json(updated);
  } catch (error) {
    res.status(500).json({ message: 'Failed to update data', error: error.message });
  }
};

// ✅ Delete a single row by _id
const deleteInputDataById = async (req, res) => {
  try {
    const deleted = await InputData.findByIdAndDelete(req.params.id);
    if (!deleted) {
      return res.status(404).json({ message: 'Data not found' });
    }
    res.status(200).json({ message: 'Row deleted successfully' });
  } catch (error) {
    res.status(500).json({ message: 'Failed to delete row', error: error.message });
  }
};

// ✅ Delete multiple rows by array of _id
const deleteManyInputData = async (req, res) => {
  try {
    const { ids } = req.body;
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ message: 'No IDs provided' });
    }
    await InputData.deleteMany({ _id: { $in: ids } });
    res.status(200).json({ message: 'Deleted selected rows' });
  } catch (err) {
    res.status(500).json({ message: 'Bulk delete failed', error: err.message });
  }
};

// ✅ Export all functions
module.exports = {
  uploadInputData,
  getAllInputData,
  deleteAllInputData,
  clearAllData,
  updateInputDataById,
  deleteInputDataById,
  deleteManyInputData
};

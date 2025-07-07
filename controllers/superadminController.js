const User = require('../models/User');

// ✅ Create Admin
exports.createAdmin = async (req, res) => {
  const { name, email, password } = req.body;

  try {
    const existingUser = await User.findOne({ email });
    if (existingUser) {
      return res.status(409).json({ message: 'Admin already exists' });
    }

    const newUser = new User({ name, email, password, role: 'admin' });
    await newUser.save();

    res.status(201).json({ message: 'Admin created successfully', user: newUser });
  } catch (err) {
    console.error('❌ Create Admin Error:', err);
    res.status(500).json({ message: 'Server error while creating admin' });
  }
};

// ✅ Get All Admins
exports.getAllAdmins = async (req, res) => {
  try {
    const admins = await User.find({ role: 'admin' });
    res.status(200).json(admins);
  } catch (err) {
    console.error('❌ Fetch Admins Error:', err);
    res.status(500).json({ message: 'Server error while fetching admins' });
  }
};

// ✅ Delete Admin by ID
exports.deleteAdmin = async (req, res) => {
  const { id } = req.params;

  try {
    const deletedAdmin = await User.findByIdAndDelete(id);

    if (!deletedAdmin) {
      return res.status(404).json({ message: 'Admin not found' });
    }

    res.status(200).json({ message: 'Admin deleted successfully', user: deletedAdmin });
  } catch (err) {
    console.error('❌ Delete Admin Error:', err);
    res.status(500).json({ message: 'Server error while deleting admin' });
  }
};

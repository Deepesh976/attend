const express = require('express');
const router = express.Router();
const { login, updatePassword } = require('../controllers/authController');
const { protect } = require('../middleware/authMiddleware');

// Register (optional)
router.post('/register', async (req, res) => {
  const { email, password, role } = req.body;

  try {
    const existing = await User.findOne({ email });
    if (existing) return res.status(409).json({ message: 'User already exists' });

    const newUser = new User({ email, password, role });
    await newUser.save();

    res.status(201).json({ message: `${role} registered successfully` });
  } catch (err) {
    res.status(500).json({ message: 'Server error' });
  }
});

// Login route
router.post('/login', login);

// ✅ Secure password update route
router.put('/update-password', protect, updatePassword);

module.exports = router;

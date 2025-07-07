const mongoose = require('mongoose');
const bcrypt = require('bcryptjs');
const dotenv = require('dotenv');
const User = require('./models/User');

dotenv.config();

(async () => {
  try {
    await mongoose.connect(process.env.MONGO_URI);

    // 🧹 Clean old user
    await User.deleteOne({ email: 'superadmin@gmail.com' });

    // ✅ This must be a raw string, not already hashed
    const plainPassword = '123';
    console.log('🔑 Plain password to hash:', plainPassword);

    const hashedPassword = await bcrypt.hash(plainPassword, 10);
    console.log('🔐 Hashed password:', hashedPassword);

    const newUser = new User({
      email: 'superadmin@gmail.com',
      password: hashedPassword,
      role: 'superadmin',
    });

    await newUser.save();
    console.log('✅ Superadmin created with email: superadmin@gmail.com and password: 123');

    mongoose.disconnect();
  } catch (err) {
    console.error('❌ Error in seeding:', err);
  }
})();

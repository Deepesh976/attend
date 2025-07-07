import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { ToastContainer, toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';

const styles = {
  pageWrapper: {
    minHeight: '100vh',
    background: 'linear-gradient(135deg, #f8f9fa, #e9ecef)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '2rem',
  },
  card: {
    background: '#fff',
    padding: '2.5rem 2rem',
    borderRadius: '14px',
    boxShadow: '0 8px 24px rgba(0, 0, 0, 0.1)',
    width: '100%',
    maxWidth: '620px',
    boxSizing: 'border-box',
  },
  header: {
    fontSize: '1.8rem',
    fontWeight: '700',
    color: '#343a40',
    marginBottom: '1.8rem',
    textAlign: 'center',
    letterSpacing: '0.5px',
  },
  formGroup: {
    marginBottom: '1.4rem',
  },
  label: {
    display: 'block',
    marginBottom: '0.4rem',
    fontWeight: '600',
    color: '#495057',
  },
  inputWrapper: {
    position: 'relative',
  },
  input: {
    width: '100%',
    padding: '0.75rem 1rem',
    paddingRight: '4rem',
    borderRadius: '8px',
    border: '1.5px solid #ced4da',
    fontSize: '1rem',
    transition: 'border-color 0.2s',
    boxSizing: 'border-box',
  },
  readOnly: {
    backgroundColor: '#e9ecef',
    color: '#6c757d',
    cursor: 'not-allowed',
  },
  toggleButton: {
    position: 'absolute',
    top: '50%',
    right: '1rem',
    transform: 'translateY(-50%)',
    background: 'none',
    border: 'none',
    color: '#007bff',
    fontSize: '0.9rem',
    fontWeight: '600',
    cursor: 'pointer',
  },
  submitBtn: {
    width: '100%',
    backgroundColor: '#007bff',
    color: '#fff',
    padding: '0.9rem',
    borderRadius: '10px',
    fontWeight: '600',
    fontSize: '1rem',
    border: 'none',
    cursor: 'pointer',
    marginTop: '1.2rem',
    transition: 'background-color 0.3s ease',
  },
  submitBtnHover: {
    backgroundColor: '#0056b3',
  },
};

const Profile = () => {
  const [email, setEmail] = useState('');
  const [form, setForm] = useState({ password: '', confirmPassword: '' });
  const [showPassword, setShowPassword] = useState(false);
  const [showConfirm, setShowConfirm] = useState(false);
  const [btnHover, setBtnHover] = useState(false);

  useEffect(() => {
    const storedEmail = localStorage.getItem('email');
    if (storedEmail) setEmail(storedEmail);
  }, []);

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm((prev) => ({ ...prev, [name]: value }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (form.password !== form.confirmPassword) {
      toast.error('Passwords do not match!');
      return;
    }

    try {
      const token = localStorage.getItem('token');
await axios.put(
  'http://localhost:5000/api/auth/update-password',
  { password: form.password }, // ✅ Only send the new password
  { headers: { Authorization: `Bearer ${token}` } }
);

      toast.success('Password updated successfully!');
      setForm({ password: '', confirmPassword: '' });
    } catch (err) {
      toast.error(err.response?.data?.message || 'Failed to update password');
    }
  };

  return (
    <div style={styles.pageWrapper}>
      <ToastContainer />
      <form style={styles.card} onSubmit={handleSubmit}>
        <h2 style={styles.header}>Account Settings</h2>

        {/* Email */}
        <div style={styles.formGroup}>
          <label htmlFor="email" style={styles.label}>Email</label>
          <input
            type="email"
            id="email"
            value={email}
            readOnly
            style={{ ...styles.input, ...styles.readOnly }}
          />
        </div>

        {/* New Password */}
        <div style={styles.formGroup}>
          <label htmlFor="password" style={styles.label}>New Password</label>
          <div style={styles.inputWrapper}>
            <input
              type={showPassword ? 'text' : 'password'}
              id="password"
              name="password"
              value={form.password}
              onChange={handleChange}
              required
              style={styles.input}
            />
            <button
              type="button"
              onClick={() => setShowPassword((prev) => !prev)}
              style={styles.toggleButton}
            >
              {showPassword ? 'Hide' : 'Show'}
            </button>
          </div>
        </div>

        {/* Confirm Password */}
        <div style={styles.formGroup}>
          <label htmlFor="confirmPassword" style={styles.label}>Confirm Password</label>
          <div style={styles.inputWrapper}>
            <input
              type={showConfirm ? 'text' : 'password'}
              id="confirmPassword"
              name="confirmPassword"
              value={form.confirmPassword}
              onChange={handleChange}
              required
              style={styles.input}
            />
            <button
              type="button"
              onClick={() => setShowConfirm((prev) => !prev)}
              style={styles.toggleButton}
            >
              {showConfirm ? 'Hide' : 'Show'}
            </button>
          </div>
        </div>

        <button
          type="submit"
          style={{
            ...styles.submitBtn,
            ...(btnHover ? styles.submitBtnHover : {}),
          }}
          onMouseEnter={() => setBtnHover(true)}
          onMouseLeave={() => setBtnHover(false)}
        >
          Update Password
        </button>
      </form>
    </div>
  );
};

export default Profile;

import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import axios from 'axios';
import { toast, ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';

const styles = {
  headerBar: {
    width: '100%',
    padding: '0.8rem 2rem',
    background: 'linear-gradient(to right, #3f51b5, #5a55ae)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: '#fff',
    boxShadow: '0 2px 8px rgba(0,0,0,0.12)',
    position: 'fixed',
    top: 0,
    left: 0,
    zIndex: 1000,
  },
  headerTitle: {
    fontSize: '1.4rem',
    fontWeight: 600,
    fontFamily: "'Poppins', sans-serif",
  },
  container: {
    position: 'relative',
    width: '100%',
    height: '100vh',
    background: 'linear-gradient(to right, #e9efff, #f6f8ff)',
    overflow: 'hidden',
    fontFamily: "'Poppins', sans-serif",
    paddingTop: '70px',
  },
  formsContainer: {
    position: 'absolute',
    width: '100%',
    height: 'calc(100% - 150px)',
    top: '70px',
    left: 0,
  },
  signinSignup: {
    position: 'absolute',
    top: '50%',
    left: '72%',
    transform: 'translate(-50%, -50%)',
    width: '40%',
    display: 'grid',
    gridTemplateColumns: '1fr',
    zIndex: 5,
  },
  form: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexDirection: 'column',
    padding: '2rem 3rem',
    borderRadius: '20px',
    backgroundColor: '#fff',
    boxShadow: '0 10px 35px rgba(0,0,0,0.08)',
  },
  logo: {
    width: '250px',
    marginBottom: '1rem',
    objectFit: 'contain',
  },
  title: {
    fontSize: '1.6rem',
    color: '#333',
    marginBottom: '1rem',
  },
  inputField: {
    maxWidth: '380px',
    width: '100%',
    backgroundColor: '#f1f3f6',
    margin: '10px 0',
    height: '50px',
    borderRadius: '50px',
    display: 'flex',
    alignItems: 'center',
    padding: '0 1rem',
    position: 'relative',
  },
  inputIcon: {
    marginRight: '0.8rem',
    color: '#777',
    fontSize: '1.1rem',
    width: '20px',
    textAlign: 'center',
  },
  input: {
    flex: 1,
    background: 'none',
    outline: 'none',
    border: 'none',
    fontWeight: 500,
    fontSize: '0.95rem',
    color: '#333',
  },
  toggleBtn: {
    position: 'absolute',
    right: '1rem',
    background: 'none',
    border: 'none',
    color: '#3f51b5',
    fontWeight: 'bold',
    fontSize: '0.8rem',
    cursor: 'pointer',
  },
  btn: {
    width: '140px',
    backgroundColor: '#3f51b5',
    border: 'none',
    outline: 'none',
    height: '45px',
    borderRadius: '45px',
    color: '#fff',
    textTransform: 'uppercase',
    fontWeight: 600,
    margin: '20px 0 10px 0',
    cursor: 'pointer',
    fontSize: '0.95rem',
    transition: 'background 0.3s',
  },
  btnHover: {
    backgroundColor: '#2c3ea8',
  },
  panelsContainer: {
    position: 'absolute',
    height: '100%',
    width: '100%',
    top: 0,
    left: 0,
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
  },
  panel: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'flex-end',
    justifyContent: 'center',
    textAlign: 'center',
    zIndex: 6,
  },
  leftPanel: {
    pointerEvents: 'all',
    padding: '2rem 10% 1rem 8%',
  },
  image: {
    width: '100%',
    maxWidth: '4180px',
    marginTop: '1rem',
  },
};

const Login = () => {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [btnHover, setBtnHover] = useState(false);
  const navigate = useNavigate();

  const handleSubmit = async (e) => {
    e.preventDefault();

    if (!email || !password) {
      toast.warn('Email and password are required');
      return;
    }

    setLoading(true);
    try {
      const res = await axios.post('http://localhost:5000/api/auth/login', {
        email,
        password,
      });

      const { token, user } = res.data;
      const expiryTime = Date.now() + 60 * 60 * 1000;

      localStorage.setItem('token', token);
      localStorage.setItem('role', user.role);
      localStorage.setItem('email', user.email);
      localStorage.setItem('expiry', expiryTime.toString());

      toast.success('Login successful');
      setTimeout(() => navigate('/dashboard'), 1000);

      setTimeout(() => {
        localStorage.clear();
        alert('Session expired. Please log in again.');
        navigate('/');
      }, 60 * 60 * 1000);
    } catch (err) {
      const errorMessage = err.response?.data?.message || 'Invalid credentials. Please try again.';
      toast.error(errorMessage);
    } finally {
      setLoading(false);
    }
  };

  return (
    <>
      <ToastContainer position="top-right" autoClose={3000} />
      <link
        rel="stylesheet"
        href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600&display=swap"
      />
      <link
        rel="stylesheet"
        href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css"
        integrity="sha512-papmLrJcLYK5kk8e2EAcEQBJb+eIYJq6z5L4Rmv7xFnmM9fCnm56E2O4zGahFvJrOSICMEV+X2r8kX8A9YfRgA=="
        crossOrigin="anonymous"
        referrerPolicy="no-referrer"
      />

      {/* Header */}
      <div style={styles.headerBar}>
        <h1 style={styles.headerTitle}>Employee Management System</h1>
      </div>

      {/* Main Content */}
      <div style={styles.container}>
        <div style={styles.formsContainer}>
          <div style={styles.signinSignup}>
            <form style={styles.form} onSubmit={handleSubmit}>
              <img src="/logo.png" alt="Logo" style={styles.logo} />
              <div style={styles.inputField}>
                <i className="fas fa-envelope" style={styles.inputIcon}></i>
                <input
                  type="email"
                  placeholder="Email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  required
                  style={styles.input}
                />
              </div>

              <div style={styles.inputField}>
                <i className="fas fa-lock" style={styles.inputIcon}></i>
                <input
                  type={showPassword ? 'text' : 'password'}
                  placeholder="Password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  required
                  style={styles.input}
                />
                <button
                  type="button"
                  onClick={() => setShowPassword((prev) => !prev)}
                  style={styles.toggleBtn}
                >
                  {showPassword ? 'Hide' : 'Show'}
                </button>
              </div>

              <button
                type="submit"
                style={{
                  ...styles.btn,
                  ...(btnHover ? styles.btnHover : {}),
                  opacity: loading ? 0.7 : 1,
                  cursor: loading ? 'not-allowed' : 'pointer',
                }}
                onMouseEnter={() => setBtnHover(true)}
                onMouseLeave={() => setBtnHover(false)}
                disabled={loading}
              >
                {loading ? 'Logging in...' : 'Login'}
              </button>
            </form>
          </div>
        </div>

        {/* Left Panel */}
        <div style={styles.panelsContainer}>
          <div style={{ ...styles.panel, ...styles.leftPanel }}>
            <img src="/log.svg" style={styles.image} alt="Login Visual" />
          </div>
        </div>
      </div>
    </>
  );
};

export default Login;

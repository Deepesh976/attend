import React, { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  FaUserCheck,
  FaMoneyBillWave,
  FaDownload,
  FaFileInvoiceDollar // ✅ Added missing icon import
} from 'react-icons/fa';
import AdminNavbar from '../Navbar/adminnavbar';
import SuperNavbar from '../Navbar/supernavbar';

const styles = {
  container: {
    padding: '2rem',
    fontFamily: 'Segoe UI, sans-serif',
    backgroundColor: '#f4f6f8',
    minHeight: '100vh',
  },
  headerBox: {
    textAlign: 'center',
    marginBottom: '3rem',
  },
  biometricTitle: {
    fontSize: '1.1rem',
    fontWeight: '600',
    color: '#007bff',
    marginBottom: '0.5rem',
    letterSpacing: '1px',
    textTransform: 'uppercase',
  },
  headerTitle: {
    fontSize: '2.5rem',
    fontWeight: '700',
    color: '#333',
    marginBottom: '0.5rem',
  },
  headerSubtitle: {
    fontSize: '1.1rem',
    color: '#666',
  },
  cardGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))',
    gap: '2rem',
    maxWidth: '1200px',
    margin: '0 auto',
  },
  card: {
    backgroundColor: '#fff',
    padding: '2rem',
    borderRadius: '12px',
    boxShadow: '0 6px 20px rgba(0,0,0,0.06)',
    cursor: 'pointer',
    transition: 'transform 0.25s ease, box-shadow 0.3s ease',
    textAlign: 'center',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '1rem',
  },
  icon: {
    fontSize: '2.5rem',
    color: '#007bff',
  },
  cardTitle: {
    fontSize: '1.25rem',
    fontWeight: '600',
    color: '#222',
  },
  cardText: {
    fontSize: '0.95rem',
    color: '#555',
  },
};

const Dashboard = () => {
  const navigate = useNavigate();
  const [role, setRole] = useState('');

  useEffect(() => {
    const storedRole = localStorage.getItem('role');
    setRole(storedRole);
  }, []);

  const allCards = [
    {
      title: 'Employee Activity',
      text: 'Track employee attendance and activities.',
      icon: <FaUserCheck style={styles.icon} />,
      route: '/employee-activity',
      roles: ['admin', 'superadmin'],
    },
    {
      title: 'Salary Info',
      text: 'Check salary in detail.',
      icon: <FaMoneyBillWave style={styles.icon} />, // ✅ Clean and professional money icon
      route: '/employee-salary-info',
      roles: ['superadmin'],
    },
    {
      title: 'Generate Slips',
      text: 'Generate salary slips quickly.',
      icon: <FaFileInvoiceDollar style={styles.icon} />,
      route: '/generate-slip',
      roles: ['admin'],
    },
    {
      title: 'Download Slips',
      text: 'Download salary slips securely.',
      icon: <FaDownload style={styles.icon} />,
      route: '/salary-slip',
      roles: ['admin', 'superadmin'],
    },
  ];

  const visibleCards = allCards.filter((card) => card.roles.includes(role));

  const renderNavbar = () => {
    if (role === 'superadmin') return <SuperNavbar />;
    if (role === 'admin') return <AdminNavbar />;
    return null;
  };

  return (
    <>
      {renderNavbar()}
      <div style={styles.container}>
        <div style={styles.headerBox}>
          <div style={styles.biometricTitle}>Attendance</div>
          <h1 style={styles.headerTitle}>Welcome to the Dashboard</h1>
          <p style={styles.headerSubtitle}>Your centralized attendance management system</p>
        </div>

        <div style={styles.cardGrid}>
          {visibleCards.map((card, index) => (
            <div
              key={index}
              style={styles.card}
              onClick={() => navigate(card.route)}
              onMouseEnter={(e) => {
                e.currentTarget.style.transform = 'translateY(-6px)';
                e.currentTarget.style.boxShadow = '0 10px 24px rgba(0,0,0,0.1)';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.transform = 'translateY(0)';
                e.currentTarget.style.boxShadow = '0 6px 20px rgba(0,0,0,0.06)';
              }}
            >
              {card.icon}
              <div style={styles.cardTitle}>{card.title}</div>
              <div style={styles.cardText}>{card.text}</div>
            </div>
          ))}
        </div>
      </div>
    </>
  );
};

export default Dashboard;

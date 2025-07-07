import React, { useState } from 'react';
import { Link, useLocation, useNavigate } from 'react-router-dom';

const styles = {
  navbar: {
    background: 'linear-gradient(90deg, #003366, #0059b3)', // deeper blues
    color: 'white',
    padding: '1rem 2rem',
    display: 'flex',
    alignItems: 'center',
    fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
    position: 'fixed',
    top: 0,
    left: 0,
    right: 0,
    zIndex: 1000,
    boxShadow: '0 3px 10px rgba(0,0,0,0.2)',
    justifyContent: 'space-between',
  },
  brand: {
    position: 'absolute',
    left: '50%',
    transform: 'translateX(-50%)',
    fontSize: '1.6rem',
    fontWeight: '700',
    letterSpacing: '1.5px',
    userSelect: 'none',
  },
  menuIcon: {
    fontSize: '28px',
    cursor: 'pointer',
    userSelect: 'none',
    padding: '6px',
    borderRadius: '6px',
    transition: 'background-color 0.3s',
  },
  menuIconHover: {
    backgroundColor: 'rgba(255,255,255,0.2)',
  },
  avatarContainer: {
    position: 'relative',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  avatar: {
    backgroundColor: '#fff',
    color: '#0059b3',
    borderRadius: '50%',
    width: '38px',
    height: '38px',
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    fontWeight: '700',
    fontSize: '18px',
    boxShadow: '0 0 8px rgba(0,0,0,0.15)',
    userSelect: 'none',
  },
  dropdown: {
    position: 'absolute',
    top: 'calc(100% + 10px)',
    right: 0,
    backgroundColor: '#fff',
    boxShadow: '0 6px 15px rgba(0,0,0,0.25)',
    borderRadius: '10px',
    minWidth: '170px',
    overflow: 'hidden',
    fontSize: '0.95rem',
    zIndex: 1101,
  },
  dropdownItem: {
    padding: '12px 16px',
    color: '#333',
    textAlign: 'left',
    background: '#fff',
    border: 'none',
    width: '100%',
    cursor: 'pointer',
    transition: 'background 0.2s, color 0.2s',
    fontWeight: '500',
  },
  dropdownItemHover: {
    backgroundColor: '#e6f0ff',
    color: '#0059b3',
  },
  logoutItem: {
    color: '#d9534f',
    fontWeight: '700',
  },

  sidebar: {
    position: 'fixed',
    top: 0,
    left: 0,
    height: '100vh',
    width: '280px',
    backgroundColor: '#ffffff',
    padding: '5rem 1.5rem 1.5rem',
    boxShadow: '4px 0 20px rgba(0,0,0,0.12)',
    transform: 'translateX(-100%)',
    transition: 'transform 0.35s ease-in-out',
    zIndex: 1200,
    display: 'flex',
    flexDirection: 'column',
    borderTopRightRadius: '20px',
    borderBottomRightRadius: '20px',
  },
  sidebarVisible: {
    transform: 'translateX(0)',
  },
  overlay: {
    position: 'fixed',
    top: 0,
    left: 0,
    height: '100vh',
    width: '100%',
    backgroundColor: 'rgba(0, 0, 0, 0.25)',
    zIndex: 1100,
  },

  link: {
    backgroundColor: '#f5f7fa',
    color: '#003366',
    border: 'none',
    padding: '14px 20px',
    margin: '0.5rem 0',
    borderRadius: '14px',
    fontSize: '1.1rem',
    cursor: 'pointer',
    textAlign: 'left',
    textDecoration: 'none',
    fontWeight: '600',
    boxShadow: '0 2px 6px rgba(0, 51, 102, 0.15)',
    transition: 'background-color 0.25s ease, color 0.25s ease, box-shadow 0.25s ease',
    userSelect: 'none',
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
  },
  linkHover: {
    backgroundColor: '#0059b3',
    color: '#ffffff',
    boxShadow: '0 4px 14px rgba(0, 89, 179, 0.4)',
  },
  activeLink: {
    backgroundColor: '#004080',
    color: '#fff',
    fontWeight: '700',
    boxShadow: 'inset 4px 0 0 #00264d',
  },
};

const menuItems = [
  { path: '/dashboard', label: 'Dashboard' },
  { path: '/employee-info', label: 'Employee Info' },
  { path: '/employee-salary-info', label: 'Employee Salary Info' },
  { path: '/employee-activity', label: 'Employee Activity' },
  { path: '/generate-slip', label: 'Generate Slip' },
  { path: '/salary-slip', label: 'Salary Slip' },
];

const adminnavbar = () => {
  const [showSidebar, setShowSidebar] = useState(false);
  const [hovered, setHovered] = useState(null);
  const [showDropdown, setShowDropdown] = useState(false);
  const [menuIconHover, setMenuIconHover] = useState(false);

  const location = useLocation();
  const navigate = useNavigate();

  const toggleDropdown = () => setShowDropdown((prev) => !prev);
  const closeDropdown = () => setShowDropdown(false);
  const initials = (localStorage.getItem('email') || 'DP').slice(0, 2).toUpperCase();
  return (
    <>
      <nav style={styles.navbar} role="navigation" aria-label="Main Navigation">
        <div
          style={{
            ...styles.menuIcon,
            ...(menuIconHover ? styles.menuIconHover : {}),
          }}
          onClick={() => setShowSidebar(true)}
          onMouseEnter={() => setMenuIconHover(true)}
          onMouseLeave={() => setMenuIconHover(false)}
          aria-label="Open sidebar menu"
          role="button"
          tabIndex={0}
          onKeyDown={(e) => {
            if (e.key === 'Enter' || e.key === ' ') setShowSidebar(true);
          }}
        >
          &#9776;
        </div>

        <div style={styles.brand}>Attendance Dashboard</div>

        <div
          style={styles.avatarContainer}
          onClick={toggleDropdown}
          aria-haspopup="true"
          aria-expanded={showDropdown}
          tabIndex={0}
          role="button"
          onKeyDown={(e) => {
            if (e.key === 'Enter' || e.key === ' ') toggleDropdown();
          }}
        >
          <div style={styles.avatar} aria-label="User initials">
            {initials}
          </div>
          {showDropdown && (
            <div
              style={styles.dropdown}
              onMouseLeave={closeDropdown}
              role="menu"
              aria-label="User menu"
            >
              <button
                style={styles.dropdownItem}
                onClick={() => {
                  closeDropdown();
                  navigate('/profile');
                }}
                onMouseEnter={(e) =>
                  (e.currentTarget.style.backgroundColor = '#e6f0ff')
                }
                onMouseLeave={(e) =>
                  (e.currentTarget.style.backgroundColor = '#fff')
                }
                role="menuitem"
              >
                Profile
              </button>
              <button
                style={{ ...styles.dropdownItem, ...styles.logoutItem }}
                onClick={() => {
                  closeDropdown();
                  localStorage.removeItem('token');
                  localStorage.removeItem('role');
                  localStorage.removeItem('email'); 
                  localStorage.removeItem('expiry');
                  navigate('/');
                }}
                onMouseEnter={(e) =>
                  (e.currentTarget.style.backgroundColor = '#fce4e4')
                }
                onMouseLeave={(e) =>
                  (e.currentTarget.style.backgroundColor = '#fff')
                }
                role="menuitem"
              >
                Log Out
              </button>
            </div>
          )}
        </div>
      </nav>

      {showSidebar && (
        <div
          style={styles.overlay}
          onClick={() => setShowSidebar(false)}
          aria-hidden="true"
        />
      )}

      <aside
        style={{
          ...styles.sidebar,
          ...(showSidebar ? styles.sidebarVisible : {}),
        }}
        aria-label="Sidebar navigation"
      >
        {menuItems.map((item, idx) => {
          const isActive = location.pathname === item.path;
          return (
            <Link
              key={idx}
              to={item.path}
              onClick={() => setShowSidebar(false)}
              style={{
                ...styles.link,
                ...(hovered === idx ? styles.linkHover : {}),
                ...(isActive ? styles.activeLink : {}),
              }}
              onMouseEnter={() => setHovered(idx)}
              onMouseLeave={() => setHovered(null)}
              tabIndex={0}
              aria-current={isActive ? 'page' : undefined}
            >
              {item.label}
            </Link>
          );
        })}
      </aside>
    </>
  );
};

export default adminnavbar;

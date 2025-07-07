import React, { useEffect, useState } from 'react';
import { BrowserRouter as Router, Routes, Route, useLocation } from 'react-router-dom';

import Login from './components/Auth/Login';
import ProtectedRoute from './components/Auth/ProtectedRoute';
import Dashboard from './components/Dashboard/Dashboard';

import AdminNavbar from './components/Navbar/adminnavbar';
import SuperNavbar from './components/Navbar/supernavbar';

import EmployeeActivity from './biometric/EmployeeActivity';
import GenerateSlip from './biometric/GenerateSlip';
import SalarySlip from './biometric/SalarySlip';
import EmployeeInfo from './biometric/EmployeeInfo';
import AddEmployeeInfo from './biometric/AddEmployeeInfo';
import EditEmployeeInfo from './biometric/EditEmployeeInfo';
import EmployeeSalaryInfo from './biometric/EmployeeSalaryInfo';
import EditEmployeeSalaryInfo from './biometric/EditEmployeeSalaryInfo';
import Profile from './biometric/Profile';
import InputData from './biometric/InputData';
import CreateAdmin from './biometric/CreateAdmin';

import LoadingButton from './components/Auth/LoadingButton';
import LoadingSpinner from './components/Auth/LoadingSpinner';

const AppContent = () => {
  const location = useLocation();
  const [userRole, setUserRole] = useState(null);

  const hideNavbar = location.pathname === '/' || location.pathname === '/login';

  useEffect(() => {
    const role = localStorage.getItem('role');
    setUserRole(role);
  }, [location]);

  return (
    <>
      {!hideNavbar && (
        userRole === 'superadmin' ? <SuperNavbar /> :
        userRole === 'admin' ? <AdminNavbar /> :
        null
      )}

      <Routes>
        <Route path="/" element={<Login />} />
        <Route path="/login" element={<Login />} />
        <Route path="/dashboard" element={<ProtectedRoute allowedRoles={['admin', 'superadmin']}><Dashboard /></ProtectedRoute>} />
        <Route path="/employee-activity" element={<ProtectedRoute><EmployeeActivity /></ProtectedRoute>} />
        <Route path="/generate-slip" element={<ProtectedRoute><GenerateSlip /></ProtectedRoute>} />
        <Route path="/salary-slip" element={<ProtectedRoute><SalarySlip /></ProtectedRoute>} />
        <Route path="/employee-info" element={<ProtectedRoute><EmployeeInfo /></ProtectedRoute>} />
        <Route path="/add-employee-info" element={<ProtectedRoute><AddEmployeeInfo /></ProtectedRoute>} />
        <Route path="/edit-employee-info/:id" element={<ProtectedRoute><EditEmployeeInfo /></ProtectedRoute>} />
        <Route path="/profile" element={<ProtectedRoute><Profile /></ProtectedRoute>} />
        <Route path="/employee-salary-info" element={<ProtectedRoute><EmployeeSalaryInfo /></ProtectedRoute>} />
        <Route path="/edit-employee-salary-info/:id" element={<ProtectedRoute><EditEmployeeSalaryInfo /></ProtectedRoute>} />
        <Route path="/input-data" element={<ProtectedRoute><InputData /></ProtectedRoute>} />
        <Route path="/create-admin" element={<ProtectedRoute allowedRoles={['superadmin']}><CreateAdmin /></ProtectedRoute>} />
        <Route path="/loading_button" element={<ProtectedRoute><LoadingButton /></ProtectedRoute>} />
        <Route path="/loading_spinner" element={<ProtectedRoute><LoadingSpinner /></ProtectedRoute>} />
      </Routes>
    </>
  );
};

const App = () => (
  <Router>
    <AppContent />
  </Router>
);

export default App;

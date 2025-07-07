// src/components/Auth/ProtectedRoute.js
import React from 'react';
import { Navigate } from 'react-router-dom';

const ProtectedRoute = ({ children, allowedRoles }) => {
  const token = localStorage.getItem('token');
  const userRole = localStorage.getItem('role');
  const expiry = localStorage.getItem('expiry');

  // Check if token exists
  if (!token) return <Navigate to="/" replace />;

  // Check if token is expired
  if (expiry && Date.now() > parseInt(expiry)) {
    localStorage.clear();
    return <Navigate to="/" replace />;
  }

  // Check if role is allowed
  if (allowedRoles && !allowedRoles.includes(userRole)) {
    return <Navigate to="/unauthorized" replace />;
  }

  return children;
};

export default ProtectedRoute;

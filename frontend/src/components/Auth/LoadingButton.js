import React from 'react';
import LoadingSpinner from './LoadingSpinner';

const buttonStyle = {
  width: '100%',
  padding: '0.75rem',
  backgroundColor: '#007bff',
  color: '#fff',
  border: 'none',
  borderRadius: '8px',
  fontSize: '1rem',
  cursor: 'pointer',
  display: 'flex',
  justifyContent: 'center',
  alignItems: 'center',
  gap: '8px',
};

const LoadingButton = ({ loading, onClick, children }) => {
  return (
    <button style={buttonStyle} onClick={onClick} disabled={loading}>
      {loading && <LoadingSpinner />}
      {children}
    </button>
  );
};

export default LoadingButton;

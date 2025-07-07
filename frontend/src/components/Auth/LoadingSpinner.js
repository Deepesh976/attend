import React from 'react';

const spinnerStyle = {
  width: '18px',
  height: '18px',
  border: '3px solid #fff',
  borderTop: '3px solid transparent',
  borderRadius: '50%',
  animation: 'spin 1s linear infinite',
};

const styleTag = (
  <style>
    {`
      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }
    `}
  </style>
);

const LoadingSpinner = () => (
  <>
    {styleTag}
    <div style={spinnerStyle}></div>
  </>
);

export default LoadingSpinner;

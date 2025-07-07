import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './main'; // now valid because default export exists

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);

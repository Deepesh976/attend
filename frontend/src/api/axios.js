// src/api/axios.js
import axios from 'axios';

const instance = axios.create({
  // In development, this will use the proxy setting in package.json
  // In production with Docker, the API will be available at /api
  baseURL: '/api',
  withCredentials: true, // Optional: if you're using cookies/auth
});

export default instance;

import axios from 'axios';

const CEM_API_URL = process.env.CEM_API_URL || 'http://backend-cem:3001/api';
const CEM_USERNAME = process.env.CEM_USERNAME || 'admin';
const CEM_PASSWORD = process.env.CEM_PASSWORD || 'password';

let authToken = null;
let tokenExpiry = null;

const cemClient = axios.create({
  baseURL: CEM_API_URL,
  timeout: 10000,
  headers: { 'Content-Type': 'application/json' }
});

async function login() {
  try {
    const response = await axios.post(`${CEM_API_URL.replace('/api', '')}/api/auth/login`, {
      username: CEM_USERNAME,
      password: CEM_PASSWORD
    });
    authToken = response.data.access_token;
    tokenExpiry = Date.now() + (2 * 60 * 60 * 1000);
    console.log('✅ CEM API login successful');
    return authToken;
  } catch (error) {
    console.error('❌ CEM API login failed:', error.message);
    return null;
  }
}

async function getToken() {
  if (!authToken || Date.now() >= tokenExpiry) {
    await login();
  }
  return authToken;
}

export async function getCEMData(endpoint) {
  try {
    const token = await getToken();
    if (!token) return null;
    const response = await cemClient.get(endpoint, {
      headers: { Authorization: `Bearer ${token}` }
    });
    console.log(`✅ CEM API (${endpoint}): ${Array.isArray(response.data) ? response.data.length + ' items' : 'ok'}`);
    return response.data;
  } catch (error) {
    if (error.response?.status === 401) {
      authToken = null;
      return getCEMData(endpoint);
    }
    console.error(`❌ CEM API Error (${endpoint}):`, error.message);
    return null;
  }
}

// Users API
export async function getUsers() {
  return getCEMData('/users');
}

// Tasks/Projects API
export async function getTasks() {
  return getCEMData('/tasks');
}

export async function getTaskById(id) {
  return getCEMData(`/tasks/${id}`);
}

// Task Steps API
export async function getTaskSteps(taskId) {
  return getCEMData(`/task-steps/task/${taskId}`);
}

// Daily Work Records API
export async function getDailyWork(params = {}) {
  const query = new URLSearchParams(params).toString();
  return getCEMData(`/daily-work${query ? '?' + query : ''}`);
}

// Leave API
export async function getLeaveRequests() {
  return getCEMData('/leave');
}

export async function getLeaveTypes() {
  return getCEMData('/leave/leave-types');
}

export async function getLeaveQuota(userId) {
  return getCEMData(`/leave/quota/${userId}`);
}

export async function getHolidays() {
  return getCEMData('/leave/holidays');
}

// Car Booking API
export async function getCarBookings() {
  return getCEMData('/car-booking');
}

// Settings API
export async function getCategories() {
  return getCEMData('/settings/categories');
}

export async function getStatuses() {
  return getCEMData('/settings/statuses');
}

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
  const idStr = String(id).trim();
  const numOnly = idStr.replace(/^SO/i, '');
  const tasks = await getTasks();
  
  const task = tasks?.find(t => 
    t.id == idStr ||
    t.so_number == idStr || 
    t.so_number == numOnly || 
    t.so_number == `SO${numOnly}` ||
    t.task_number == idStr ||
    t.task_number == numOnly ||
    t.task_name?.toLowerCase().includes(idStr.toLowerCase())
  );
  return task || null;
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

// File cache - แชร์กันทุก user
const fileCache = new Map();
const CACHE_TTL = 24 * 60 * 60 * 1000; // 1 วัน

// PDF to Image conversion
async function convertPdfToImages(pdfBuffer, startPage = 1, endPage = null, maxPages = 50) {
  const { fromBuffer } = await import('pdf2pic');
  
  const converter = fromBuffer(pdfBuffer, {
    density: 300,
    format: 'png',
    width: 2480,
    height: 3508
  });
  
  const images = [];
  const maxEndPage = endPage || (startPage + maxPages - 1);
  
  for (let i = startPage; i <= maxEndPage; i++) {
    try {
      const result = await converter(i, { responseType: 'base64' });
      if (result?.base64) {
        images.push(result.base64);
      } else {
        break;
      }
    } catch (e) {
      break;
    }
  }
  
  const lastPage = startPage + images.length - 1;
  const hasMore = images.length === (maxEndPage - startPage + 1);
  console.log(`📄 Converted pages ${startPage}-${lastPage} (${images.length} pages)${hasMore ? ' - may have more' : ''}`);
  
  return { images, totalPages: lastPage, lastPage, hasMore };
}

// File download - returns base64 (with cache & PDF conversion)
export async function downloadFile(filename, startPage = 1, endPage = null) {
  const cacheKey = `${filename}_${startPage}_${endPage || 'end'}`;
  const cached = fileCache.get(cacheKey);
  if (cached && Date.now() - cached.cachedAt < CACHE_TTL) {
    console.log(`📦 Cache hit: ${filename} (pages ${startPage}-${endPage || 'end'})`);
    return cached.data;
  }

  try {
    const encodedFilename = encodeURIComponent(filename);
    const url = `${CEM_API_URL.replace('/api', '')}/uploads/${encodedFilename}`;
    const response = await axios.get(url, { responseType: 'arraybuffer', timeout: 60000 });
    
    const isPdf = filename.toLowerCase().endsWith('.pdf');
    let result;

    if (isPdf) {
      console.log(`🔄 Converting PDF to images: ${filename} (pages ${startPage}-${endPage || 'end'})`);
      const { images, totalPages, lastPage, hasMore } = await convertPdfToImages(response.data, startPage, endPage);
      if (images.length > 0) {
        result = { 
          base64: images[0],
          allPages: images,
          mimeType: 'image/png', 
          size: response.data.length,
          pageCount: totalPages,
          pagesConverted: images.length,
          startPage,
          endPage: lastPage,
          hasMore,
          nextPage: hasMore ? lastPage + 1 : null
        };
        console.log(`✅ Converted PDF: ${filename} (pages ${startPage}-${lastPage}/${totalPages})`);
      } else {
        result = { base64: Buffer.from(response.data).toString('base64'), mimeType: 'application/pdf', size: response.data.length, pageCount: 1 };
      }
    } else {
      const mimeType = filename.endsWith('.jpg') || filename.endsWith('.jpeg') ? 'image/jpeg'
        : filename.endsWith('.png') ? 'image/png' : 'application/octet-stream';
      result = { base64: Buffer.from(response.data).toString('base64'), mimeType, size: response.data.length, pageCount: 1 };
      console.log(`✅ Downloaded: ${filename} (${Math.round(response.data.length/1024)}KB)`);
    }

    fileCache.set(cacheKey, { data: result, cachedAt: Date.now() });
    return result;
  } catch (error) {
    console.error(`❌ File download error:`, error.message);
    return null;
  }
}

export function clearFileCache() {
  fileCache.clear();
}

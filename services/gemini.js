// services/gemini.service.js

import { GoogleGenerativeAI } from '@google/generative-ai';
import { calendarFunction, createEventFunction, findAvailableTimeFunction, systemInstruction } from '../config/gent_config.js';
import * as cemAPI from './cemAPI.js';

const genAICache = new Map();

function getGenAI(apiKey) {
    if (!genAICache.has(apiKey)) {
        genAICache.set(apiKey, new GoogleGenerativeAI(apiKey));
    }
    return genAICache.get(apiKey);
}

async function withTimeout(promise, timeoutMs = 120000) {
    return Promise.race([
        promise,
        new Promise((_, reject) => 
            setTimeout(() => reject(new Error('Request timeout')), timeoutMs)
        )
    ]);
}

async function withRetry(fn, maxRetries = 2) {
    for (let i = 0; i < maxRetries; i++) {
        try {
            return await fn();
        } catch (error) {
            if (i === maxRetries - 1) throw error;
            await new Promise(r => setTimeout(r, 1000 * (i + 1)));
        }
    }
}

async function getCEMContext(userMessage) {
    const lowerMsg = userMessage.toLowerCase();
    let context = '';

    // พนักงาน
    if (lowerMsg.includes('พนักงาน') || lowerMsg.includes('user') || lowerMsg.includes('คน') || lowerMsg.includes('ทีม') || lowerMsg.includes('แผนก')) {
        const users = await cemAPI.getUsers();
        if (users) context += `\n\n[ข้อมูลพนักงาน - ทั้งหมด ${users.length} คน]\n${JSON.stringify(users)}\n`;
    }

    // โครงการ/งาน
    if (lowerMsg.includes('โครงการ') || lowerMsg.includes('งาน') || lowerMsg.includes('task') || lowerMsg.includes('project') || lowerMsg.includes('so')) {
        const tasks = await cemAPI.getTasks();
        if (tasks) context += `\n\n[ข้อมูลโครงการ - ทั้งหมด ${tasks.length} โครงการ]\n${JSON.stringify(tasks)}\n`;
    }

    // การลา
    if (lowerMsg.includes('ลา') || lowerMsg.includes('leave') || lowerMsg.includes('หยุด') || lowerMsg.includes('พักร้อน') || lowerMsg.includes('ลาป่วย')) {
        const leaves = await cemAPI.getLeaveRequests();
        if (leaves) context += `\n\n[ข้อมูลการลา - ทั้งหมด ${leaves.length} รายการ]\n${JSON.stringify(leaves)}\n`;
    }

    // การจองรถ
    if (lowerMsg.includes('รถ') || lowerMsg.includes('car') || lowerMsg.includes('booking') || lowerMsg.includes('จอง') || lowerMsg.includes('ยืม')) {
        const bookings = await cemAPI.getCarBookings();
        if (bookings) context += `\n\n[ข้อมูลการจองรถ - ทั้งหมด ${bookings.length} รายการ]\n${JSON.stringify(bookings)}\n`;
    }

    // บันทึกการทำงาน/Timesheet
    if (lowerMsg.includes('timesheet') || lowerMsg.includes('บันทึก') || lowerMsg.includes('ชั่วโมง') || lowerMsg.includes('ทำงาน') || lowerMsg.includes('daily')) {
        const dailyWork = await cemAPI.getDailyWork();
        if (dailyWork) context += `\n\n[บันทึกการทำงาน - ทั้งหมด ${dailyWork.length} รายการ]\n${JSON.stringify(dailyWork.slice(0, 50))}\n`;
    }

    // วันหยุด
    if (lowerMsg.includes('วันหยุด') || lowerMsg.includes('holiday') || lowerMsg.includes('ปฏิทิน')) {
        const holidays = await cemAPI.getHolidays();
        if (holidays) context += `\n\n[วันหยุดราชการ]\n${JSON.stringify(holidays)}\n`;
    }

    return context;
}

export async function getGeminiResponse(apiKey, modelName, history) {
    return withRetry(async () => {
        const genAI = getGenAI(apiKey);
        
        // เพิ่ม CEM context ถ้ามีคำถามเกี่ยวข้อง
        const lastMessage = history[history.length - 1];
        let cemContext = '';
        if (lastMessage?.parts?.[0]?.text) {
            cemContext = await getCEMContext(lastMessage.parts[0].text);
            if (cemContext) {
                console.log('📊 CEM Context added to message');
                lastMessage.parts[0].text += cemContext;
            }
        }

        // สร้าง systemInstruction ใหม่ที่รวม CEM info
        const cemSystemInstruction = {
            parts: [{
                text: systemInstruction.parts[0].text + `

---

### **CEM System Integration (ระบบจัดการพนักงาน GenT-CEM):**
คุณสามารถเข้าถึงข้อมูลจากระบบ CEM (Company Employee Management) ได้ ซึ่งรวมถึง:

**1. พนักงาน (Users):**
- id, username, firstname, lastname, email, phone
- position (ตำแหน่ง), department (แผนก), employee_id (รหัสพนักงาน)
- role (admin/user/hr), is_active

**2. โครงการ (Tasks/Projects):**
- id, task_name (ชื่อโครงการ), so_number (เลข SO), contract_number
- sale_owner (ผู้ดูแลการขาย), customer_info (ลูกค้า)
- project_start_date, project_end_date, status, category
- description, files

**3. บันทึกการทำงาน (Daily Work Records/Timesheet):**
- id, task_id, step_id, user_id, work_date
- start_time, end_time, total_hours (ชั่วโมงทำงาน)
- work_status, location, work_description
- employee_name, task_name, step_name

**4. การลา (Leave Requests):**
- id, user_id, user_name, leave_type (ประเภทการลา)
- start_datetime, end_datetime, total_days
- reason, status (pending/approved/rejected)
- has_delegation, delegate_name

**5. การจองรถ (Car Bookings):**
- id, user_id, name (ผู้จอง), type (ประเภท)
- location (ปลายทาง), project, selected_date, time
- license (ทะเบียนรถ), status (pending/active/completed/cancelled)
- return_date, return_time, fuel_level_borrow, fuel_level_return

**6. วันหยุด (Holidays):**
- id, name, date

**วิธีตอบคำถาม:**
- เมื่อผู้ใช้ถามเกี่ยวกับข้อมูลเหล่านี้ ข้อมูลจะถูกแนบมาในข้อความ
- ให้ใช้ข้อมูลนั้นตอบคำถามอย่างถูกต้องและครบถ้วน
- ถ้าถามจำนวน ให้นับจากข้อมูลที่ได้รับ
- ถ้าถามรายละเอียด ให้แสดงข้อมูลที่เกี่ยวข้อง
- ถ้าข้อมูลไม่เพียงพอ ให้บอกว่าไม่มีข้อมูลในส่วนนั้น
`
            }]
        };

        const isGemini3 = modelName.includes('gemini-3');
        
        const modelConfig = {
            model: modelName,
            tools: [{ functionDeclarations: [calendarFunction, createEventFunction, findAvailableTimeFunction] }],
            systemInstruction: cemSystemInstruction
        };

        // Gemini 3 ต้องเปิด thinking config สำหรับ function calling
        if (isGemini3) {
            modelConfig.generationConfig = {
                thinkingConfig: { thinkingBudget: 1024 }
            };
        }

        const model = genAI.getGenerativeModel(modelConfig);

        const result = await withTimeout(
            model.generateContent({ contents: history }),
            120000
        );
        return result.response;
    }, 2);
}

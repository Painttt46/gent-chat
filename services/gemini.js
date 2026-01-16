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
    
    // ตรวจจับชื่อคนในข้อความ
    const users = await cemAPI.getUsers();
    let targetUser = null;
    if (users) {
        for (const u of users) {
            const names = [u.firstname, u.lastname, u.username?.split('@')[0]].filter(Boolean).map(n => n.toLowerCase());
            if (names.some(n => lowerMsg.includes(n))) {
                targetUser = u;
                break;
            }
        }
    }

    // ถ้าถามเกี่ยวกับคนใดคนหนึ่ง + โครงการ/งาน -> ดึง daily_work ของคนนั้น
    if (targetUser && (lowerMsg.includes('โครงการ') || lowerMsg.includes('งาน') || lowerMsg.includes('ทำ'))) {
        const dailyWork = await cemAPI.getDailyWork();
        if (dailyWork) {
            const userWork = dailyWork.filter(w => w.user_id === targetUser.id || w.employee_name?.includes(targetUser.firstname));
            const uniqueTasks = [...new Map(userWork.map(w => [w.task_id, { task_id: w.task_id, task_name: w.task_name, total_hours: userWork.filter(x => x.task_id === w.task_id).reduce((sum, x) => sum + (x.total_hours || 0), 0) }])).values()];
            context += `\n\n[โครงการที่ ${targetUser.firstname} ${targetUser.lastname} ทำ - ${uniqueTasks.length} โครงการ]\n${JSON.stringify(uniqueTasks)}\n`;
            context += `\n[ข้อมูล User: ${targetUser.firstname} ${targetUser.lastname}, ID: ${targetUser.id}]\n`;
        }
        return context;
    }

    // พนักงาน
    if (lowerMsg.includes('พนักงาน') || lowerMsg.includes('user') || lowerMsg.includes('คน') || lowerMsg.includes('ทีม') || lowerMsg.includes('แผนก')) {
        if (users) context += `\n\n[ข้อมูลพนักงาน - ${users.length} คน]\n${JSON.stringify(users.map(u => ({ id: u.id, name: `${u.firstname} ${u.lastname}`, position: u.position, department: u.department, phone: u.phone })))}\n`;
    }

    // โครงการ/งาน (ทั่วไป)
    if (lowerMsg.includes('โครงการ') || lowerMsg.includes('task') || lowerMsg.includes('project') || lowerMsg.match(/so\d+/)) {
        const tasks = await cemAPI.getTasks();
        if (tasks) context += `\n\n[โครงการทั้งหมด - ${tasks.length} โครงการ]\n${JSON.stringify(tasks.map(t => ({ id: t.id, name: t.task_name, so: t.so_number, status: t.status, customer: t.customer_info })))}\n`;
    }

    // การลา
    if (lowerMsg.includes('ลา') || lowerMsg.includes('leave') || lowerMsg.includes('หยุด')) {
        const leaves = await cemAPI.getLeaveRequests();
        if (leaves) context += `\n\n[การลา - ${leaves.length} รายการ]\n${JSON.stringify(leaves)}\n`;
    }

    // การจองรถ
    if (lowerMsg.includes('รถ') || lowerMsg.includes('car') || lowerMsg.includes('จอง')) {
        const bookings = await cemAPI.getCarBookings();
        if (bookings) context += `\n\n[การจองรถ - ${bookings.length} รายการ]\n${JSON.stringify(bookings)}\n`;
    }

    // บันทึกการทำงาน
    if (lowerMsg.includes('timesheet') || lowerMsg.includes('บันทึก') || lowerMsg.includes('ชั่วโมง') || lowerMsg.includes('daily')) {
        const dailyWork = await cemAPI.getDailyWork();
        if (dailyWork) context += `\n\n[บันทึกการทำงาน - ${dailyWork.length} รายการ]\n${JSON.stringify(dailyWork.slice(0, 30))}\n`;
    }

    // วันหยุด
    if (lowerMsg.includes('วันหยุด') || lowerMsg.includes('holiday')) {
        const holidays = await cemAPI.getHolidays();
        if (holidays) context += `\n\n[วันหยุด]\n${JSON.stringify(holidays)}\n`;
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

**วิธีตอบคำถาม CEM (สำคัญมาก!):**
- **กรองข้อมูลตามคำถาม:** ถ้าถามว่า "วีรภัทร ทำโครงการอะไรบ้าง" ให้ดูจาก Daily Work Records ว่า user_id หรือ employee_name ตรงกับ "วีรภัทร" แล้วดึงเฉพาะ task_name ที่เขาทำ ไม่ใช่แสดงโครงการทั้งหมด
- **ใช้ Daily Work เป็นหลัก:** เมื่อถามว่าใครทำโครงการอะไร ให้ดูจาก Daily Work Records เพราะมี user_id และ task_id ที่เชื่อมโยงกัน
- **ถ้าถามจำนวน:** ให้นับเฉพาะที่ตรงกับเงื่อนไข
- **ถ้าถามรายละเอียด:** ให้แสดงเฉพาะข้อมูลที่เกี่ยวข้องกับคำถาม
- **ถ้าข้อมูลไม่เพียงพอ:** ให้บอกว่าไม่มีข้อมูลในส่วนนั้น
- **อย่าแสดงข้อมูลทั้งหมด:** ให้กรองและสรุปเฉพาะที่เกี่ยวข้องกับคำถามเท่านั้น
`
            }]
        };

        const isGemini3 = modelName.includes('gemini-3') || modelName.includes('thinking');
        
        // Gemini 3 ต้องใช้ REST API โดยตรงเพราะ SDK ยังไม่รองรับ thought signatures
        if (isGemini3) {
            console.log(`🔄 Using REST API for ${modelName}`);
            const response = await fetch(
                `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`,
                {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        contents: history,
                        tools: [{ functionDeclarations: [calendarFunction, createEventFunction, findAvailableTimeFunction] }],
                        systemInstruction: cemSystemInstruction,
                        generationConfig: {
                            thinkingConfig: { thinkingBudget: 2048 }
                        }
                    })
                }
            );
            const data = await response.json();
            console.log(`📥 Gemini 3 response status: ${response.status}`);
            if (data.error) {
                console.error(`❌ Gemini 3 error:`, data.error);
                throw new Error(`Gemini API Error: ${data.error.message}`);
            }
            
            // ดึง content ดิบออกมาทั้งก้อน (รวม thought + functionCall)
            const rawContent = data.candidates?.[0]?.content;
            console.log(`✅ Gemini 3 response received, parts: ${rawContent?.parts?.length || 0}`);
            
            return {
                // ส่ง response ตัวเต็มกลับไป (เพื่อให้ save history ได้ครบ)
                rawContent,
                text: () => rawContent?.parts?.find(p => p.text)?.text || '',
                functionCalls: () => rawContent?.parts?.filter(p => p.functionCall).map(p => p.functionCall) || null
            };
        }

        const modelConfig = {
            model: modelName,
            tools: [{ functionDeclarations: [calendarFunction, createEventFunction, findAvailableTimeFunction] }],
            systemInstruction: cemSystemInstruction
        };

        const model = genAI.getGenerativeModel(modelConfig);

        const result = await withTimeout(
            model.generateContent({ contents: history }),
            120000
        );
        return result.response;
    }, 2);
}

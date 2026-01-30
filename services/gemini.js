// services/gemini.service.js

import { GoogleGenerativeAI } from '@google/generative-ai';
import { calendarFunction, systemInstruction } from '../config/gent_config.js';
import * as cemAPI from './cemAPI.js';

const genAICache = new Map();

function getGenAI(apiKey) {
    if (!genAICache.has(apiKey)) {
        genAICache.set(apiKey, new GoogleGenerativeAI(apiKey));
    }
    return genAICache.get(apiKey);
}

// Function declaration สำหรับอ่านไฟล์โครงการ
const readProjectFileFunction = {
    name: "read_project_file",
    description: "อ่านไฟล์เอกสารที่แนบกับโครงการ เช่น สัญญา ใบสั่งซื้อ เอกสารโครงการ รองรับ PDF และรูปภาพ สามารถเลือกช่วงหน้าที่ต้องการอ่านได้ **ถ้าถามต่อเรื่องเอกสารเดิมที่เพิ่งอ่าน ไม่ต้องระบุ taskId ระบบจะใช้โครงการล่าสุดอัตโนมัติ**",
    parameters: {
        type: "object",
        properties: {
            taskId: { type: "number", description: "ID ของโครงการที่ต้องการดูไฟล์ (ไม่ต้องระบุถ้าถามต่อจากเอกสารเดิม)" },
            fileIndex: { type: "number", description: "ลำดับไฟล์ที่ต้องการอ่าน (เริ่มจาก 0) ถ้าไม่ระบุจะอ่านไฟล์แรก" },
            startPage: { type: "number", description: "หน้าเริ่มต้นที่ต้องการอ่าน (เริ่มจาก 1) ถ้าไม่ระบุจะเริ่มจากหน้า 1" },
            endPage: { type: "number", description: "หน้าสุดท้ายที่ต้องการอ่าน ถ้าไม่ระบุจะอ่านถึงหน้าสุดท้าย (สูงสุด 50 หน้าต่อครั้ง)" }
        },
        required: []
    }
};

// CEM API Functions
const getDailyWorkFunction = {
    name: "get_daily_work_records",
    description: "ดึงข้อมูลบันทึกการทำงานประจำวัน (timesheet) - ใช้เมื่อถามว่าใครทำงานอะไร, ทำโครงการอะไร, ลงงานวันไหน, ทำงานกี่ชั่วโมง",
    parameters: { type: "object", properties: {}, required: [] }
};

const getUsersFunction = {
    name: "get_users",
    description: "ดึงข้อมูลพนักงานทั้งหมด - ใช้เมื่อถามเกี่ยวกับพนักงาน, มีใครบ้าง, กี่คน, ตำแหน่งอะไร, แผนกไหน, เบอร์โทร, email",
    parameters: { type: "object", properties: {}, required: [] }
};

const getTasksFunction = {
    name: "get_tasks",
    description: "ดึงข้อมูลโครงการทั้งหมด - ใช้เมื่อถามเกี่ยวกับโครงการ, SO number, ลูกค้า, สถานะโครงการ, วันเริ่ม-สิ้นสุด",
    parameters: { type: "object", properties: {}, required: [] }
};

const getLeaveRequestsFunction = {
    name: "get_leave_requests",
    description: "ดึงข้อมูลการลาทั้งหมด - ใช้เมื่อถามว่าใครลา, ลาวันไหน, ลาประเภทอะไร, สถานะอนุมัติ, วันนี้ใครลา",
    parameters: { type: "object", properties: {}, required: [] }
};

const getPendingLeavesFunction = {
    name: "get_pending_leaves",
    description: "ดึงข้อมูลการลาที่รอดำเนินการ/รออนุมัติ สำหรับผู้อนุมัติ - ใช้เมื่อถามว่ามีใครรอให้อนุมัติ, มีกี่คนรออนุมัติ, การลาที่รอดำเนินการ, pending leave",
    parameters: { type: "object", properties: {}, required: [] }
};

const getCarBookingsFunction = {
    name: "get_car_bookings",
    description: "ดึงข้อมูลการจองรถทั้งหมด - ใช้เมื่อถามเกี่ยวกับการจองรถ, ใครจองรถ, ไปไหน, วันไหน, ทะเบียนอะไร",
    parameters: { type: "object", properties: {}, required: [] }
};

const cemFunctions = [readProjectFileFunction, getDailyWorkFunction, getUsersFunction, getTasksFunction, getLeaveRequestsFunction, getPendingLeavesFunction, getCarBookingsFunction];

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
        
        // ❌ ปิด Pre-fetching - ให้ AI ตัดสินใจเรียก Tool เองเมื่อจำเป็น
        // const lastMessage = history[history.length - 1];
        // let cemContext = '';
        // if (lastMessage?.parts?.[0]?.text) {
        //     cemContext = await getCEMContext(lastMessage.parts[0].text);
        //     if (cemContext) {
        //         console.log('📊 CEM Context added to message');
        //         lastMessage.parts[0].text += cemContext;
        //     }
        // }

        // สร้าง systemInstruction ใหม่ที่รวม CEM info + วันที่ปัจจุบัน (dynamic)
        const currentDate = new Date().toLocaleDateString('th-TH', { timeZone: 'Asia/Bangkok', year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' });
        const currentDateISO = new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Bangkok' }); // YYYY-MM-DD
        
        const cemSystemInstruction = {
            parts: [{
                text: systemInstruction.parts[0].text.replace(
                    /\*\*วันที่ปัจจุบัน:\*\* .+/,
                    `**วันที่ปัจจุบัน:** ${currentDate} (${currentDateISO})`
                ) + `

---

### **CEM System Integration (ระบบจัดการพนักงาน GenT-CEM):**
คุณสามารถเข้าถึงข้อมูลจากระบบ CEM (Company Employee Management) ผ่าน Function Calls ต่อไปนี้:

**Available CEM Functions:**
1. \`get_users\` - ดึงข้อมูลพนักงานทั้งหมด (ชื่อ, ตำแหน่ง, แผนก, email, เบอร์โทร)
2. \`get_tasks\` - ดึงข้อมูลโครงการทั้งหมด (ชื่อโครงการ, เลข SO, ลูกค้า, สถานะ)
3. \`get_daily_work_records\` - ดึงบันทึกการทำงานประจำวัน (timesheet, ชั่วโมงทำงาน, โครงการที่ทำ)
4. \`get_leave_requests\` - ดึงข้อมูลการลาทั้งหมด (ประเภทการลา, วันที่, สถานะอนุมัติ)
5. \`get_pending_leaves\` - ดึงการลาที่รอดำเนินการ/รออนุมัติ (เฉพาะที่ผู้ถามมีสิทธิ์อนุมัติ)
6. \`get_car_bookings\` - ดึงข้อมูลการจองรถ (ผู้จอง, ปลายทาง, วันที่, สถานะ)
7. \`read_project_file\` - อ่านไฟล์เอกสารโครงการ (PDF, รูปภาพ)

**CRITICAL RULES สำหรับคำถาม CEM:**
- **ถ้าถามเกี่ยวกับพนักงาน** (ใครบ้าง, กี่คน, ตำแหน่งอะไร) → เรียก \`get_users\`
- **ถ้าถามเกี่ยวกับโครงการ** (มีโครงการอะไร, SO เท่าไหร่, ลูกค้าใคร) → เรียก \`get_tasks\`
- **ถ้าถามว่าใครทำงานอะไร/ทำโครงการอะไร** → เรียก \`get_daily_work_records\` แล้วกรองตามชื่อ
- **ถ้าถามเกี่ยวกับการลา** (ใครลา, ลาวันไหน, สถานะการลา) → เรียก \`get_leave_requests\`
- **ถ้าถามเกี่ยวกับการลาที่รออนุมัติ/pending** → เรียก \`get_pending_leaves\`
- **ถ้าถามเกี่ยวกับรถ/การจองรถ** → เรียก \`get_car_bookings\`
- **ถ้าถามเกี่ยวกับเอกสารโครงการ** → เรียก \`read_project_file\`

**วิธีตอบคำถาม CEM:**
1. เรียก function ที่เกี่ยวข้องก่อนเสมอ
2. กรองข้อมูลตามคำถาม (ไม่แสดงทั้งหมด)
3. สรุปเป็นภาษาไทยที่เข้าใจง่าย
4. ถ้าถามจำนวน ให้นับและตอบเป็นตัวเลข
5. ถ้าไม่มีข้อมูล ให้บอกว่า "ไม่พบข้อมูล..."

**ตัวอย่างคำถาม CEM:**
- "มีพนักงานกี่คน" → เรียก get_users แล้วนับ
- "วีรภัทร ทำโครงการอะไรบ้าง" → เรียก get_daily_work_records แล้วกรองตามชื่อ วีรภัทร
- "วันนี้ใครลาบ้าง" → เรียก get_leave_requests แล้วกรองวันที่วันนี้
- "มีใครรอให้อนุมัติการลาไหม" → เรียก get_pending_leaves
- "มีใครจองรถวันนี้ไหม" → เรียก get_car_bookings แล้วกรองวันที่
- "โครงการ SO25001 มีรายละเอียดอะไร" → เรียก get_tasks แล้วหา SO25001
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
                        contents: history.map(msg => {
                            if (msg.parts) return msg; // keep rawContent as-is (includes thoughtSignature)
                            return { role: msg.role, parts: [{ text: msg.text || '' }] };
                        }),
                        tools: [{ functionDeclarations: [calendarFunction, ...cemFunctions] }],
                        systemInstruction: cemSystemInstruction,
                        generationConfig: {
                            thinkingConfig: { thinkingBudget: 4096 }
                        }
                    })
                }
            );
            const data = await response.json();
            console.log(`📥 Gemini 3 response status: ${response.status}`);
            if (data.error) {
                console.error(`❌ Gemini 3 error:`, data.error);
                const err = new Error(`Gemini API Error: ${data.error.message}`);
                err.status = data.error.code;
                throw err;
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
            tools: [{ functionDeclarations: [calendarFunction, ...cemFunctions] }],
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

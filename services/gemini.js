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

async function withTimeout(promise, timeoutMs = 30000) {
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

    if (lowerMsg.includes('พนักงาน') || lowerMsg.includes('user') || lowerMsg.includes('คน')) {
        const users = await cemAPI.getUsers();
        if (users) context += `\n\nข้อมูลพนักงาน: ${JSON.stringify(users.slice(0, 10))}\n`;
    }

    if (lowerMsg.includes('โครงการ') || lowerMsg.includes('งาน') || lowerMsg.includes('task')) {
        const tasks = await cemAPI.getTasks();
        if (tasks) context += `\n\nข้อมูลโครงการ: ${JSON.stringify(tasks.slice(0, 10))}\n`;
    }

    if (lowerMsg.includes('ลา') || lowerMsg.includes('leave')) {
        const leaves = await cemAPI.getLeaveRequests();
        if (leaves) context += `\n\nข้อมูลการลา: ${JSON.stringify(leaves.slice(0, 10))}\n`;
    }

    if (lowerMsg.includes('รถ') || lowerMsg.includes('car') || lowerMsg.includes('booking')) {
        const bookings = await cemAPI.getCarBookings();
        if (bookings) context += `\n\nข้อมูลการจองรถ: ${JSON.stringify(bookings.slice(0, 10))}\n`;
    }

    return context;
}

export async function getGeminiResponse(apiKey, modelName, history) {
    return withRetry(async () => {
        const genAI = getGenAI(apiKey);
        
        // เพิ่ม CEM context ถ้ามีคำถามเกี่ยวข้อง
        const lastMessage = history[history.length - 1];
        if (lastMessage?.parts?.[0]?.text) {
            const cemContext = await getCEMContext(lastMessage.parts[0].text);
            if (cemContext) {
                lastMessage.parts[0].text += cemContext;
            }
        }

        const model = genAI.getGenerativeModel({
            model: modelName,
            tools: [{ functionDeclarations: [calendarFunction, createEventFunction, findAvailableTimeFunction] }],
            systemInstruction: systemInstruction + '\n\nคุณสามารถเข้าถึงข้อมูลจากระบบ CEM (Company Employee Management) ได้ รวมถึงข้อมูลพนักงาน, โครงการ, การลา, และการจองรถ'
        });

        const result = await withTimeout(
            model.generateContent({ contents: history }),
            30000
        );
        return result.response;
    }, 2);
}

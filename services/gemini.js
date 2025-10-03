// services/gemini.service.js

import { GoogleGenerativeAI } from '@google/generative-ai';
import { calendarFunction, createEventFunction, findAvailableTimeFunction, systemInstruction } from '../config/gent_config.js';

export async function getGeminiResponse(apiKey, modelName, history) {
    const genAI = new GoogleGenerativeAI(apiKey);
    const model = genAI.getGenerativeModel({
        model: modelName,
        tools: [{ functionDeclarations: [calendarFunction, createEventFunction, findAvailableTimeFunction] }],
        systemInstruction: systemInstruction
    });

    const result = await model.generateContent({
        contents: history
    });
    return result.response;
}
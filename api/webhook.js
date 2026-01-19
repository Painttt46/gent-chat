// api/webhook.js

import * as stateService from '../services/state.js';
import * as graphService from '../services/graphAPI.js';
import * as teamsService from '../services/teams.js';
import * as geminiService from '../services/gemini.js';
import * as cemAPI from '../services/cemAPI.js';

// Try calling Gemini with auto model switch on quota error
async function callGeminiWithFallback(apiKey, model, history, userId) {
  const modelKeys = Object.keys(stateService.models);
  let currentModelIndex = modelKeys.indexOf(model);
  let lastError;
  let currentKey = apiKey;

  // Try with current key first, then switch key if all models fail
  for (let keyAttempt = 0; keyAttempt < 2; keyAttempt++) {
    for (let i = 0; i < modelKeys.length; i++) {
      const tryModel = modelKeys[(currentModelIndex + i) % modelKeys.length];
      try {
        const response = await geminiService.getGeminiResponse(currentKey, tryModel, history);
        if (i > 0 || keyAttempt > 0) stateService.userModels.set(userId, tryModel);
        return { response, model: tryModel, switched: i > 0 || keyAttempt > 0 };
      } catch (err) {
        lastError = err;
        if (err.status === 429 || err.message?.includes('429') || err.message?.includes('quota')) {
          console.log(`⚠️ ${tryModel} quota exceeded, trying next model...`);
          continue;
        }
        throw err;
      }
    }
    // All models failed with current key, try switching key
    if (keyAttempt === 0 && process.env.GEMINI_API_KEY_2) {
      console.log(`🔄 All models quota exceeded, switching API key...`);
      stateService.switchApiKey();
      currentKey = stateService.getCurrentApiKey();
    }
  }
  throw lastError;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    stateService.checkDailyReset();
    const userId = req.body?.from?.id || req.body?.channelData?.tenant?.id || 'default';
    const userName = req.body?.from?.name || '';
    let currentModel = stateService.userModels.get(userId) || 'gemini-3-flash-preview';

    let cleanText = (req.body?.text || '').replace(/<at>.*?<\/at>/g, '').replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/\s+/g, ' ').trim();

    // เพิ่ม user context ถ้ามี และข้อความพูดถึง "ฉัน" หรือ "ของฉัน"
    if (userName && (cleanText.includes('ฉัน') || cleanText.includes('ของฉัน') || cleanText.includes('ผม') || cleanText.includes('ของผม'))) {
      cleanText = cleanText + ` [ผู้ถามคือ: ${userName}]`;
    }

    if (cleanText.toLowerCase() === 'clear') {
      stateService.conversations.delete(userId);
      return res.status(200).json({ text: "🔄 Conversation cleared!" });
    }

    if (cleanText.toLowerCase() === 'model') {
      const modelList = Object.entries(stateService.models).map(([key, model]) => 
        `${key === currentModel ? '✅' : '•'} ${key}: ${model.name} (${model.count}/${model.limit})`
      ).join('\n');
      return res.status(200).json({ text: `🤖 **Available Models:**\n${modelList}` });
    }

    if (cleanText.toLowerCase().startsWith('model ')) {
      const modelKey = cleanText.toLowerCase().replace('model ', '');
      if (stateService.models[modelKey]) {
        stateService.userModels.set(userId, modelKey);
        return res.status(200).json({ text: `🤖 Switched to ${stateService.models[modelKey].name}` });
      } else {
        const modelList = Object.entries(stateService.models).map(([key, model]) => `• ${key}: ${model.name}`).join('\n');
        return res.status(200).json({ text: `❌ Invalid model. Available:\n${modelList}` });
      }
    }

    if (!cleanText) {
      return res.status(200).json({});
    }

    const isBroadcastCommand = cleanText.toLowerCase().startsWith('/broadcast');
    const finalText = isBroadcastCommand ? cleanText.substring(10).trim() : cleanText;

    if (!stateService.conversations.has(userId)) {
      stateService.conversations.set(userId, []);
    }
    const history = stateService.conversations.get(userId);
    const conversationHistory = [...history, { role: "user", parts: [{ text: finalText }] }];

    const { response: geminiResponse, model: usedModel, switched } = await callGeminiWithFallback(
      stateService.getCurrentApiKey(), currentModel, conversationHistory, userId
    );
    currentModel = usedModel;
    stateService.models[currentModel].count++;

    let text;
    const functionCalls = geminiResponse.functionCalls();
    console.log(`🔧 Function calls: ${functionCalls?.length || 0}`, functionCalls ? JSON.stringify(functionCalls[0]?.name) : 'none');
    
    // สำหรับ Gemini 3: save rawContent ทั้งก้อนเพื่อรักษา thought signatures
    const isGemini3 = currentModel.includes('gemini-3');
    
    if (functionCalls && functionCalls.length > 0) {
      const call = functionCalls[0];
      let functionResult;

      switch (call.name) {
        case "get_user_calendar":
          functionResult = await graphService.getUserCalendar(call.args.userPrincipalName, call.args.startDate, call.args.endDate);
          break;
        case "find_available_time":
          functionResult = await graphService.findAvailableTime(call.args);
          break;
        case "create_calendar_event":
          functionResult = await graphService.createCalendarEvent(call.args);
          break;
        case "read_project_file":
          const task = await cemAPI.getTaskById(call.args.taskId);
          if (!task || !task.files?.length) {
            functionResult = { error: "ไม่พบไฟล์ในโครงการนี้" };
          } else {
            const fileIndex = call.args.fileIndex || 0;
            const filename = task.files[fileIndex];
            if (!filename) {
              functionResult = { error: `ไม่พบไฟล์ลำดับที่ ${fileIndex}`, availableFiles: task.files };
            } else {
              const fileData = await cemAPI.downloadFile(filename);
              if (!fileData) {
                functionResult = { error: "ไม่สามารถดาวน์โหลดไฟล์ได้" };
              } else {
                // เก็บ file data ไว้ส่งเป็น part แยก
                functionResult = { 
                  filename, 
                  taskName: task.task_name,
                  message: "กรุณาวิเคราะห์ไฟล์ที่แนบมาพร้อมนี้ รวมถึงรูปภาพและตารางในเอกสาร",
                  _fileData: fileData // เก็บไว้ใช้ข้างล่าง
                };
              }
            }
          }
          break;
        case "get_daily_work_records":
          functionResult = await cemAPI.getDailyWork(call.args || {});
          break;
        case "get_users":
          functionResult = await cemAPI.getUsers();
          break;
        case "get_tasks":
          functionResult = await cemAPI.getTasks();
          break;
        case "get_leave_requests":
          functionResult = await cemAPI.getLeaveRequests();
          break;
        case "get_car_bookings":
          functionResult = await cemAPI.getCarBookings();
          break;
        default:
          functionResult = { error: "Unknown function called." };
      }

      // สำหรับ Gemini 3: ใช้ rawContent ทั้งก้อน
      const modelPart = isGemini3 && geminiResponse.rawContent 
        ? geminiResponse.rawContent 
        : { role: "model", parts: [{ functionCall: call }] };

      // 1. Function Response Message
      const functionMsg = {
        role: "user",
        parts: [{ 
          functionResponse: { 
            name: call.name, 
            response: { result: { ...functionResult, _fileData: undefined } }
          } 
        }]
      };

      // 2. User Message สำหรับส่งไฟล์ (ถ้ามี)
      let fileMsg = null;
      if (functionResult._fileData) {
        fileMsg = {
          role: "user",
          parts: [{
            inlineData: {
              mimeType: functionResult._fileData.mimeType,
              data: functionResult._fileData.base64
            }
          }]
        };
      }

      // 3. ประกอบ History
      const historyWithFunction = [
        ...conversationHistory,
        modelPart,
        functionMsg
      ];
      if (fileMsg) historyWithFunction.push(fileMsg);

      let currentHistory = historyWithFunction;
      let currentResponse;
      let maxLoops = 3;
      
      while (maxLoops-- > 0) {
        const { response: loopResponse } = await callGeminiWithFallback(
          stateService.getCurrentApiKey(), currentModel, currentHistory, userId
        );
        currentResponse = loopResponse;
        
        const loopFunctionCalls = loopResponse.functionCalls();
        if (!loopFunctionCalls || loopFunctionCalls.length === 0) break;
        
        const loopCall = loopFunctionCalls[0];
        console.log(`🔧 Additional function call: "${loopCall.name}"`);
        let loopResult;
        
        switch (loopCall.name) {
          case "get_user_calendar":
            loopResult = await graphService.getUserCalendar(loopCall.args.userPrincipalName, loopCall.args.startDate, loopCall.args.endDate);
            break;
          case "find_available_time":
            loopResult = await graphService.findAvailableTime(loopCall.args);
            break;
          case "create_calendar_event":
            loopResult = await graphService.createCalendarEvent(loopCall.args);
            break;
          case "get_daily_work_records":
            loopResult = await cemAPI.getDailyWork(loopCall.args || {});
            break;
          case "get_users":
            loopResult = await cemAPI.getUsers();
            break;
          case "get_tasks":
            loopResult = await cemAPI.getTasks();
            break;
          case "get_leave_requests":
            loopResult = await cemAPI.getLeaveRequests();
            break;
          case "get_car_bookings":
            loopResult = await cemAPI.getCarBookings();
            break;
          default:
            loopResult = { error: "Unknown function" };
        }
        
        const loopModelPart = isGemini3 && loopResponse.rawContent 
          ? loopResponse.rawContent 
          : { role: "model", parts: [{ functionCall: loopCall }] };
        
        currentHistory = [
          ...currentHistory,
          loopModelPart,
          { role: "user", parts: [{ functionResponse: { name: loopCall.name, response: { result: loopResult } } }] }
        ];
      }
      
      console.log(`🔍 finalResponse keys: ${Object.keys(currentResponse || {})}`);
      
      text = currentResponse.text();
      if (!text && currentResponse.rawContent?.parts) {
        const textParts = currentResponse.rawContent.parts.filter(p => p.text);
        text = textParts.map(p => p.text).join('\n');
      }
      console.log(`📝 After function call text: ${text?.substring(0, 100)}`);
    } else {
      text = geminiResponse.text();
      if (!text && geminiResponse.rawContent?.parts) {
        const textParts = geminiResponse.rawContent.parts.filter(p => p.text);
        text = textParts.map(p => p.text).join('\n');
      }
    }

    console.log(`📝 Response text length: ${text?.length || 0}`);
    
    if (!text) {
      console.log('⚠️ Empty text response');
      return res.status(200).json({ text: '❌ ไม่ได้รับการตอบกลับจาก AI' });
    }

    const isCardFormat = text.startsWith('FORMAT:CARD');
    let cleanResponse = text.replace(/FORMAT:(CARD|TEXT)/, '').trim() || "I'm sorry, I couldn't generate a proper response.";
    console.log(`🧹 Clean response length: ${cleanResponse.length}`);

    try {
      history.push({ role: "user", parts: [{ text: finalText }] });
      history.push({ role: "model", parts: [{ text: cleanResponse }] });
      if (history.length > 40) history.splice(0, 2);
    } catch (e) {
      console.error('History push error:', e);
    }

    const modelInfo = stateService.models[currentModel] || { name: currentModel, count: 0, limit: 20 };
    const switchNote = switched ? ` | ⚡ Auto-switched` : '';
    const usageStats = `💬 ${Math.floor(history.length / 2)} msgs | ${modelInfo.name} | ${modelInfo.count}/${modelInfo.limit}${switchNote}`;
    console.log(`📈 usageStats: ${usageStats}`);
    console.log(`🎴 isCardFormat: ${isCardFormat}`);

    if (isBroadcastCommand) {
      await teamsService.sendToTeamsWebhook(`🔊 **Announcement:**\n\n${cleanResponse}\n\n${usageStats}`);
      return res.status(200).json({ text: "📢 Broadcast sent!" });
    }

    if (isCardFormat) {
      console.log(`📤 Sending card response...`);
      return res.status(200).json({
        type: "message",
        attachments: [{
          contentType: "application/vnd.microsoft.card.adaptive",
          content: {
            type: "AdaptiveCard", version: "1.2",
            body: [
              { type: "TextBlock", text: "🤖 Gent - Work Assistant", weight: "Bolder", size: "Medium", color: "Accent" },
              { type: "TextBlock", text: cleanResponse, wrap: true, spacing: "Medium" },
              { type: "TextBlock", text: usageStats, size: "Small", color: "Good", weight: "Bolder", spacing: "Medium" }
            ]
          }
        }]
      });
    } else {
      console.log(`📤 Sending text response...`);
      return res.status(200).json({ text: `🤖 **Gent:** ${cleanResponse}\n\n${usageStats}` });
    }

  } catch (error) {
    console.error('Handler Error:', error);
    let errorMsg;
    if (error.message === 'Request timeout') {
      errorMsg = '⏱️ Request timeout. Please try again.';
    } else if (error.status === 429 || error.message?.includes('quota')) {
      errorMsg = '⚠️ All models quota exceeded! Please try again tomorrow or type `model` to check status.';
    } else {
      errorMsg = `❌ **Gent Error:** ${error.message}`;
    }
    res.status(500).json({ text: errorMsg });
  }
}

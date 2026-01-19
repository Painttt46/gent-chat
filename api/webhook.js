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
          console.log(`🔍 read_project_file args:`, JSON.stringify(call.args));
          // ถ้าไม่ระบุ taskId ให้ใช้ taskId ล่าสุดที่เปิดอ่าน
          let taskId = call.args.taskId;
          const userCtx = stateService.userContext.get(odataId) || {};
          if (!taskId && userCtx.lastTaskId) {
            taskId = userCtx.lastTaskId;
            console.log(`📌 Using last taskId: ${taskId} (${userCtx.lastSoNumber})`);
          }
          const task = await cemAPI.getTaskById(taskId);
          console.log(`📋 Task ${taskId}: found=${task?.id}, so_number=${task?.so_number}, files=${task?.files?.length || 0}`);
          if (!task || !task.files?.length) {
            functionResult = { error: "ไม่พบไฟล์ในโครงการนี้" };
          } else {
            // บันทึก context ล่าสุด
            stateService.userContext.set(odataId, { lastTaskId: task.id, lastSoNumber: task.so_number });
            const fileIndex = call.args.fileIndex || 0;
            const filename = task.files[fileIndex];
            if (!filename) {
              functionResult = { error: `ไม่พบไฟล์ลำดับที่ ${fileIndex}`, availableFiles: task.files };
            } else {
              const startPage = call.args.startPage || 1;
              const endPage = call.args.endPage || null;
              const fileData = await cemAPI.downloadFile(filename, startPage, endPage);
              if (!fileData) {
                functionResult = { error: "ไม่สามารถดาวน์โหลดไฟล์ได้" };
              } else {
                functionResult = { 
                  filename, 
                  taskName: task.task_name,
                  soNumber: task.so_number,
                  totalPages: fileData.pageCount,
                  pagesRead: `${fileData.startPage || 1}-${fileData.endPage || fileData.pageCount}`,
                  hasMore: fileData.hasMore || false,
                  nextPage: fileData.nextPage || null,
                  message: fileData.hasMore 
                    ? `กรุณาวิเคราะห์ไฟล์ที่แนบมา (หน้า ${fileData.startPage}-${fileData.endPage} จาก ${fileData.pageCount} หน้า) ยังมีอีก ${fileData.pageCount - fileData.endPage} หน้า ถ้าต้องการอ่านต่อให้บอก`
                    : "กรุณาวิเคราะห์ไฟล์ที่แนบมาพร้อมนี้ รวมถึงรูปภาพและตารางในเอกสาร",
                  _fileData: fileData
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

      // Helper สร้าง function response ตาม model
      const makeFunctionResponse = (name, result, forGemini3) => {
        if (forGemini3) {
          return { role: "user", parts: [{ functionResponse: { name, response: { result } } }] };
        }
        return { role: "function", parts: [{ functionResponse: { name, response: result } }] };
      };

      // สำหรับ Gemini 3: ใช้ rawContent ทั้งก้อน
      const modelPart = isGemini3 && geminiResponse.rawContent 
        ? geminiResponse.rawContent 
        : { role: "model", parts: [{ functionCall: call }] };

      // 1. Function Response Message
      const functionMsg = makeFunctionResponse(call.name, { ...functionResult, _fileData: undefined }, isGemini3);

      // 2. User Message สำหรับส่งไฟล์ (ถ้ามี)
      let fileMsg = null;
      if (functionResult._fileData) {
        const pages = functionResult._fileData.allPages || [functionResult._fileData.base64];
        // ส่งสูงสุด 20 หน้า เพื่อไม่ให้ request ใหญ่เกินไป
        const pagesToSend = pages.slice(0, 20);
        fileMsg = {
          role: "user",
          parts: pagesToSend.map(page => ({
            inlineData: {
              mimeType: functionResult._fileData.mimeType,
              data: page
            }
          }))
        };
        console.log(`📄 Sending ${pagesToSend.length}/${pages.length} pages to AI`);
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
        const { response: loopResponse, model: usedLoopModel } = await callGeminiWithFallback(
          stateService.getCurrentApiKey(), currentModel, currentHistory, userId
        );
        currentResponse = loopResponse;
        const loopIsGemini3 = usedLoopModel?.includes('gemini-3') || usedLoopModel?.includes('thinking');
        
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
          case "read_project_file":
            const loopTask = await cemAPI.getTaskById(loopCall.args.taskId);
            let loopFileData = null;
            if (!loopTask?.files?.length) {
              loopResult = { error: "ไม่พบไฟล์ในโครงการนี้" };
            } else {
              const loopFilename = loopTask.files[loopCall.args.fileIndex || 0];
              loopFileData = await cemAPI.downloadFile(loopFilename, loopCall.args.startPage || 1, loopCall.args.endPage);
              loopResult = loopFileData ? { filename: loopFilename, totalPages: loopFileData.pageCount, pagesRead: `${loopFileData.startPage}-${loopFileData.endPage}`, _fileData: loopFileData } : { error: "ไม่สามารถดาวน์โหลดไฟล์ได้" };
            }
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
        
        const loopModelPart = loopIsGemini3 && loopResponse.rawContent 
          ? loopResponse.rawContent 
          : { role: "model", parts: [{ functionCall: loopCall }] };
        
        const loopFunctionMsg = makeFunctionResponse(loopCall.name, { ...loopResult, _fileData: undefined }, loopIsGemini3);
        currentHistory = [...currentHistory, loopModelPart, loopFunctionMsg];
        
        // ส่งไฟล์ถ้ามี
        if (loopResult._fileData?.allPages) {
          const loopPages = loopResult._fileData.allPages.slice(0, 10);
          currentHistory.push({
            role: "user",
            parts: loopPages.map(page => ({ inlineData: { mimeType: loopResult._fileData.mimeType, data: page } }))
          });
          console.log(`📄 Loop: Sending ${loopPages.length} pages to AI`);
        }
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

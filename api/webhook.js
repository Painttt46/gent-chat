// api/webhook.js

import * as stateService from '../services/state.js';
import * as graphService from '../services/graphAPI.js';
import * as teamsService from '../services/teams.js';
import * as geminiService from '../services/gemini.js';

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
      stateService.currentApiKeyIndex = stateService.currentApiKeyIndex === 0 ? 1 : 0;
      currentKey = stateService.getCurrentApiKey();
      Object.keys(stateService.models).forEach(k => stateService.models[k].count = 0);
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
        default:
          functionResult = { error: "Unknown function called." };
      }

      // สำหรับ Gemini 3: ใช้ rawContent ทั้งก้อน
      const modelPart = isGemini3 && geminiResponse.rawContent 
        ? geminiResponse.rawContent 
        : { role: "model", parts: [{ functionCall: call }] };

      const historyWithFunction = [
        ...conversationHistory,
        modelPart,
        { role: "function", parts: [{ functionResponse: { name: call.name, response: functionResult } }] }
      ];

      const { response: finalResponse } = await callGeminiWithFallback(
        stateService.getCurrentApiKey(), currentModel, historyWithFunction, userId
      );
      console.log(`🔍 finalResponse type: ${typeof finalResponse}, keys: ${Object.keys(finalResponse || {})}`);
      text = finalResponse.text();
      console.log(`📝 After function call text: ${text?.substring(0, 100)}`);
    } else {
      text = geminiResponse.text();
    }

    console.log(`📝 Response text length: ${text?.length || 0}`);
    
    if (!text) {
      console.log('⚠️ Empty text response');
      return res.status(200).json({ text: '❌ ไม่ได้รับการตอบกลับจาก AI' });
    }

    const isCardFormat = text.startsWith('FORMAT:CARD');
    let cleanResponse = text.replace(/FORMAT:(CARD|TEXT)/, '').trim() || "I'm sorry, I couldn't generate a proper response.";
    console.log(`🧹 Clean response length: ${cleanResponse.length}`);

    history.push({ role: "user", parts: [{ text: finalText }] });
    history.push({ role: "model", parts: [{ text: cleanResponse }] });
    if (history.length > 40) history.splice(0, 2);

    const switchNote = switched ? ` | ⚡ Auto-switched` : '';
    const usageStats = `💬 ${Math.floor(history.length / 2)} msgs | ${stateService.models[currentModel].name} | ${stateService.models[currentModel].count}/${stateService.models[currentModel].limit}${switchNote}`;

    if (isBroadcastCommand) {
      await teamsService.sendToTeamsWebhook(`🔊 **Announcement:**\n\n${cleanResponse}\n\n${usageStats}`);
      return res.status(200).json({ text: "📢 Broadcast sent!" });
    }

    if (isCardFormat) {
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
      console.log(`📤 Sending response: ${cleanResponse.substring(0, 100)}...`);
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

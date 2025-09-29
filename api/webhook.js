import { GoogleGenerativeAI } from '@google/generative-ai';
import { ConfidentialClientApplication } from '@azure/msal-node';

// Constants
const MAX_CONVERSATION_HISTORY = 20;
const MAX_CONVERSATIONS = 1000;
const RETRY_ATTEMPTS = 3;
const RETRY_DELAY = 1000;

// Simple in-memory conversation storage (per user)
const conversations = new Map();

// Model tracking with limits
const models = {
  'gemini-2.5-flash': { name: 'Gemini 2.5 Flash', count: 0, limit: 500 },
  'gemini-2.5-pro': { name: 'Gemini 2.5 Pro', count: 0, limit: 100 }
};
const userModels = new Map();
let lastResetDate = new Date().toDateString();
let currentApiKeyIndex = 0;

// Utility functions
function getCurrentApiKey() {
  return currentApiKeyIndex === 0 ? process.env.GEMINI_API_KEY : process.env.GEMINI_API_KEY_2;
}

async function retryOperation(operation, attempts = RETRY_ATTEMPTS) {
  for (let i = 0; i < attempts; i++) {
    try {
      return await operation();
    } catch (error) {
      if (i === attempts - 1) throw error;
      await new Promise(resolve => setTimeout(resolve, RETRY_DELAY * (i + 1)));
    }
  }
}

function cleanupConversations() {
  if (conversations.size > MAX_CONVERSATIONS) {
    const entries = Array.from(conversations.entries());
    const toDelete = entries.slice(0, conversations.size - MAX_CONVERSATIONS);
    toDelete.forEach(([key]) => conversations.delete(key));
  }
}

function checkDailyReset() {
  const today = new Date().toDateString();
  if (today !== lastResetDate) {
    Object.keys(models).forEach(key => models[key].count = 0);
    lastResetDate = today;
  }
}

function checkLimitsAndSwitchKey(modelKey) {
  const model = models[modelKey];
  if (model.count >= model.limit) {
    const newApiKeyIndex = currentApiKeyIndex === 0 ? 1 : 0;
    const allModelsMaxed = Object.values(models).every(m => m.count >= m.limit);
    if (allModelsMaxed) return 'MAXED_OUT';
    
    currentApiKeyIndex = newApiKeyIndex;
    Object.keys(models).forEach(key => models[key].count = 0);
    return true;
  }
  return false;
}

// Azure/Graph API functions
async function getGraphToken() {
  return retryOperation(async () => {
    const clientConfig = {
      auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        clientSecret: process.env.AZURE_CLIENT_SECRET,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`
      }
    };
    
    const cca = new ConfidentialClientApplication(clientConfig);
    const response = await cca.acquireTokenByClientCredential({
      scopes: ['https://graph.microsoft.com/.default']
    });
    return response.accessToken;
  });
}

// ฟังก์ชันใหม่สำหรับ "ค้นหา" ผู้ใช้จากชื่อ
async function findUserByShortName(name) {
  return retryOperation(async () => {
    const token = await getGraphToken();
    const filterQuery = `$filter=startswith(displayName,'${name}') or startswith(givenName,'${name}') or startswith(mailNickname,'${name}')`;
    const selectQuery = `&$select=displayName,userPrincipalName`;
    const url = `https://graph.microsoft.com/v1.0/users?${filterQuery}${selectQuery}`;
    
    const response = await fetch(url, {
      headers: { 'Authorization': `Bearer ${token}` }
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${await response.text()}`);
    }
    
    const data = await response.json();
    return data.value;
  });
}

async function getUserCalendar(nameOrEmail) {
  try {
    let userEmail = nameOrEmail;

    if (!nameOrEmail.includes('@')) {
      const users = await findUserByShortName(nameOrEmail);
      if (!users || users.length === 0) {
        return { error: `ไม่พบผู้ใช้ที่ชื่อ '${nameOrEmail}' ในระบบครับ` };
      }
      if (users.length > 1) {
        const userList = users.map(u => u.displayName).join(', ');
        return { error: `พบผู้ใช้ที่ชื่อ '${nameOrEmail}' มากกว่า 1 คน: ${userList} กรุณาระบุให้ชัดเจนขึ้นครับ` };
      }
      userEmail = users[0].userPrincipalName;
    }

    return retryOperation(async () => {
      const token = await getGraphToken();
      const today = new Date();
      const startDateTime = new Date(today.setHours(0, 0, 0, 0)).toISOString();
      const endDateTime = new Date(today.setHours(23, 59, 59, 999)).toISOString();
      const url = `https://graph.microsoft.com/v1.0/users/${userEmail}/calendarView?startDateTime=${startDateTime}&endDateTime=${endDateTime}&$select=subject,body,bodyPreview,organizer,attendees,start,end,location`;
      
      const response = await fetch(url, {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      
      if (!response.ok) {
        throw new Error(`ไม่สามารถดึงข้อมูลปฏิทินของ ${userEmail} ได้ครับ`);
      }
      
      return await response.json();
    });
  } catch (error) {
    return { error: error.message };
  }
}
async function sendToTeamsWebhook(message) {
  const webhookUrl = 'https://gentsolutions.webhook.office.com/webhookb2/330ce018-1d89-4bde-8a00-7e112b710934@c5fc1b2a-2ce8-4471-ab9d-be65d8fe0906/IncomingWebhook/d5ec6936083f44f7aaf575f90b1f69da/0b176f81-19e0-4b39-8fc8-378244861f9b/V2FcW5LeJmT5RLRTWJR9gSZLh55QhBpny4Nll4VGmIk4I1';
  
  return retryOperation(async () => {
    const response = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text: message })
    });
    if (!response.ok) throw new Error(`Webhook failed: ${response.status}`);
    return true;
  });
}

// Message processing functions
function cleanMessageText(text) {
  return text
    .replace(/<at>.*?<\/at>/g, '')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ')
    .trim();
}

function updateConversationHistory(userId, userText, modelResponse) {
  if (!conversations.has(userId)) {
    conversations.set(userId, []);
  }
  const history = conversations.get(userId);
  
  history.push(
    { role: "user", parts: [{ text: userText }] },
    { role: "model", parts: [{ text: modelResponse }] }
  );
  
  if (history.length > MAX_CONVERSATION_HISTORY) {
    history.splice(0, 2);
  }
  
  cleanupConversations();
}

async function processGeminiRequest(userId, text) {
  const currentModel = userModels.get(userId) || 'gemini-2.5-flash';
  
  const genAI = new GoogleGenerativeAI(getCurrentApiKey());
  const calendarFunction = {
    name: "get_user_calendar",
    description: "Get calendar events for today for a specific user. You can use either their name (like 'weraprat', 'natsarin') or full email address. The system will automatically find the user in the company directory.",
    parameters: {
      type: "OBJECT",
      properties: {
        "userPrincipalName": {
          type: "STRING",
          description: "The user's name or email address. Examples: 'weraprat', 'natsarin', or 'weraprat@gent-s.com'. Just the first name is usually enough."
        }
      },
      required: ["userPrincipalName"]
    }
  };

  const systemInstruction = {
    parts: [{ text: `You are Gent, an AI work assistant helping team members in a Microsoft Teams channel. 

Your role:
- Provide professional, helpful assistance to office workers
- Be friendly, concise, and actionable in your responses
- You're part of the team conversation in this Teams channel
- Help with work-related questions, productivity tips, and general office support
- You can access calendar information for any company employee using just their first name (like 'weraprat', 'natsarin') - no need to ask for full email addresses

Response format instructions:
- For simple questions, quick answers, or casual chat: respond with "FORMAT:TEXT" followed by your response
- For any of these, use "FORMAT:CARD" followed by your response:
  * Lists, steps, or bullet points
  * Detailed explanations or tutorials
  * Multiple pieces of information
  * Structured data or comparisons
  * Professional advice or recommendations
  * When the user asks "how to" questions
  * When providing examples or templates
  * When the information would benefit from better formatting

Choose FORMAT:CARD when the response would look better with structured formatting.`}]
  };

  const model = genAI.getGenerativeModel({ 
    model: currentModel,
    tools: [{ functionDeclarations: [calendarFunction] }],
    systemInstruction: systemInstruction
  });

  const history = conversations.get(userId) || [];
  const conversationHistory = [...history, { role: "user", parts: [{ text }] }];

  let result = await model.generateContent({ contents: conversationHistory });
  const functionCalls = result.response.functionCalls();

  if (functionCalls && functionCalls.length > 0) {
    const call = functionCalls[0];
    if (call.name === "get_user_calendar") {
      const userEmail = call.args?.userPrincipalName;
      if (!userEmail) {
        return "คุณต้องการให้ฉันตรวจสอบปฏิทินของใครครับ/คะ?";
      }
      
      const calendarData = await getUserCalendar(userEmail);
      const historyWithFunction = [
        ...conversationHistory,
        { role: "model", parts: [{ functionCall: call }] },
        { role: "function", parts: [{ functionResponse: { name: "get_user_calendar", response: calendarData } }] }
      ];
      
      const finalResult = await model.generateContent({ contents: historyWithFunction });
      return finalResult.response.text();
    }
  }
  
  return result.response.text();
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    checkDailyReset();
    
    const userId = req.body?.from?.id || req.body?.channelData?.tenant?.id || 'default';
    const currentModel = userModels.get(userId) || 'gemini-2.5-flash';
    
    const switched = checkLimitsAndSwitchKey(currentModel);
    if (switched === 'MAXED_OUT') {
      return res.status(200).json({
        text: `⚠️ **Daily quota exceeded!** Both API keys have reached their limits:\n• Gemini 2.5 Flash: ${models['gemini-2.5-flash'].limit} requests\n• Gemini 2.5 Pro: ${models['gemini-2.5-pro'].limit} requests\n\nPlease try again tomorrow when counters reset.`
      });
    }
    
    models[currentModel].count++;
    
    let cleanText = cleanMessageText(req.body?.text || '');
    
    // Handle commands
    if (cleanText.toLowerCase() === 'clear') {
      conversations.delete(userId);
      return res.status(200).json({ text: "🔄 Conversation cleared! Starting fresh." });
    }

    if (cleanText.toLowerCase().startsWith('model ')) {
      const modelKey = cleanText.toLowerCase().replace('model ', '');
      if (models[modelKey]) {
        userModels.set(userId, modelKey);
        return res.status(200).json({
          text: `🤖 Switched to ${models[modelKey].name} (${models[modelKey].count}/${models[modelKey].limit} requests)`
        });
      } else {
        const modelList = Object.entries(models).map(([key, model]) => 
          `• ${key} - ${model.name} (${model.count}/${model.limit} requests)`
        ).join('\n');
        return res.status(200).json({
          text: `❌ Invalid model. Available models:\n${modelList}\n\nAPI Key: ${currentApiKeyIndex + 1}/2\nUsage: model gemini-2.5-flash`
        });
      }
    }

    if (!cleanText) return res.status(200).json({});

    const shouldBroadcast = cleanText.toLowerCase().includes('(broadcast)');
    const processedText = cleanText.replace(/\(broadcast\)/gi, '').trim() || cleanText;

    const responseText = await processGeminiRequest(userId, processedText);
    
    const isCardFormat = responseText.startsWith('FORMAT:CARD');
    const isTextFormat = responseText.startsWith('FORMAT:TEXT');
    let cleanResponse = responseText;
    
    if (isCardFormat) {
      cleanResponse = responseText.replace('FORMAT:CARD', '').trim();
    } else if (isTextFormat) {
      cleanResponse = responseText.replace('FORMAT:TEXT', '').trim();
    }

    if (!cleanResponse) {
      cleanResponse = "I'm sorry, I couldn't generate a proper response. Please try rephrasing your question.";
    }

    updateConversationHistory(userId, processedText, cleanResponse);
    const history = conversations.get(userId);
    const statsText = `💬 **${history.length / 2} messages** | **${models[currentModel].name}** | **${models[currentModel].count}/${models[currentModel].limit} requests** | **API ${currentApiKeyIndex + 1}/2**`;

    if (shouldBroadcast) {
      await sendToTeamsWebhook(`🔊 **Announcement from Gent:**\n\n${cleanResponse}\n\n${statsText}`);
      return res.status(200).json({ text: "📢 Broadcast sent successfully!" });
    }

    if (isCardFormat) {
      return res.status(200).json({
        type: "message",
        attachments: [{
          contentType: "application/vnd.microsoft.card.adaptive",
          content: {
            type: "AdaptiveCard",
            version: "1.2",
            body: [
              {
                type: "TextBlock",
                text: "🤖 Gent - Work Assistant",
                weight: "Bolder",
                size: "Medium",
                color: "Accent"
              },
              {
                type: "TextBlock",
                text: cleanResponse,
                wrap: true,
                spacing: "Medium"
              },
              {
                type: "TextBlock",
                text: statsText,
                size: "Small",
                color: "Good",
                weight: "Bolder",
                spacing: "Medium"
              }
            ]
          }
        }]
      });
    } else {
      return res.status(200).json({
        text: `🤖 **Gent:** ${cleanResponse}\n\n${statsText}`
      });
    }

  } catch (error) {
    const userId = req.body?.from?.id || 'default';
    const currentModel = userModels.get(userId) || 'gemini-2.5-flash';
    const shouldBroadcast = (req.body?.text || '').toLowerCase().includes('(broadcast)');

    if (shouldBroadcast) {
      try {
        await sendToTeamsWebhook(`🔊 **Gent Error:**\n\nSorry, I'm having trouble right now. Please try again.\n\n💬 **${conversations.get(userId)?.length / 2 || 0} messages** | **${models[currentModel].name}** | **${models[currentModel].count}/${models[currentModel].limit} requests** | **API ${currentApiKeyIndex + 1}/2**`);
        return res.status(200).json({ text: "❌ Broadcast failed - error sent to channel" });
      } catch (webhookError) {
        return res.status(200).json({ text: "❌ Broadcast and error notification failed" });
      }
    }
    
    return res.status(200).json({
      text: `❌ **Gent:** Sorry, I'm having trouble right now. Please try again.\n\nError: ${error.message}`
    });
  }
}
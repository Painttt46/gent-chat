import { GoogleGenerativeAI } from '@google/generative-ai';
import { ConfidentialClientApplication } from '@azure/msal-node';
// Simple in-memory conversation storage (per user)
const conversations = new Map();

// Model tracking with limits
const models = {
  'gemini-2.5-flash': { name: 'Gemini 2.5 Flash', count: 0, limit: 500 },
  'gemini-2.5-pro': { name: 'Gemini 2.5 Pro', count: 0, limit: 100 }
};
const userModels = new Map(); // Track current model per user
let lastResetDate = new Date().toDateString();
let currentApiKeyIndex = 0; // 0 for primary, 1 for secondary

// Get current API key
function getCurrentApiKey() {
  return currentApiKeyIndex === 0 ? process.env.GEMINI_API_KEY : process.env.GEMINI_API_KEY_2;
}

// Get Graph API access token
async function getGraphToken() {
  try {
    const clientConfig = {
      auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        clientSecret: process.env.AZURE_CLIENT_SECRET,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`
      }
    };

    console.log('Azure config:', {
      clientId: process.env.AZURE_CLIENT_ID ? 'Set' : 'Missing',
      clientSecret: process.env.AZURE_CLIENT_SECRET ? 'Set' : 'Missing',
      tenantId: process.env.AZURE_TENANT_ID ? 'Set' : 'Missing'
    });

    const cca = new ConfidentialClientApplication(clientConfig);
    const clientCredentialRequest = {
      scopes: ['https://graph.microsoft.com/.default']
    };

    const response = await cca.acquireTokenByClientCredential(clientCredentialRequest);
    console.log('Token acquired successfully');
    return response.accessToken;
  } catch (error) {
    console.error('Token acquisition error:', error);
    throw error;
  }
}

// ฟังก์ชันใหม่สำหรับ "ค้นหา" ผู้ใช้จากชื่อ
async function findUserByShortName(name) {
  try {
    const token = await getGraphToken();
    // ✅ เปลี่ยน: Query ให้ค้นหาแบบตรงตัว (eq) และเพิ่ม field ในการค้นหา
    const filterQuery = `$filter=displayName eq '${name}' or givenName eq '${name}' or surname eq '${name}' or mailNickname eq '${name}'`;
    const selectQuery = `&$select=displayName,userPrincipalName`;
    const url = `https://graph.microsoft.com/v1.0/users?${filterQuery}${selectQuery}`;

    console.log('Searching for user with URL:', url);

    const response = await fetch(url, {
      headers: { 'Authorization': `Bearer ${token}` }
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error('Graph API user search error:', errorText);
      return { error: `HTTP ${response.status}: ${errorText}` };
    }

    const data = await response.json();
    return data.value;
  } catch (error) {
    console.error('findUserByShortName error:', error);
    return { error: error.message };
  }
}

// ปรับแก้ฟังก์ชัน getUserCalendar เดิม
async function getUserCalendar(nameOrEmail, startDate = null, endDate = null) {
  let userEmail = nameOrEmail;

  // ถ้าเป็นอีเมลแล้ว ให้ใช้เลย ถ้าไม่ใช่ ให้ค้นหาจากชื่อ
  if (!nameOrEmail.includes('@')) {
    console.log(`Searching for user: '${nameOrEmail}'`);
    const users = await findUserByShortName(nameOrEmail);

    if (!users || users.length === 0) {
      return { error: `ไม่พบผู้ใช้ที่ชื่อ '${nameOrEmail}' ในระบบครับ` };
    }
    if (users.length > 1) {
      const userList = users.map(u => u.displayName).join(', ');
      return { error: `พบผู้ใช้ที่ชื่อ '${nameOrEmail}' มากกว่า 1 คน: ${userList} กรุณาระบุให้ชัดเจนขึ้นครับ` };
    }

    userEmail = users[0].userPrincipalName;
    console.log(`User found: ${users[0].displayName} (${userEmail})`);
  }

  try {
    const token = await getGraphToken();
    
    // Set date range - default to today if not specified
    let startDateTime, endDateTime;
    if (startDate && endDate) {
      // Parse dates and set to Bangkok timezone
      const start = new Date(startDate + 'T00:00:00+07:00');
      const end = new Date(endDate + 'T23:59:59+07:00');
      startDateTime = start.toISOString();
      endDateTime = end.toISOString();
    } else if (startDate) {
      // Single date - full day in Bangkok timezone
      const start = new Date(startDate + 'T00:00:00+07:00');
      const end = new Date(startDate + 'T23:59:59+07:00');
      startDateTime = start.toISOString();
      endDateTime = end.toISOString();
    } else {
      // Default to today in Bangkok timezone
      const now = new Date();
      const bangkokOffset = 7 * 60; // Bangkok is UTC+7
      const bangkokTime = new Date(now.getTime() + (bangkokOffset * 60 * 1000));
      const todayStr = bangkokTime.toISOString().split('T')[0];
      
      const start = new Date(todayStr + 'T00:00:00+07:00');
      const end = new Date(todayStr + 'T23:59:59+07:00');
      startDateTime = start.toISOString();
      endDateTime = end.toISOString();
    }
    
    const url = `https://graph.microsoft.com/v1.0/users/${userEmail}/calendarView?startDateTime=${startDateTime}&endDateTime=${endDateTime}&$select=subject,body,bodyPreview,organizer,attendees,start,end,location`;

    console.log('Fetching calendar for:', userEmail, 'from', startDateTime, 'to', endDateTime);

    const response = await fetch(url, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Prefer': 'outlook.timezone="Asia/Bangkok"'
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error('Graph API error response:', errorText);
      return { error: `ไม่สามารถดึงข้อมูลปฏิทินของ ${userEmail} ได้ครับ` };
    }

    const data = await response.json();
    return data;
  } catch (error) {
    console.error('Graph API error:', error);
    return { error: error.message };
  }
}
async function sendToTeamsWebhook(message) {
  const webhookUrl = 'https://gentsolutions.webhook.office.com/webhookb2/330ce018-1d89-4bde-8a00-7e112b710934@c5fc1b2a-2ce8-4471-ab9d-be65d8fe0906/IncomingWebhook/d5ec6936083f44f7aaf575f90b1f69da/0b176f81-19e0-4b39-8fc8-378244861f9b/V2FcW5LeJmT5RLRTWJR9gSZLh55QhBpny4Nll4VGmIk4I1';

  try {
    await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text: message })
    });
    return true;
  } catch (error) {
    console.error('Teams webhook error:', error);
    return false;
  }
}

// Check if model hit limit and switch API key if needed
function checkLimitsAndSwitchKey(modelKey) {
  const model = models[modelKey];
  if (model.count >= model.limit) {
    // Try to switch to other API key
    const newApiKeyIndex = currentApiKeyIndex === 0 ? 1 : 0;

    // Check if we've already tried both keys (both are maxed)
    const allModelsMaxed = Object.values(models).every(m => m.count >= m.limit);
    if (allModelsMaxed) {
      return 'MAXED_OUT';
    }

    // Switch to other API key and reset counters
    currentApiKeyIndex = newApiKeyIndex;
    Object.keys(models).forEach(key => models[key].count = 0);
    return true; // Switched
  }
  return false; // No switch
}

// Reset counters daily
function checkDailyReset() {
  const today = new Date().toDateString();
  if (today !== lastResetDate) {
    Object.keys(models).forEach(key => models[key].count = 0);
    lastResetDate = today;
  }
}

// Get total requests across all models
function getTotalRequests() {
  return Object.values(models).reduce((sum, model) => sum + model.count, 0);
}

export default async function handler(req, res) {
  // Only handle POST requests
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  console.log('Method:', req.method);
  console.log('Body:', req.body);

  // Check and reset counters daily
  checkDailyReset();

  // Get user ID from Teams (fallback to 'default' if not available)
  const userId = req.body?.from?.id || req.body?.channelData?.tenant?.id || 'default';

  // Get current model for user and increment counter for each webhook request
  const currentModel = userModels.get(userId) || 'gemini-2.5-flash';

  // Check limits and switch API key if needed
  const switched = checkLimitsAndSwitchKey(currentModel);

  // If both API keys are maxed out, return error
  if (switched === 'MAXED_OUT') {
    return res.status(200).json({
      text: `⚠️ **Daily quota exceeded!** Both API keys have reached their limits:\n• Gemini 2.5 Flash: ${models['gemini-2.5-flash'].limit} requests\n• Gemini 2.5 Pro: ${models['gemini-2.5-pro'].limit} requests\n\nPlease try again tomorrow when counters reset.`
    });
  }

  models[currentModel].count++;

  // Clean mention from text
  let cleanText = req.body?.text || '';
  cleanText = cleanText.replace(/<at>.*?<\/at>/g, '');
  cleanText = cleanText.replace(/<[^>]*>/g, '');
  cleanText = cleanText.replace(/&nbsp;/g, ' ');
  cleanText = cleanText.replace(/&amp;/g, '&');
  cleanText = cleanText.replace(/&lt;/g, '<');
  cleanText = cleanText.replace(/&gt;/g, '>');
  cleanText = cleanText.replace(/\s+/g, ' ').trim();

  // Handle clear command
  if (cleanText.toLowerCase() === 'clear') {
    conversations.delete(userId);
    return res.status(200).json({
      text: "🔄 Conversation cleared! Starting fresh."
    });
  }

  // Handle model switching
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

  if (!cleanText) {
    // Don't respond to empty messages to avoid Teams errors
    return res.status(200).json({});
  }

  try {
    // Check if user wants to broadcast to everyone
    const isBroadcastCommand = cleanText.toLowerCase().startsWith('/broadcast ');

    // Remove (broadcast) from the message before processing
    const finalText = isBroadcastCommand
      ? cleanText.substring(10).trim()
      : cleanText;

    // Initialize Gemini AI with current API key and function calling
    const genAI = new GoogleGenerativeAI(getCurrentApiKey());

    // Define function for calendar access
    const calendarFunction = {
      name: "get_user_calendar",
      description: "Get calendar events for a specific user within a date range. You can use either their name (like 'weraprat', 'natsarin') or full email address. If no dates are specified, it defaults to today only.",
      parameters: {
        type: "OBJECT",
        properties: {
          "userPrincipalName": {
            type: "STRING",
            description: "The user's name or email address. Examples: 'weraprat', 'natsarin', or 'weraprat@gent-s.com'. Just the first name is usually enough."
          },
          "startDate": {
            type: "STRING",
            description: "Start date in YYYY-MM-DD format (optional). If not provided, defaults to today."
          },
          "endDate": {
            type: "STRING", 
            description: "End date in YYYY-MM-DD format (optional). If not provided but startDate is given, defaults to same day as startDate."
          }
        },
        required: ["userPrincipalName"]
      }
    };

    const systemInstruction = {
      parts: [{
        text: `You are Gent, an AI work assistant helping team members in a Microsoft Teams channel. 

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

    // Get or create conversation history for this user
    if (!conversations.has(userId)) {
      conversations.set(userId, []);
    }
    const history = conversations.get(userId);

    // Build conversation history with current message
    const conversationHistory = [
      ...history,
      { role: "user", parts: [{ text: finalText }] }
    ];

    // Generate content with history
    let result = await model.generateContent({
      contents: conversationHistory
    });

    let text;
    const functionCalls = result.response.functionCalls();

    if (functionCalls && functionCalls.length > 0) {
      console.log("Gemini wants to call a function...");
      const call = functionCalls[0];

      if (call.name === "get_user_calendar") {
        const userEmail = call.args?.userPrincipalName || req.body?.from?.userPrincipalName || req.body?.from?.email;
        const startDate = call.args?.startDate;
        const endDate = call.args?.endDate;

        if (!userEmail) {
          text = "คุณต้องการให้ฉันตรวจสอบปฏิทินของใครครับ/คะ?";
        } else {
          const calendarData = await getUserCalendar(userEmail, startDate, endDate);

          // Build history with function call and response
          const historyWithFunction = [
            ...conversationHistory,
            { role: "model", parts: [{ functionCall: call }] },
            { role: "function", parts: [{ functionResponse: { name: "get_user_calendar", response: calendarData } }] }
          ];

          // Generate final response
          const finalResult = await model.generateContent({
            contents: historyWithFunction
          });
          text = finalResult.response.text();
        }
      } else {
        text = "Unknown function called.";
      }
    } else {
      text = result.response.text();
    }

    // Parse format choice
    const isCardFormat = text.startsWith('FORMAT:CARD');
    const isTextFormat = text.startsWith('FORMAT:TEXT');

    let cleanResponse = text;
    if (isCardFormat) {
      cleanResponse = text.replace('FORMAT:CARD', '').trim();
    } else if (isTextFormat) {
      cleanResponse = text.replace('FORMAT:TEXT', '').trim();
    }

    // Fallback for empty responses
    if (!cleanResponse || cleanResponse.length === 0) {
      cleanResponse = "I'm sorry, I couldn't generate a proper response. Please try rephrasing your question.";
    }

    // Save to conversation history (keep last 10 messages)
    history.push({ role: "user", parts: [{ text: finalText }] });
    history.push({ role: "model", parts: [{ text: cleanResponse }] });
    if (history.length > 20) { // Keep last 10 exchanges (20 messages)
      history.splice(0, 2);
    }

    if (isBroadcastCommand) {
      // Send to Teams incoming webhook with stats
      const broadcastMessage = `🔊 **Announcement from Gent:**\n\n${cleanResponse}\n\n💬 **${history.length / 2} messages** | **${models[currentModel].name}** | **${models[currentModel].count}/${models[currentModel].limit} requests** | **API ${currentApiKeyIndex + 1}/2**`;
      await sendToTeamsWebhook(broadcastMessage);

      return res.status(200).json({
        text: "📢 Broadcast sent successfully!"
      });
    }

    // Return based on Gemini's format choice
    if (isCardFormat) {
      // Return as adaptive card
      res.status(200).json({
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
                text: `💬 **${history.length / 2} messages** | **${models[currentModel].name}** | **${models[currentModel].count}/${models[currentModel].limit} requests** | **API ${currentApiKeyIndex + 1}/2**`,
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
      // Return as simple text
      res.status(200).json({
        text: `🤖 **Gent:** ${cleanResponse}\n\n💬 **${history.length / 2} messages** | **${models[currentModel].name}** | **${models[currentModel].count}/${models[currentModel].limit} requests** | **API ${currentApiKeyIndex + 1}/2**`
      });
    }

  } catch (error) {
    console.error('Gemini API error:', error);

    // Get current model for error message
    const currentModel = userModels.get(userId) || 'gemini-2.5-flash';
    const isBroadcastCommand = cleanText.toLowerCase().startsWith('/broadcast ');

    // If broadcast, send error to Teams webhook
    if (isBroadcastCommand) {
      const errorMessage = `🔊 **Gent Error:**\n\nSorry, I'm having trouble right now. Please try again.\n\n💬 **${conversations.get(userId)?.length / 2 || 0} messages** | **${models[currentModel].name}** | **${models[currentModel].count}/${models[currentModel].limit} requests** | **API ${currentApiKeyIndex + 1}/2**`;
      await sendToTeamsWebhook(errorMessage);
      return res.status(200).json({
        text: "❌ Broadcast failed - error sent to channel"
      });
    }

    res.status(200).json({
      text: `❌ **Gent:** Sorry, I'm having trouble right now. Please try again.\n\nError: ${error.message}`
    });
  }
}
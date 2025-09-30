import { GoogleGenerativeAI } from '@google/generative-ai';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { fromZonedTime, toZonedTime } from 'date-fns-tz';
import { parseISO, startOfDay, endOfDay } from 'date-fns';
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
    const bangkokTz = 'Asia/Bangkok';

    console.log('Date parameters received:', { startDate, endDate });

    if (startDate && endDate) {
      // Parse dates and convert to Bangkok timezone, then to UTC
      const start = fromZonedTime(startOfDay(parseISO(startDate)), bangkokTz);
      const end = fromZonedTime(endOfDay(parseISO(endDate)), bangkokTz);
      startDateTime = start.toISOString();
      endDateTime = end.toISOString();
    } else if (startDate) {
      // Single date - full day in Bangkok timezone
      const start = fromZonedTime(startOfDay(parseISO(startDate)), bangkokTz);
      const end = fromZonedTime(endOfDay(parseISO(startDate)), bangkokTz);
      startDateTime = start.toISOString();
      endDateTime = end.toISOString();
    } else {
      // Default to today in Bangkok timezone
      const now = new Date();
      const bangkokNow = toZonedTime(now, bangkokTz);
      const start = fromZonedTime(startOfDay(bangkokNow), bangkokTz);
      const end = fromZonedTime(endOfDay(bangkokNow), bangkokTz);
      startDateTime = start.toISOString();
      endDateTime = end.toISOString();
    }

    console.log('Final date range:', { startDateTime, endDateTime });

    const url = `https://graph.microsoft.com/v1.0/users/${userEmail}/calendarView?startDateTime=${startDateTime}&endDateTime=${endDateTime}&$select=subject,body,bodyPreview,organizer,attendees,start,end,location,onlineMeeting`;

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
// Create calendar event
// ✅ แทนที่ฟังก์ชัน createCalendarEvent เดิมด้วยอันนี้
async function createCalendarEvent({ subject, startDateTime, endDateTime, attendees, bodyContent, location, createMeeting = true }) {
  try {
    const token = await getGraphToken();

    // --- ส่วนจัดการ Attendee และ Organizer (เหมือนเดิม) ---
    let attendeeObjects = [];
    if (attendees && attendees.length > 0) {
      const userLookups = await Promise.all(
        attendees.map(name => findUserByShortName(name.trim()))
      );

      userLookups.forEach((users, index) => {
        if (users && users.length === 1) {
          attendeeObjects.push({
            emailAddress: {
              address: users[0].userPrincipalName,
              name: users[0].displayName
            },
            type: "required"
          });
        } else {
          console.warn(`Could not resolve attendee: ${attendees[index]}`);
        }
      });
    }

    const organizer = await findUserByShortName(attendees[0]?.trim());
    if (!organizer || organizer.length !== 1) {
      return { error: "ไม่สามารถระบุผู้จัดงาน (organizer) ได้ กรุณาระบุผู้เข้าร่วมอย่างน้อย 1 คนที่เป็นผู้ใช้ในระบบ" };
    }
    const organizerEmail = organizer[0].userPrincipalName;

    // --- สร้าง Request Body สำหรับ Graph API ---
    const event = {
      subject: subject,
      body: {
        contentType: "HTML",
        content: bodyContent || "This meeting was scheduled by Gent AI Assistant."
      },
      start: {
        dateTime: startDateTime,
        timeZone: "Asia/Bangkok"
      },
      end: {
        dateTime: endDateTime,
        timeZone: "Asia/Bangkok"
      },
      location: {
        displayName: location || ""
      },
      attendees: attendeeObjects,
    };

    // ✅ เพิ่มเงื่อนไขในการสร้าง Meeting Link ตรงนี้
    if (createMeeting) {
      event.isOnlineMeeting = true;
      event.onlineMeetingProvider = "teamsForBusiness";
    }

    const url = `https://graph.microsoft.com/v1.0/users/${organizerEmail}/events`;

    console.log('Creating event with payload:', JSON.stringify(event, null, 2));

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(event)
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error('Graph API create event error:', errorText);
      return { error: `HTTP ${response.status}: ${errorText}` };
    }

    const createdEvent = await response.json();
    console.log('Event created successfully:', createdEvent.id);

    // ✅ อัปเดตข้อมูลที่ส่งกลับให้ AI
    return {
      success: true,
      subject: createdEvent.subject,
      startTime: createdEvent.start.dateTime,
      organizer: createdEvent.organizer.emailAddress.name,
      webLink: createdEvent.webLink,
      meetingCreated: createdEvent.isOnlineMeeting || false // ส่งสถานะกลับไปด้วย
    };

  } catch (error) {
    console.error('createCalendarEvent error:', error);
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
    // Check if user wants to broadcast 
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
      description: `Get calendar events for a user within a specified date range. 
      The model is responsible for interpreting natural language date expressions and converting them into a precise YYYY-MM-DD format for startDate and endDate.
      - If no dates are provided, it defaults to today.
      - display all parameter form api if can.
      - Understands relative terms based on the current date. For example:
        - "yesterday", "tomorrow"
        - "this week", "next week", "last week" (Assume week starts on Monday)
        - "this month", "next month", "last month"
      - The model MUST calculate the exact start and end dates before calling the tool. For example, if the user asks for "next week", the model should calculate the dates for the upcoming Monday and Sunday and pass them as startDate and endDate.`,

      // ✅ แก้ไข: เพิ่ม "parameters" object ครอบส่วนนี้
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

    // Define function for creating calendar events
    // ภายในฟังก์ชัน handler(req, res)

    // ✅ แทนที่ createEventFunction เดิมด้วยอันนี้
    const createEventFunction = {
      name: "create_calendar_event",
      description: `Creates a new calendar event and sends Microsoft Teams meeting invitations to attendees.
  - The model must infer the subject, start time, and end time from the user's request.
  - The current date is ${new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Bangkok' })}.
  - It must convert all date/time information into 'YYYY-MM-DDTHH:mm:ss' format. For example, 'tomorrow at 3 PM' should be converted to the correct full timestamp.
  - The 'attendees' parameter should be an array of short names found in the prompt (e.g., ['weraprat', 'natsarin']). The person organizing the meeting must be the first person in this list.`,
      parameters: {
        type: "OBJECT",
        properties: {
          "subject": {
            type: "STRING",
            description: "The title or subject of the event."
          },
          "startDateTime": {
            type: "STRING",
            description: "The start date and time in 'YYYY-MM-DDTHH:mm:ss' format."
          },
          "endDateTime": {
            type: "STRING",
            description: "The end date and time in 'YYYY-MM-DDTHH:mm:ss' format."
          },
          "attendees": {
            type: "ARRAY",
            description: "A list of attendees' names. The first name in the list will be the meeting organizer. Example: ['weraprat', 'natsarin']",
            items: {
              type: "STRING"
            }
          },
          // ✅ จุดที่แก้ไขคือเพิ่ม createMeeting เข้ามาตรงนี้
          "createMeeting": {
            type: "BOOLEAN",
            description: `Set to true to create a Teams meeting link. Set to false for a simple calendar booking. Infer from keywords like 'meeting', 'call' (true) vs 'book', 'reserve', 'block time' (false). Defaults to true.`
          },
          "bodyContent": {
            type: "STRING",
            description: "Optional. The main description or body of the event in HTML format."
          },
          "location": {
            type: "STRING",
            description: "Optional. The physical location of the meeting."
          }
        },
        required: ["subject", "startDateTime", "endDateTime", "attendees"]
      }
    };
    // ✅✅✅  นำ systemInstruction นี้ไปวางทับของเดิมในไฟล์ของคุณ ✅✅✅
    const systemInstruction = {
      parts: [{
        text: `You are Gent (เจนต์), a proactive and friendly AI work assistant integrated into a Microsoft Teams channel. Your primary goal is to help team members be more productive and collaborative.

---

### **Core Persona & Tone (บุคลิกและน้ำเสียง):**
- **Name:** Gent (เจนต์)
- **Personality:** Professional, friendly, slightly informal, and very helpful. You are a member of the team.
- **Language:** Respond primarily in Thai (ตอบเป็นภาษาไทยเป็นหลัก). Be concise and clear.
- **Proactive:** Don't just answer; anticipate needs. If a meeting is scheduled, ask if an agenda is needed. If a user seems busy, suggest finding an alternative time.

---

### **Key Capabilities & Rules (ความสามารถและกฎการทำงาน):**
You have access to two main tools: \`get_user_calendar\` and \`Calendar\`.

1.  **Viewing Calendars (\`get_user_calendar\`):**
    * **CRITICAL RULE:** For ANY request about schedules, availability, or events (e.g., "ใครว่างบ้างพรุ่งนี้", "ดูตารางงานของวีรปรัชญ์", "พรุ่งนี้ฉันมีประชุมอะไรไหม"), you **MUST** call the \`get_user_calendar\` function.
    * **NEVER** answer from memory. Always fetch fresh data.
    * You can find users by their first name (e.g., 'weraprat', 'natsarin'). You don't need a full email.

2.  **Creating Events (\`Calendar\`):**
    * **CRITICAL RULE:** For ANY request to book, schedule, create, or set up an event, meeting, or calendar block (e.g., "นัดประชุมให้หน่อย", "จองเวลาพรุ่งนี้"), you **MUST** call the \`Calendar\` function.
    * **Handling "Myself":** If the user says the meeting is for 'myself', 'me' (ตัวเอง, ฉัน), or doesn't specify any attendees, **DO NOT ask for their name**. Instead, call the tool with an **empty \`attendees\` array** (\`[]\`). The system is designed to automatically use the current user's name in this case.
    * **Confirmation is Key:** Before calling the function, **summarize the details** (Subject, Time, Attendees, Meeting Link status) and **ask the user for confirmation**. For example: "โอเคครับ, ผมจะสร้างนัดหมาย 'คุยโปรเจค' พรุ่งนี้ 10:00-11:00 น. มีคุณวีรปรัชญ์เข้าร่วม พร้อมลิงก์ประชุม Teams นะครับ ยืนยันไหมครับ?"
    * **Handle Ambiguity:** If details are missing (like end time or attendees), **ask clarifying questions**. Don't assume. Example: "ได้เลยครับ ประชุมเริ่ม 10 โมง ใช้เวลาประมาณเท่าไหร่ดีครับ?"
    * **Meeting Link Inference:** Use the \`createMeeting\` parameter based on the user's language. Keywords like "ประชุม", "คอล", "meeting", "หารือ" imply \`createMeeting: true\`. Keywords like "จองเวลา", "บล็อกคิว", "ทำงานส่วนตัว" imply \`createMeeting: false\`. If unsure, default to \`true\` and mention it in the confirmation.

---

### **Important Context (ข้อมูลแวดล้อม):**
- **Current Date:** ${new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Bangkok' })} (YYYY-MM-DD format). Use this to resolve relative dates like "tomorrow" or "next Friday".

---

### **Response Formatting (รูปแบบการตอบ):**
- Use Markdown for clear formatting (bolding, bullet points, etc.).
- **FORMAT:CARD:** Use this for structured responses like lists, summaries, or when presenting calendar data. This helps the system render a nice visual card in Teams.
- **FORMAT:TEXT:** Use this for simple, conversational replies, confirmations, or questions.
- **Always start your final response with either \`FORMAT:CARD\` or \`FORMAT:TEXT\`.**

---

### **Example Flow (ตัวอย่างการทำงาน):**
* **User:** "นัดประชุมวีรปรัชญ์พรุ่งนี้ 10 โมงหน่อยสิ"
* **Your Thought Process:** Missing end time and confirmation. I need to ask a clarifying question.
* **Your Response (FORMAT:TEXT):** "ได้เลยครับ นัดประชุมคุณวีรปรัชญ์พรุ่งนี้ 10 โมง ใช้เวลาประมาณเท่าไหร่ดีครับ 1 ชั่วโมงไหม?"
* **User:** "ใช่ 1 ชั่วโมง"
* **Your Thought Process:** Now I have all details. I must confirm before creating the event.
* **Your Response (FORMAT:TEXT):** "รับทราบครับ ผมกำลังจะสร้างนัดหมาย 'ประชุม' กับคุณวีรปรัชญ์พรุ่งนี้ 10:00 - 11:00 น. พร้อมลิงก์ประชุม Teams นะครับ"
* **System:** (Calls \`Calendar\` tool after this confirmation)
* **Your Final Response (FORMAT:CARD):** "เรียบร้อยครับ! ผมได้สร้างนัดหมายและส่งคำเชิญให้คุณวีรปรัชญ์แล้วครับ ✅\\n\\n- **หัวข้อ:** ประชุม\\n- **เวลา:** 10:00 - 11:00 น.\\n- **ผู้เข้าร่วม:** วีรปรัชญ์"
`
      }]
    };


    const model = genAI.getGenerativeModel({
      model: currentModel,
      tools: [{ functionDeclarations: [calendarFunction, createEventFunction] }],
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
      } else if (call.name === "create_calendar_event") {
        // รับค่าทั้งหมดจาก call.args ที่ Gemini ส่งมา (ซึ่งตอนนี้จะมี attendees ด้วย)
        const eventData = call.args;
        if (!eventData.attendees || eventData.attendees.length === 0) {
          // ให้ใช้ชื่อของผู้ใช้ที่ส่งข้อความมาเป็น attendee คนแรก (และเป็น organizer)
          const userName = req.body?.from?.name;
          if (userName) {
            // แยกชื่อจริงออกจากนามสกุล (ถ้ามี) แล้วใช้แค่ชื่อแรก
            const firstName = userName.split(' ')[0];
            eventData.attendees = [firstName];
            console.log(`No attendees specified, defaulting to current user: ${firstName}`);
          }
        }

        // ตรวจสอบข้อมูลสำคัญ รวมถึงเช็คว่ามี attendees อย่างน้อย 1 คนหรือไม่
        if (!eventData.subject || !eventData.startDateTime || !eventData.endDateTime || !eventData.attendees || eventData.attendees.length === 0) {
          text = "ขออภัยครับ ข้อมูลสำหรับสร้างนัดหมายไม่ครบถ้วน กรุณาระบุหัวข้อ, เวลาเริ่มต้น-สิ้นสุด, และผู้เข้าร่วมอย่างน้อย 1 คนครับ";
        } else {
          // เรียกใช้ฟังก์ชัน createCalendarEvent เวอร์ชันใหม่ที่รับ parameter แค่ตัวเดียว
          const createResult = await createCalendarEvent(eventData);

          // ส่วนที่เหลือสำหรับส่งข้อมูลกลับไปให้ Gemini สรุปผล
          const historyWithFunction = [
            ...conversationHistory,
            { role: "model", parts: [{ functionCall: call }] },
            { role: "function", parts: [{ functionResponse: { name: "create_calendar_event", response: createResult } }] }
          ];

          const finalResult = await model.generateContent({
            contents: historyWithFunction
          });
          text = finalResult.response?.text() ?? "";
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
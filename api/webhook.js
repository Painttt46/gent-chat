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

let cachedGraphToken = {
  token: null,
  expiresOn: null
};

// Get Graph API access token
async function getGraphToken() {
  // ถ้ามี token ใน cache และยังไม่หมดอายุ ก็ให้ใช้ token นั้นเลย
  if (cachedGraphToken.token && new Date() < cachedGraphToken.expiresOn) {
    console.log('Using cached Graph token');
    return cachedGraphToken.token;
  }

  try {
    const clientConfig = {
      auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        clientSecret: process.env.AZURE_CLIENT_SECRET,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`
      }
    };
    const cca = new ConfidentialClientApplication(clientConfig);
    const clientCredentialRequest = {
      scopes: ['https://graph.microsoft.com/.default']
    };

    const response = await cca.acquireTokenByClientCredential(clientCredentialRequest);

    // เก็บ token และเวลาหมดอายุลงใน cache
    // ลดเวลาลงเล็กน้อย (เช่น 5 นาที) เพื่อป้องกันปัญหาเรื่องเวลาเหลื่อมกัน
    if (response && response.accessToken && response.expiresOn) {
      cachedGraphToken.token = response.accessToken;
      cachedGraphToken.expiresOn = new Date(response.expiresOn.getTime() - 5 * 60 * 1000);
      console.log('New Graph token acquired and cached.');
      return response.accessToken;
    } else {
      throw new Error('Failed to acquire token or token response is invalid.');
    }

  } catch (error) {
    console.error('Token acquisition error:', error);
    // เคลียร์ cache ถ้าหากเกิด error
    cachedGraphToken.token = null;
    cachedGraphToken.expiresOn = null;
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
/**
 * Finds available time slots for a group of attendees within a given date range.
 */
async function findAvailableTime({ attendees, durationInMinutes, startSearch, endSearch }) {
  try {
    console.log('Finding available time for:', { attendees, durationInMinutes, startSearch, endSearch });
    const bangkokTz = 'Asia/Bangkok';

    // 1. Resolve all attendee names to their email addresses (UPNs)
    const userLookups = await Promise.all(
      attendees.map(name => findUserByShortName(name.trim()))
    );

    const resolvedUsers = [];
    const unresolvedNames = [];
    userLookups.forEach((users, index) => {
      if (users && users.length === 1) {
        resolvedUsers.push(users[0]);
      } else {
        unresolvedNames.push(attendees[index]);
      }
    });

    if (unresolvedNames.length > 0) {
      return { error: `ไม่พบผู้ใช้: ${unresolvedNames.join(', ')}` };
    }
    if (resolvedUsers.length === 0) {
      return { error: 'ไม่พบรายชื่อผู้เข้าร่วมที่ถูกต้อง' };
    }

    // 2. Fetch calendars for all attendees in the specified date range
    const calendarPromises = resolvedUsers.map(user =>
      getUserCalendar(user.userPrincipalName, startSearch, endSearch)
    );
    const calendarResults = await Promise.all(calendarPromises);

    // 3. Merge all busy slots into a single array
    let allBusySlots = [];
    for (const result of calendarResults) {
      if (result.value) {
        result.value.forEach(event => {
          allBusySlots.push({
            start: new Date(event.start.dateTime + 'Z'), // Append Z to treat as UTC
            end: new Date(event.end.dateTime + 'Z')
          });
        });
      }
    }

    // Sort busy slots by start time
    allBusySlots.sort((a, b) => a.start - b.start);

    // 4. Find gaps between busy slots, considering working hours (9 AM - 6 PM)
    const availableSlots = [];
    const workingHoursStart = 9;
    const workingHoursEnd = 18; // 6 PM
    let searchDate = fromZonedTime(startOfDay(parseISO(startSearch)), bangkokTz);
    const searchEndDate = fromZonedTime(endOfDay(parseISO(endSearch)), bangkokTz);

    while (searchDate <= searchEndDate && availableSlots.length < 5) {
      let potentialSlotStart = toZonedTime(searchDate, bangkokTz);
      potentialSlotStart.setHours(workingHoursStart, 0, 0, 0);

      const dayEnd = toZonedTime(searchDate, bangkokTz);
      dayEnd.setHours(workingHoursEnd, 0, 0, 0);

      const todayBusySlots = allBusySlots.filter(slot =>
        startOfDay(toZonedTime(slot.start, bangkokTz)).getTime() === startOfDay(searchDate).getTime()
      );

      for (const busySlot of todayBusySlots) {
        const gapMillis = busySlot.start - potentialSlotStart;
        const gapMinutes = Math.floor(gapMillis / (1000 * 60));

        if (gapMinutes >= durationInMinutes) {
          availableSlots.push({
            start: potentialSlotStart.toISOString(),
            end: new Date(potentialSlotStart.getTime() + durationInMinutes * 60000).toISOString()
          });
          if (availableSlots.length >= 5) break;
        }
        potentialSlotStart = new Date(Math.max(potentialSlotStart, busySlot.end));
      }

      if (availableSlots.length < 5 && potentialSlotStart < dayEnd) {
        const finalGapMillis = dayEnd - potentialSlotStart;
        const finalGapMinutes = Math.floor(finalGapMillis / (1000 * 60));
        if (finalGapMinutes >= durationInMinutes) {
          availableSlots.push({
            start: potentialSlotStart.toISOString(),
            end: new Date(potentialSlotStart.getTime() + durationInMinutes * 60000).toISOString()
          });
        }
      }

      searchDate.setDate(searchDate.getDate() + 1);
    }

    return { availableSlots: availableSlots.slice(0, 5) };
  } catch (error) {
    console.error('findAvailableTime error:', error);
    return { error: error.message };
  }
}
async function createCalendarEvent({
  subject,
  startDateTime,
  endDateTime,
  attendees,
  optionalAttendees = [], // 👥 เพิ่มพารามิเตอร์ใหม่
  bodyContent,
  location,
  createMeeting = true,
  recurrence = null // 💡 เพิ่มพารามิเตอร์ใหม่
}) {
  try {
    const token = await getGraphToken();

    // --- ส่วนจัดการ Attendee ---
    let attendeeObjects = [];

    // 1. จัดการ Required Attendees
    const requiredLookups = await Promise.all(
      attendees.map(name => findUserByShortName(name.trim()))
    );

    let organizerEmail = null;
    if (requiredLookups.length > 0 && requiredLookups[0] && requiredLookups[0].length === 1) {
      organizerEmail = requiredLookups[0][0].userPrincipalName;
    } else {
      // ถ้าไม่เจอผู้จัดงานใน required list ให้ลองหาจาก optional list
      if (optionalAttendees.length > 0) {
        const optionalLookupsForOrganizer = await Promise.all(optionalAttendees.map(name => findUserByShortName(name.trim())));
        if (optionalLookupsForOrganizer.length > 0 && optionalLookupsForOrganizer[0] && optionalLookupsForOrganizer[0].length === 1) {
          organizerEmail = optionalLookupsForOrganizer[0][0].userPrincipalName;
        }
      }
    }

    if (!organizerEmail) {
      return { error: "ไม่สามารถระบุผู้จัดงาน (organizer) ที่ถูกต้องได้ กรุณาระบุผู้เข้าร่วมอย่างน้อย 1 คน" };
    }

    requiredLookups.forEach((users) => {
      if (users && users.length === 1) {
        attendeeObjects.push({
          emailAddress: { address: users[0].userPrincipalName, name: users[0].displayName },
          type: "required"
        });
      }
    });

    // 2. 🆕 จัดการ Optional Attendees
    if (optionalAttendees.length > 0) {
      const optionalLookups = await Promise.all(
        optionalAttendees.map(name => findUserByShortName(name.trim()))
      );
      optionalLookups.forEach((users) => {
        if (users && users.length === 1) {
          attendeeObjects.push({
            emailAddress: { address: users[0].userPrincipalName, name: users[0].displayName },
            type: "optional"
          });
        }
      });
    }

    // --- (ส่วนตรวจจับ CONFLICT สามารถลบออกได้ก่อนถ้ายังไม่ได้เพิ่มฟังก์ชัน find_available_time) ---
    // แต่ถ้าจะเก็บไว้ก็ไม่เป็นปัญหาครับ

    // --- สร้าง Request Body ---
    const event = {
      subject: subject,
      body: { contentType: "HTML", content: bodyContent || "" },
      start: { dateTime: startDateTime, timeZone: "Asia/Bangkok" },
      end: { dateTime: endDateTime, timeZone: "Asia/Bangkok" },
      location: { displayName: location || "" },
      attendees: attendeeObjects,
    };

    // 3. 🆕 เพิ่ม Recurrence เข้าไปใน event
    if (recurrence) {
      event.recurrence = recurrence;
    }

    if (createMeeting) {
      event.isOnlineMeeting = true;
      event.onlineMeetingProvider = "teamsForBusiness";
    }

    const url = `https://graph.microsoft.com/v1.0/users/${organizerEmail}/events`;
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(event)
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error('Graph API create event error:', errorText);
      return { error: `HTTP ${response.status}: ${errorText}` };
    }

    const createdEvent = await response.json();
    return {
      success: true,
      subject: createdEvent.subject,
      startTime: createdEvent.start.dateTime,
      organizer: createdEvent.organizer.emailAddress.name,
      webLink: createdEvent.webLink,
      meetingCreated: createdEvent.isOnlineMeeting || false
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
          "optionalAttendees": {
            type: "ARRAY",
            description: "A list of OPTIONAL attendees' names. Use for people who are invited but not required to come. Example: ['natsarin']",
            items: { type: "STRING" }
          },
          // 💡 เพิ่ม Property ใหม่สำหรับ Recurrence
          "recurrence": {
            type: "OBJECT",
            description: "Describes the recurrence pattern and range of the event. Use for events that repeat.",
            properties: {
              "pattern": {
                type: "OBJECT",
                properties: {
                  "type": { type: "STRING", enum: ["daily", "weekly", "absoluteMonthly", "relativeMonthly", "absoluteYearly", "relativeYearly"] },
                  "interval": { type: "NUMBER", description: "The number of units between occurrences. E.g., 1 for every week, 2 for every other week." },
                  "daysOfWeek": { type: "ARRAY", items: { type: "STRING" }, description: "e.g., ['monday', 'wednesday']" },
                  "dayOfMonth": { type: "NUMBER", description: "Day of the month (1-31) for monthly patterns." }
                },
                required: ["type", "interval"]
              },
              "range": {
                type: "OBJECT",
                properties: {
                  "type": { type: "STRING", enum: ["endDate", "noEnd", "numberedOccurrences"] },
                  "startDate": { type: "STRING", description: "The start date of the recurrence in YYYY-MM-DD format." },
                  "endDate": { type: "STRING", description: "The end date of the recurrence in YYYY-MM-DD format." },
                  "numberOfOccurrences": { type: "NUMBER", description: "The number of times the event repeats." }
                },
                required: ["type", "startDate"]
              }
            },
            required: ["pattern", "range"]
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
    const findAvailableTimeFunction = {
      name: "find_available_time",
      description: "ค้นหาช่วงเวลาที่ว่างตรงกันสำหรับผู้เข้าร่วมทั้งหมดภายในช่วงวันที่กำหนด ใช้เมื่อผู้ใช้ต้องการ 'หาเวลาว่าง', 'หาคิวว่าง', 'นัดประชุมตอนไหนดี'",
      parameters: {
        type: "OBJECT",
        properties: {
          "attendees": { type: "ARRAY", description: "รายชื่อผู้เข้าร่วมประชุม เช่น ['weraprat', 'natsarin']", items: { type: "STRING" } },
          "durationInMinutes": { type: "NUMBER", description: "ระยะเวลาประชุมที่ต้องการ (นาที) เช่น 30 หรือ 60" },
          "startSearch": { type: "STRING", description: "วันเริ่มต้นสำหรับค้นหาในรูปแบบ YYYY-MM-DD" },
          "endSearch": { type: "STRING", description: "วันสิ้นสุดสำหรับค้นหาในรูปแบบ YYYY-MM-DD" }
        },
        required: ["attendees", "durationInMinutes", "startSearch", "endSearch"]
      }
    };

    const systemInstruction = {
      parts: [{
        text: `You are Gent, a proactive and highly intelligent AI work assistant integrated into Microsoft Teams. Your primary goal is to facilitate seamless scheduling and calendar management for the team. You must respond in Thai.

---

### **Core Persona & Tone :**
- **Name:** Gent
- **Personality:** Professional, friendly, proactive, and a bit like a smart strategist.
- **Language:** Respond primarily in Thai (ตอบเป็นภาษาไทยเป็นหลัก). Be concise and clear.

---

### **Key Capabilities & Rules (ความสามารถและกฎการทำงาน):**
You have access to three main tools: \`get_user_calendar\`, \`find_available_time\`, and \`create_calendar_event\`.

1.  **Viewing Calendars (\`get_user_calendar\`):**
    * **RULE:** For simple requests to view schedules or events (e.g., "ดูตารางงานของ weraprat"), you **MUST** call the \`get_user_calendar\` function.

2.  **Finding Available Time (\`find_available_time\`):**
    * **CRITICAL RULE:** For ANY request to "find a time", "when are we free?", "หาเวลาว่างให้หน่อย", "หาคิวว่าง", you **MUST** call the \`find_available_time\` function.
    * **Action:** After the tool returns available slots, you MUST present these options to the user and ask which one they'd like to book.

3.  **Creating Events (\`create_calendar_event\`):**
    * **CRITICAL RULE:** For ANY request to book, schedule, create, or set up an event, meeting, or calendar block, you **MUST** call the \`create_calendar_event\` function.
    * **Recurrence:** You can now create repeating events. You MUST infer the recurrence pattern and range from user requests.
        * "ประชุมทุกวันจันทร์" -> \`recurrence: { pattern: { type: 'weekly', interval: 1, daysOfWeek: ['monday'] }, range: { type: 'noEnd', startDate: '...' } }\`
        * "Townhall ทุกวันที่ 15 ของเดือน" -> \`recurrence: { pattern: { type: 'absoluteMonthly', interval: 1, dayOfMonth: 15 }, range: { type: 'noEnd', startDate: '...' } }\`
    * **Attendees:** You can now distinguish between required and optional attendees.
        * "นัดประชุม weraprat ส่วนนัฏสรินทร์จะเข้าหรือไม่ก็ได้" -> \`attendees: ['weraprat']\`, \`optionalAttendees: ['natsarin']\`
        * If the user doesn't specify, assume everyone is **required**.
    * **PROACTIVE CONFLICT DETECTION:** The tool automatically checks for conflicts.
        * If the tool returns \`{ "conflict": true, "conflictingAttendees": ["User A"] }\`, it means the creation **failed** because those users are busy.
        * In this situation, you **MUST NOT** say the event was created. Instead, you must inform the user about the conflict and suggest a next action.
        * **Your Response MUST be:** "ไม่สามารถสร้างนัดหมายได้ครับ เนื่องจากคุณ [User A] มีนัดหมายอื่นคาบเกี่ยวอยู่ ต้องการให้ผมช่วยหาเวลาว่างอื่นให้แทนไหมครับ?" Then, use the \`find_available_time\` tool if the user agrees.
    * **Confirmation is Key:** Before calling the function, **summarize all details** (Subject, Time, Attendees, Recurrence) and **ask the user for confirmation**.

---

### **Important Context (ข้อมูลแวดล้อม):**
- **Current Date:** ${new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Bangkok' })}. Use this to resolve relative dates.

---

### **Response Formatting (รูปแบบการตอบ):**
- Use Markdown for clear formatting.
- **FORMAT:CARD:** Use for structured responses like lists or summaries.
- **FORMAT:TEXT:** Use for simple, conversational replies.
- **Always start your final response with either \`FORMAT:CARD\` or \`FORMAT:TEXT\`.**

---

### **Example Flow (ตัวอย่างการทำงาน):**

**Flow 1: Handling a booking conflict**
* **User:** "นัดประชุม Project X ตอนบ่ายสองพรุ่งนี้ให้หน่อย มีผมกับนัฏสรินทร์"
* **Your Thought Process:** User wants to create an event. I will call \`create_calendar_event\`.
* **System:** (Calls \`create_calendar_event\` tool. The tool finds a conflict and returns \`{ "conflict": true, "conflictingAttendees": ["Natsarin"] }\`)
* **Your Response (FORMAT:TEXT):** "ขออภัยครับ ไม่สามารถสร้างนัดหมายได้ เนื่องจากคุณ 'Natsarin' มีนัดหมายอื่นคาบเกี่ยวอยู่ตอนบ่ายสองพอดีครับ ต้องการให้ผมช่วยหาเวลาว่างอื่นสำหรับวันพรุ่งนี้ให้แทนไหมครับ?"

**Flow 2: Creating a Recurring Event**
* **User:** "นัด Sync ทีมหน่อย ทุกวันศุกร์ 4 โมงเย็น เริ่มศุกร์นี้เลย ส่วน Manager ให้เป็น optional นะ"
* **Your Thought Process:** This is a recurring event with an optional attendee. I need to build a recurrence object.
* **Your Confirmation (FORMAT:TEXT):** "รับทราบครับ ผมจะสร้างนัดหมาย 'Sync ทีม' ให้ทุกวันศุกร์ เวลา 16:00 - 17:00 น. โดยเริ่มตั้งแต่วันศุกร์นี้เป็นต้นไป และเชิญ Manager แบบ optional นะครับ ยืนยันไหมครับ?"
`
      }]
    };



    const model = genAI.getGenerativeModel({
      model: currentModel,
      tools: [{ functionDeclarations: [calendarFunction, createEventFunction, findAvailableTimeFunction] }], // <<-- ✅ ต้องมีครบ 3 ตัว
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
      } else if (call.name === "find_available_time") {
        const result = await findAvailableTime(call.args);

        const historyWithFunction = [
          ...conversationHistory,
          { role: "model", parts: [{ functionCall: call }] },
          { role: "function", parts: [{ functionResponse: { name: "find_available_time", response: result } }] }
        ];

        const finalResult = await model.generateContent({
          contents: historyWithFunction
        });
        text = finalResult.response.text();

      } else if (call.name === "create_calendar_event") {
        // รับค่าทั้งหมดจาก call.args ที่ Gemini ส่งมา (ซึ่งตอนนี้จะมี attendees ด้วย)
        const eventData = call.args;


        // ตรวจสอบข้อมูลสำคัญ รวมถึงเช็คว่ามี attendees อย่างน้อย 1 คนหรือไม่
        if (!eventData.subject || !eventData.startDateTime || !eventData.endDateTime || !eventData.attendees || eventData.attendees.length === 0) {
          text = "ขออภัยครับ ข้อมูลสำหรับสร้างนัดหมายไม่ครบถ้วน กรุณาระบุหัวข้อ, เวลาเริ่มต้น-สิ้นสุด, และผู้เข้าร่วมอย่างน้อย 1 คนครับ";
        } else {
          // เรียกใช้ฟังก์ชัน createCalendarEvent เวอร์ชันใหม่ที่รับ parameter แค่ตัวเดียว
          const createResult = await createCalendarEvent(eventData);

          // ตรวจสอบผลลัพธ์จากการสร้าง Event ก่อน
          if (createResult.error) {
            // ถ้าการสร้าง Event ล้มเหลว ก็แจ้ง Error ไปตามจริง
            text = `เกิดข้อผิดพลาดในการสร้างนัดหมายครับ: ${createResult.error}`;
          } else {
            // ถ้าการสร้าง Event สำเร็จ ให้เตรียมข้อความยืนยันพื้นฐานไว้ก่อน
            let successMessage = `เรียบร้อยครับ! ผมได้สร้างนัดหมาย '${eventData.subject}' และส่งคำเชิญให้แล้วครับ ✅`;

            // ตอนนี้ เราจะ "พยายาม" เรียก Gemini เพื่อสรุปผลให้สวยงาม
            try {
              const historyWithFunction = [
                ...conversationHistory,
                { role: "model", parts: [{ functionCall: call }] },
                { role: "function", parts: [{ functionResponse: { name: "create_calendar_event", response: createResult } }] }
              ];

              const finalResult = await model.generateContent({
                contents: historyWithFunction
              });

              const geminiText = finalResult.response?.text();

              // ถ้า Gemini ตอบกลับมาเป็นข้อความที่มีเนื้อหา ก็ใช้ข้อความนั้น
              if (geminiText) {
                text = geminiText;
              } else {
                // ถ้า Gemini ตอบกลับมาว่างเปล่า ก็ใช้ข้อความที่เราเตรียมไว้
                text = successMessage;
              }

            } catch (summarizationError) {
              // ถ้าการเรียก Gemini ครั้งที่สองนี้ "พัง" หรือ "แครช"
              console.error("Gemini summarization failed, using fallback message.", summarizationError);
              // ไม่ต้องสนใจ Error! ให้ใช้ข้อความยืนยันที่เราเตรียมไว้แทน เพื่อให้ผู้ใช้รู้ว่างานสำเร็จแล้ว
              text = successMessage;
            }
          }

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
    if (history.length > 40) { // Keep last 10 exchanges (20 messages)
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
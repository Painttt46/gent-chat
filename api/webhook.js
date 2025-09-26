import { GoogleGenerativeAI } from '@google/generative-ai';

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

// Send message to Teams incoming webhook
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

  // Check if user wants to broadcast to everyone (move outside try-catch)
  const shouldBroadcast = cleanText.toLowerCase().includes('(broadcast)');

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
    const currentModel = userModels.get(userId) || 'gemini-2.5-flash';
    return res.status(200).json({
      text: `Hi! I'm Gent, your AI work assistant in this Teams channel. How can I help you today?\n\nCommands:\n• 'clear' - reset conversation\n• 'model <name>' - switch AI model\n\nCurrent: ${models[currentModel].name} (${models[currentModel].count}/${models[currentModel].limit} requests) | API Key: ${currentApiKeyIndex + 1}/2`
    });
  }

  try {
    // Initialize Gemini AI with current API key
    const genAI = new GoogleGenerativeAI(getCurrentApiKey());
    const model = genAI.getGenerativeModel({ model: currentModel });

    // Check if user wants to broadcast to everyone
    // const shouldBroadcast = cleanText.toLowerCase().includes('(broadcast)'); // Moved outside try-catch
    
    // Remove (broadcast) from the message before processing
    const processedText = cleanText.replace(/\(broadcast\)/gi, '').trim();
    const finalText = processedText || cleanText;

    // Get or create conversation history for this user
    if (!conversations.has(userId)) {
      conversations.set(userId, []);
    }
    const history = conversations.get(userId);

    // Build conversation context with agent prompt
    let conversationContext = `You are Gent, an AI work assistant helping team members in a Microsoft Teams channel. 

Your role:
- Provide professional, helpful assistance to office workers
- Be friendly, concise, and actionable in your responses
- You're part of the team conversation in this Teams channel
- Help with work-related questions, productivity tips, and general office support
- You can research APIs and URLs when asked

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

Choose FORMAT:CARD when the response would look better with structured formatting.

`;

    // Add previous conversation history
    if (history.length > 0) {
      conversationContext += "Previous conversation in this channel:\n";
      history.forEach((msg, index) => {
        conversationContext += `${msg.role}: ${msg.content}\n`;
      });
      conversationContext += "\n";
    }

    conversationContext += `Current message from team member: ${finalText}`;

    // Check for "home" API command
    if (finalText.toLowerCase().includes('home')) {
      try {
        const apiData = await fetch('https://api.zippopotam.us/us/33162');
        const jsonData = await apiData.json();
        conversationContext += `\n\nHome API Response from https://api.zippopotam.us/us/33162:\n${JSON.stringify(jsonData, null, 2)}`;
      } catch (error) {
        conversationContext += `\n\nNote: Could not fetch home API data - ${error.message}`;
      }
    }

    // Check if user is asking about an API/URL and fetch it
    const urlRegex = /https?:\/\/[^\s]+/g;
    const urls = finalText.match(urlRegex);
    if (urls && (finalText.toLowerCase().includes('research') || finalText.toLowerCase().includes('api') || finalText.toLowerCase().includes('check'))) {
      try {
        const apiData = await fetch(urls[0]);
        const jsonData = await apiData.json();
        conversationContext += `\n\nAPI Response from ${urls[0]}:\n${JSON.stringify(jsonData, null, 2)}`;
      } catch (error) {
        conversationContext += `\n\nNote: Could not fetch data from ${urls[0]} - ${error.message}`;
      }
    }

    // Generate response
    const result = await model.generateContent(conversationContext);
    const response = result.response;
    const text = response.text();

    // Parse format choice
    const isCardFormat = text.startsWith('FORMAT:CARD');
    const isTextFormat = text.startsWith('FORMAT:TEXT');

    let cleanResponse = text;
    if (isCardFormat) {
      cleanResponse = text.replace('FORMAT:CARD', '').trim();
    } else if (isTextFormat) {
      cleanResponse = text.replace('FORMAT:TEXT', '').trim();
    }

    // Save to conversation history (keep last 10 messages)
    history.push({ role: "User", content: finalText });
    history.push({ role: "Gent", content: cleanResponse });
    if (history.length > 20) { // Keep last 10 exchanges (20 messages)
      history.splice(0, 2);
    }
    
    if (shouldBroadcast) {
      // Send to Teams incoming webhook with stats
      const broadcastMessage = `🔊 **Announcement from Gent:**\n\n${cleanResponse}\n\n💬 **${history.length / 2} messages** | **${models[currentModel].name}** | **${models[currentModel].count}/${models[currentModel].limit} requests** | **API ${currentApiKeyIndex + 1}/2**`;
      await sendToTeamsWebhook(broadcastMessage);
      
      // Return empty response (no reply to user)
      return res.status(200).json({});
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

    // If broadcast, send error to Teams webhook
    if (shouldBroadcast) {
      const errorMessage = `🔊 **Gent Error:**\n\nSorry, I'm having trouble right now. Please try again.\n\n💬 **${conversations.get(userId)?.length / 2 || 0} messages** | **${models[currentModel].name}** | **${models[currentModel].count}/${models[currentModel].limit} requests** | **API ${currentApiKeyIndex + 1}/2**`;
      await sendToTeamsWebhook(errorMessage);
      return res.status(200).json({});
    }

    res.status(200).json({
      text: `❌ **Gent:** Sorry, I'm having trouble right now. Please try again.\n\nError: ${error.message}`
    });
  }
}

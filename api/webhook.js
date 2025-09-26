import { GoogleGenerativeAI } from '@google/generative-ai';

// Simple in-memory conversation storage (per user)
const conversations = new Map();

// Model tracking
const models = {
  'gemini-2.5-flash': { name: 'Gemini 2.5 Flash', count: 0 },
  'gemini-2.5-pro': { name: 'Gemini 2.5 Pro', count: 0 }
};
const userModels = new Map(); // Track current model per user
let lastResetDate = new Date().toDateString();

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
        text: `🤖 Switched to ${models[modelKey].name} (${models[modelKey].count} requests)`
      });
    } else {
      const modelList = Object.entries(models).map(([key, model]) => 
        `• ${key} - ${model.name} (${model.count} requests)`
      ).join('\n');
      return res.status(200).json({
        text: `❌ Invalid model. Available models:\n${modelList}\n\nTotal: ${getTotalRequests()} requests\nUsage: model gemini-2.5-flash`
      });
    }
  }

  if (!cleanText) {
    const currentModel = userModels.get(userId) || 'gemini-2.5-flash';
    return res.status(200).json({
      text: `Hi! I'm Gent, your AI work assistant in this Teams channel. How can I help you today?\n\nCommands:\n• 'clear' - reset conversation\n• 'model <name>' - switch AI model\n\nCurrent: ${models[currentModel].name} (${models[currentModel].count} requests)`
    });
  }

  try {
    // Initialize Gemini AI
    const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
    const model = genAI.getGenerativeModel({ model: currentModel });

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

    conversationContext += `Current message from team member: ${cleanText}`;

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
    history.push({ role: "User", content: cleanText });
    history.push({ role: "Gent", content: cleanResponse });
    if (history.length > 20) { // Keep last 10 exchanges (20 messages)
      history.splice(0, 2);
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
                text: `💬 **${history.length / 2} messages** | **${models[currentModel].name}** | **${models[currentModel].count} requests**`,
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
        text: `🤖 **Gent:** ${cleanResponse}\n\n💬 **${history.length / 2} messages** | **${models[currentModel].name}** | **${models[currentModel].count} requests**`
      });
    }

  } catch (error) {
    console.error('Gemini API error:', error);

    res.status(200).json({
      text: `❌ **Gent:** Sorry, I'm having trouble right now. Please try again.\n\nError: ${error.message}`
    });
  }
}

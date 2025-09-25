import { GoogleGenerativeAI } from '@google/generative-ai';

// Simple in-memory conversation storage (per user)
const conversations = new Map();

export default async function handler(req, res) {
  console.log('Method:', req.method);
  console.log('Body:', req.body);

  // Get user ID from Teams (fallback to 'default' if not available)
  const userId = req.body?.from?.id || req.body?.channelData?.tenant?.id || 'default';

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

  if (!cleanText) {
    return res.status(200).json({
      text: "Hi! I'm your work assistant. How can I help you today? (Type 'clear' to reset conversation)"
    });
  }

  try {
    // Initialize Gemini AI
    const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
    const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });

    // Get or create conversation history for this user
    if (!conversations.has(userId)) {
      conversations.set(userId, []);
    }
    const history = conversations.get(userId);

    // Build conversation context
    let conversationContext = "You are a helpful work assistant for office workers. Provide practical, professional advice and answers. Keep responses concise and actionable.\n\n";
    
    // Add previous conversation history
    if (history.length > 0) {
      conversationContext += "Previous conversation:\n";
      history.forEach((msg, index) => {
        conversationContext += `${msg.role}: ${msg.content}\n`;
      });
      conversationContext += "\n";
    }
    
    conversationContext += `Current question: ${cleanText}`;

    // Generate response
    const result = await model.generateContent(conversationContext);
    const response = result.response;
    const text = response.text();

    // Save to conversation history (keep last 10 messages)
    history.push({ role: "User", content: cleanText });
    history.push({ role: "Assistant", content: text });
    if (history.length > 20) { // Keep last 10 exchanges (20 messages)
      history.splice(0, 2);
    }

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
              text: "🤖 Work Assistant",
              weight: "Bolder",
              size: "Medium",
              color: "Accent"
            },
            {
              type: "TextBlock",
              text: `**Q:** ${cleanText}`,
              wrap: true,
              spacing: "Medium"
            },
            {
              type: "TextBlock",
              text: text,
              wrap: true,
              spacing: "Medium"
            },
            {
              type: "TextBlock",
              text: `💬 Messages in conversation: ${history.length / 2} | Type 'clear' to reset`,
              size: "Small",
              color: "Accent",
              spacing: "Medium"
            }
          ]
        }
      }]
    });

  } catch (error) {
    console.error('Gemini API error:', error);
    
    res.status(200).json({
      text: `❌ Sorry, I'm having trouble right now. Please try again.\n\nError: ${error.message}`
    });
  }
}

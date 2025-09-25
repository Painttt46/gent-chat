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
      text: "Hi! I'm Gent, your AI work assistant in this Teams channel. How can I help you today? (Type 'clear' to reset conversation)"
    });
  }

  try {
    // Initialize Gemini AI
    const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
    const model = genai.GenerativeModel('gemini-2.5-pro')

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
                text: `💬 ${history.length / 2} messages in conversation | Type 'clear' to reset`,
                size: "Small",
                color: "Accent",
                spacing: "Medium"
              }
            ]
          }
        }]
      });
    } else {
      // Return as simple text
      res.status(200).json({
        text: `🤖 **Gent:** ${cleanResponse}\n\n💬 ${history.length / 2} messages | Type 'clear' to reset`
      });
    }

  } catch (error) {
    console.error('Gemini API error:', error);

    res.status(200).json({
      text: `❌ **Gent:** Sorry, I'm having trouble right now. Please try again.\n\nError: ${error.message}`
    });
  }
}

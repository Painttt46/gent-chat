import { GoogleGenerativeAI } from '@google/generative-ai';

export default async function handler(req, res) {
  console.log('Method:', req.method);
  console.log('Body:', req.body);

  // Clean mention from text
  let cleanText = req.body?.text || '';
  cleanText = cleanText.replace(/<at>.*?<\/at>/g, '');
  cleanText = cleanText.replace(/<[^>]*>/g, '');
  cleanText = cleanText.replace(/&nbsp;/g, ' ');
  cleanText = cleanText.replace(/&amp;/g, '&');
  cleanText = cleanText.replace(/&lt;/g, '<');
  cleanText = cleanText.replace(/&gt;/g, '>');
  cleanText = cleanText.replace(/\s+/g, ' ').trim();

  if (!cleanText) {
    return res.status(200).json({
      text: "Hi! I'm your work assistant. How can I help you today?"
    });
  }

  try {
    // Initialize Gemini AI (following official quickstart)
    const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

    const prompt = `You are a helpful work assistant for office workers. 
    Provide practical, professional advice and answers.
    Keep responses concise and actionable.
    
    User question: ${cleanText}`;

    // Generate content (following official pattern)
    const result = await model.generateContent(prompt);
    const response = result.response;
    const text = response.text();

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

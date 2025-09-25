export default async function handler(req, res) {
  console.log('Method:', req.method);
  console.log('Body:', req.body);

  // Clean mention from text
  let cleanText = req.body?.text || '';
  // Remove HTML at-mentions
  cleanText = cleanText.replace(/<at>.*?<\/at>/g, '');
  // Remove HTML tags
  cleanText = cleanText.replace(/<[^>]*>/g, '');
  // Decode HTML entities
  cleanText = cleanText.replace(/&nbsp;/g, ' ');
  cleanText = cleanText.replace(/&amp;/g, '&');
  cleanText = cleanText.replace(/&lt;/g, '<');
  cleanText = cleanText.replace(/&gt;/g, '>');
  // Clean whitespace
  cleanText = cleanText.replace(/\s+/g, ' ').trim();

  // TODO: Add Gemini API integration here
  // For now, just echo the message
  res.status(200).json({
    text: `You said: "${cleanText}"`
  });
}

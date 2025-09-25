export default function handler(req, res) {
  console.log('Method:', req.method);
  console.log('Body:', req.body);

  // Clean mention from text
  let cleanText = req.body?.text || '';
  // Remove HTML at-mentions
  cleanText = cleanText.replace(/<at>.*?<\/at>/g, '');
  // Remove HTML tags
  cleanText = cleanText.replace(/<[^>]*>/g, '');
  // Decode HTML entities like &nbsp;
  cleanText = cleanText.replace(/&nbsp;/g, ' ');
  cleanText = cleanText.replace(/&amp;/g, '&');
  cleanText = cleanText.replace(/&lt;/g, '<');
  cleanText = cleanText.replace(/&gt;/g, '>');
  // Remove ALL whitespace characters and replace with single space
  cleanText = cleanText.replace(/\s+/g, ' ');
  // Trim spaces
  cleanText = cleanText.trim();

  // Check if it's exactly "form"
  let isForm = cleanText.toLowerCase() === 'form';

  if (isForm) {
    // Return adaptive card form
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
              text: "Feedback Form",
              weight: "Bolder",
              size: "Medium"
            },
            {
              type: "Input.Text",
              id: "name",
              placeholder: "Your name"
            },
            {
              type: "Input.Text",
              id: "email",
              placeholder: "Your email"
            },
            {
              type: "Input.ChoiceSet",
              id: "rating",
              style: "compact",
              placeholder: "Rate our service",
              choices: [
                { title: "Excellent", value: "5" },
                { title: "Good", value: "4" },
                { title: "Average", value: "3" },
                { title: "Poor", value: "2" },
                { title: "Very Poor", value: "1" }
              ]
            },
            {
              type: "Input.Text",
              id: "feedback",
              placeholder: "Your feedback",
              isMultiline: true
            }
          ],
          actions: [
            {
              type: "Action.Submit",
              title: "Submit"
            }
          ]
        }
      }]
    });
  } else {
    // Return custom message format
    res.status(200).json({
      text: `The message going out is "${cleanText}"`
    });
  }
}
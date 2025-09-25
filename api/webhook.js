export default async function handler(req, res) {
  console.log('=== WEBHOOK CALLED ===');
  console.log('Method:', req.method);
  console.log('Headers:', JSON.stringify(req.headers, null, 2));
  console.log('Body:', JSON.stringify(req.body, null, 2));
  console.log('Query:', req.query);
  console.log('========================');

  // Handle form submission - Teams sends in different formats
  if (req.body?.value || req.body?.type === 'invoke') {
    let formData;
    
    // Handle different Teams submission formats
    if (req.body?.value?.action?.data) {
      formData = req.body.value.action.data;
    } else if (req.body?.value) {
      formData = req.body.value;
    } else if (req.body?.data) {
      formData = req.body.data;
    }
    
    console.log('Form submission received:', formData);
    
    try {
      // Test MongoDB connection and save
      console.log('Attempting to connect to MongoDB...');
      await saveToDatabase(formData);
      console.log('Successfully saved to MongoDB');
      
      return res.status(200).json({
        text: `✅ Thank you ${formData.name}! Your feedback has been saved to MongoDB.`
      });
    } catch (error) {
      console.error('MongoDB connection/save error:', error.message);
      return res.status(200).json({
        text: `❌ Database Error: ${error.message}. Please check MongoDB connection.`
      });
    }
  }

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
              title: "Submit",
              data: {
                "submitType": "feedback"
              }
            }
          ]
        }
      }]
    });
  } else {
    // Return custom message format
    res.status(200).json({
      text: `The message is "${cleanText}"`
    });
  }
}

// Database save function with better error handling
async function saveToDatabase(formData) {
  const { MongoClient } = require('mongodb');
  
  console.log('MongoDB URI:', process.env.MONGODB_URI ? 'Set' : 'Not set');
  
  if (!process.env.MONGODB_URI) {
    throw new Error('MONGODB_URI environment variable not set');
  }
  
  const client = new MongoClient(process.env.MONGODB_URI);
  
  try {
    console.log('Connecting to MongoDB...');
    await client.connect();
    console.log('Connected successfully');
    
    const db = client.db('feedback');
    const result = await db.collection('submissions').insertOne({
      name: formData.name,
      email: formData.email,
      rating: formData.rating,
      feedback: formData.feedback,
      timestamp: new Date()
    });
    
    console.log('Document inserted with ID:', result.insertedId);
    return result;
  } catch (error) {
    console.error('MongoDB operation failed:', error);
    throw error;
  } finally {
    await client.close();
    console.log('MongoDB connection closed');
  }
}

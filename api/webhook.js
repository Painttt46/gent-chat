export default async function handler(req, res) {
  // A) ถ้าเป็นการ Submit Adaptive Card
  if (req.body?.type === "invoke" || req.body?.value) {
    let formData =
      req.body?.value?.action?.data ??
      req.body?.value?.data ??
      req.body?.value ??
      req.body?.data ??
      {};

    console.log("Form submission:", formData);

    try {
      await saveToDatabase(formData);
      return res.status(200).json({
        statusCode: 200,
        type: "application/vnd.microsoft.activity",
        value: {
          type: "message",
          text: `✅ Thanks ${formData?.name ?? ""}! Your feedback has been saved.`
        }
      });
    } catch (err) {
      console.error("DB error:", err);
      return res.status(200).json({
        statusCode: 200,
        type: "application/vnd.microsoft.activity",
        value: {
          type: "message",
          text: `❌ Database Error: ${err.message}`
        }
      });
    }
  }

  // B) ถ้าเป็นข้อความ "form" → ส่ง Adaptive Card
  let cleanText = (req.body?.text || "")
    .replace(/<at>.*?<\/at>/g, "")
    .replace(/<[^>]*>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/\s+/g, " ")
    .trim();

  if (cleanText.toLowerCase() === "form") {
    return res.status(200).json({
      type: "message",
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: {
            $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
            type: "AdaptiveCard",
            version: "1.4",
            body: [
              { type: "TextBlock", text: "Feedback Form", weight: "Bolder", size: "Medium" },
              { type: "Input.Text", id: "name", placeholder: "Your name" },
              { type: "Input.Text", id: "email", placeholder: "Your email" },
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
              { type: "Input.Text", id: "feedback", placeholder: "Your feedback", isMultiline: true }
            ],
            actions: [{ type: "Action.Submit", title: "Submit", data: { submitType: "feedback" } }]
          }
        }
      ]
    });
  }

  // C) ข้อความทั่วไป
  return res.status(200).json({ text: `The message is "${cleanText}"` });
}

// บันทึกลง MongoDB
async function saveToDatabase(formData) {
  const { MongoClient } = require("mongodb");
  if (!process.env.MONGODB_URI) throw new Error("MONGODB_URI not set");

  const client = new MongoClient(process.env.MONGODB_URI);
  try {
    await client.connect();
    const db = client.db("feedback");
    return await db.collection("submissions").insertOne({
      name: formData?.name,
      email: formData?.email,
      rating: formData?.rating,
      feedback: formData?.feedback,
      timestamp: new Date()
    });
  } finally {
    await client.close();
  }
}

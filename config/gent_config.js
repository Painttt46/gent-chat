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


const systemInstruction = {
    parts: [{
        text:
            `You are Gent, a proactive and highly intelligent AI work assistant integrated into Microsoft Teams. Your primary goal is to facilitate seamless scheduling and calendar management for the team. You must respond in Thai.

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
export { calendarFunction, createEventFunction, systemInstruction };

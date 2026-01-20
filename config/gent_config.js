const calendarFunction = {
    name: "get_user_calendar",
    description: `ดูตารางนัดหมายของพนักงาน ใช้เมื่อถามว่า "ดูตาราง", "วันนี้มีนัดอะไร", "ตารางงานสัปดาห์นี้"
      - ถ้าไม่ระบุวันที่ จะแสดงวันนี้
      - เข้าใจคำว่า "วันนี้", "พรุ่งนี้", "สัปดาห์นี้", "เดือนนี้"
      - ต้องแปลงวันที่เป็น YYYY-MM-DD ก่อนเรียก`,
    parameters: {
        type: "OBJECT",
        properties: {
            "userPrincipalName": {
                type: "STRING",
                description: "ชื่อพนักงาน เช่น 'weraprat', 'natsarin'"
            },
            "startDate": {
                type: "STRING",
                description: "วันเริ่มต้น YYYY-MM-DD (ถ้าไม่ระบุ = วันนี้)"
            },
            "endDate": {
                type: "STRING",
                description: "วันสิ้นสุด YYYY-MM-DD (ถ้าไม่ระบุ = เท่ากับ startDate)"
            }
        },
        required: ["userPrincipalName"]
    }
};

/*
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

const findAvailableTimeFunction = {
    name: "find_available_time",
    description: "ค้นหาช่วงเวลาที่ว่างตรงกันสำหรับผู้เข้าร่วมทั้งหมดภายในช่วงวันที่กำหนด ใช้เมื่อผู้ใช้ต้องการ 'หาเวลาว่าง', 'หาคิวว่าง', 'นัดประชุมตอนไหนดี'",
    parameters: {
        type: "OBJECT",
        properties: {
            "attendees": { type: "ARRAY", description: "รายชื่อผู้เข้าร่วมประชุม เช่น ['weraprat', 'natsarin']", items: { type: "STRING" } },
            "durationInMinutes": { type: "NUMBER", description: "ระยะเวลาประชุมที่ต้องการ (นาที) เช่น 30 หรือ 60" },
            "startSearch": { type: "STRING", description: "วันเริ่มต้นสำหรับค้นหาในรูปแบบ YYYY-MM-DD" },
            "endSearch": { type: "STRING", description: "วันสิ้นสุดสำหรับค้นหาในรูปแบบ YYYY-MM-DD" }
        },
        required: ["attendees", "durationInMinutes", "startSearch", "endSearch"]
    }
};
*/

const systemInstruction = {
    parts: [{
        text:
            `You are Gent, a proactive and highly intelligent AI work assistant integrated into Microsoft Teams. Your goal is not just to answer, but to solve problems and manage tasks efficiently.

        ---

        ### 1. Core Persona (ตัวตนและหน้าที่):
        - **Name:** Gent (เจนท์)
        - **Role:** เลขาส่วนตัวอัจฉริยะ (AI Executive Assistant) ที่รู้ใจและทำงานไว
        - **Tone:** มืออาชีพ (Professional), มั่นใจ (Confident), กระตือรือร้น (Proactive), และเป็นมิตร (Friendly)
        - **Language:** ตอบเป็น **ภาษาไทยธุรกิจ** ที่สละสลวย กระชับ และสุภาพ (ใช้ "ครับ" เสมอ)
        - **Mindset:** "Think Ahead" - อย่ารอให้สั่งทุกอย่าง ถ้าเห็นว่าอะไรจำเป็น ให้เสนอทันที

        ---

        ### 2. Advanced Document Analysis Rules (กฎการอ่านเอกสารขั้นสูง - สำคัญมาก):
        1. **Thai Numerals (เลขไทย):** คุณต้องแปลงเลขไทย (๐-๙) ในภาพเอกสารให้เป็นเลขอารบิก (0-9) เสมอ โดยเฉพาะใน "วันที่", "จำนวนเงิน", และ "เลขที่สัญญา"
        2. **Handling Document Typos:** เอกสารราชการไทยมักมีการพิมพ์เลขข้อผิด (เช่น ข้ามจากข้อ 6 ไป 17 หรือพิมพ์ข้อ 4 ซ้ำ)
           - **กฎเหล็ก:** ห้ามหยุดอ่านหรือตัดเนื้อหาทิ้งเพียงเพราะเลขข้อไม่เรียงกัน!
           - ให้ดึง "หัวข้อ" (Header) ทั้งหมดออกมาตามจริงที่ปรากฏในภาพ หากเลขข้อผิดให้ระบุในวงเล็บ เช่น "ข้อ 17 (ในเอกสารพิมพ์ผิด น่าจะเป็นข้อ 7)"
        3. **Exhaustive Reading:** ต้องอ่านเอกสารให้ครบทุกหน้าและทุกบรรทัด ห้ามสรุปเอาเองว่าจบแล้วจนกว่าจะถึงหน้าสุดท้ายหรือลายเซ็นคู่สัญญา
        4. **Table Structure:** หากเจอภาพตาราง ให้พยายามรักษาโครงสร้างแถวและคอลัมน์ (Row/Column) เมื่อสรุปข้อมูล อย่าเอาตัวเลขมาปนกันมั่ว
        5. **Form Headers:** สังเกตหัวกระดาษ (Header) และลายเซ็นท้ายกระดาษ เพื่อระบุประเภทเอกสาร (เช่น ใบสั่งซื้อ, สัญญา, ใบเสนอราคา)

        ---

        ### 3. Continuous Context & Proactive Rules (การทำงานต่อเนื่องเชิงรุก):
        1. **Context Retention (การจำบริบทเอกสาร):**
           - เมื่ออ่านไฟล์เสร็จแล้ว หากผู้ใช้ถามต่อ (เช่น "สัญญาหมดอายุเมื่อไหร่?", "มีค่าปรับไหม?") **ห้าม** ถามกลับว่า "ไฟล์ไหน?"
           - ให้ใช้ข้อมูลจากไฟล์ล่าสุดที่เพิ่งอ่านไปตอบทันที
        2. **Anticipate Needs (คาดเดาความต้องการ):**
           - อ่านสัญญาเสร็จ -> เสนอ "ต้องการให้ลงนัดหมายวันส่งมอบในปฏิทินเลยไหมครับ?"
           - อ่านใบเสนอราคาเสร็จ -> เสนอ "ต้องการสรุปยอดรวมเพื่อทำเรื่องเบิกไหมครับ?"
        3. **Error Recovery:**
           - ถ้าหาข้อมูลไม่เจอ ห้ามตอบ Error Code ให้ถามข้อมูลเพิ่มอย่างสุภาพ เช่น "ขออภัยครับ ไม่พบข้อมูลโครงการนี้ ยืนยันว่าเป็นเลข SO นี้ใช่ไหมครับ?"

        ---

        ### 4. Thinking Process (กระบวนการคิดก่อนตอบ):
        ก่อนตอบทุกครั้ง ให้ประมวลผลตามลำดับนี้:
        1. **Analyze Intent:** ผู้ใช้ต้องการ "ข้อมูล" (Fact), "การกระทำ" (Action), หรือ "คำแนะนำ" (Advice)?
        2. **Check Context:** มีข้อมูลเก่าที่ต้องใช้ไหม? (เช่น ไฟล์สัญญาที่เพิ่งอ่าน, เลข SO ที่คุยค้างไว้)
        3. **Select Tools:** ต้องใช้ Tool ตัวไหน? (ถ้าข้อมูลมีอยู่แล้วใน History ไม่ต้องเรียก Tool ซ้ำ ให้ตอบได้เลย)
        4. **Formulate Response:** สรุปผลลัพธ์เป็น **FORMAT:CARD** หรือ **FORMAT:TEXT** ให้อ่านง่ายที่สุด

        ---

        ### ข้อมูลแวดล้อม:
        - **วันที่ปัจจุบัน:** ${new Date().toLocaleDateString('th-TH', { timeZone: 'Asia/Bangkok', year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' })}

        ---

        ### รูปแบบการตอบ (Response Formatting):
        - **FORMAT:CARD**
           - ใช้เมื่อ: สรุปข้อมูลโครงการ, รายการนัดหมาย, สาระสำคัญของสัญญา, หรือข้อมูลที่มีหัวข้อย่อยเยอะๆ
           - สไตล์: ใช้ Markdown List, Bold ตัวเลขสำคัญ, แยกบรรทัดให้ชัดเจน
        - **FORMAT:TEXT**
           - ใช้เมื่อ: ตอบคำถามทั่วไป, ยืนยันการทำรายการ, หรือคุยเล่น
           - สไตล์: กระชับ เป็นธรรมชาติ

        **สำคัญ:** ขึ้นต้นคำตอบด้วย \`FORMAT:CARD\` หรือ \`FORMAT:TEXT\` เสมอ เพื่อให้ระบบแสดงผลถูกต้อง
        `
    }]
};
export { calendarFunction, systemInstruction };

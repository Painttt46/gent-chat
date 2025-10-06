
// services/graph.service.js

import { ConfidentialClientApplication } from '@azure/msal-node';
import { fromZonedTime, toZonedTime } from 'date-fns-tz';
import { parseISO, startOfDay, endOfDay } from 'date-fns';

let cachedGraphToken = {
    token: null,
    expiresOn: null
};

// This function is internal to this service and doesn't need to be exported.
async function getGraphToken() {
    if (cachedGraphToken.token && new Date() < cachedGraphToken.expiresOn) {
        console.log('Using cached Graph token');
        return cachedGraphToken.token;
    }
    try {
        const clientConfig = {
            auth: {
                clientId: process.env.AZURE_CLIENT_ID,
                clientSecret: process.env.AZURE_CLIENT_SECRET,
                authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`
            }
        };
        const cca = new ConfidentialClientApplication(clientConfig);
        const clientCredentialRequest = {
            scopes: ['https://graph.microsoft.com/.default']
        };
        const response = await cca.acquireTokenByClientCredential(clientCredentialRequest);
        if (response && response.accessToken && response.expiresOn) {
            cachedGraphToken.token = response.accessToken;
            cachedGraphToken.expiresOn = new Date(response.expiresOn.getTime() - 5 * 60 * 1000);
            console.log('New Graph token acquired and cached.');
            return response.accessToken;
        } else {
            throw new Error('Failed to acquire token or token response is invalid.');
        }
    } catch (error) {
        console.error('Token acquisition error:', error);
        cachedGraphToken.token = null;
        cachedGraphToken.expiresOn = null;
        throw error;
    }
}

export async function findUserByShortName(name) {
    try {
        const token = await getGraphToken();
        const filterQuery = `$filter=displayName eq '${name}' or givenName eq '${name}' or surname eq '${name}' or mailNickname eq '${name}'`;
        const selectQuery = `&$select=displayName,userPrincipalName`;
        const url = `https://graph.microsoft.com/v1.0/users?${filterQuery}${selectQuery}`;
        console.log('Searching for user with URL:', url);
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        if (!response.ok) {
            const errorText = await response.text();
            console.error('Graph API user search error:', errorText);
            return { error: `HTTP ${response.status}: ${errorText}` };
        }
        const data = await response.json();
        return data.value;
    } catch (error) {
        console.error('findUserByShortName error:', error);
        return { error: error.message };
    }
}

export async function getUserCalendar(nameOrEmail, startDate = null, endDate = null) {
    let userEmail = nameOrEmail;
    if (!nameOrEmail.includes('@')) {
        console.log(`Searching for user: '${nameOrEmail}'`);
        const users = await findUserByShortName(nameOrEmail);
        if (!users || users.length === 0) {
            return { error: `ไม่พบผู้ใช้ที่ชื่อ '${nameOrEmail}' ในระบบครับ` };
        }
        if (users.length > 1) {
            const userList = users.map(u => u.displayName).join(', ');
            return { error: `พบผู้ใช้ที่ชื่อ '${nameOrEmail}' มากกว่า 1 คน: ${userList} กรุณาระบุให้ชัดเจนขึ้นครับ` };
        }
        userEmail = users[0].userPrincipalName;
        console.log(`User found: ${users[0].displayName} (${userEmail})`);
    }
    try {
        const token = await getGraphToken();
        let startDateTime, endDateTime;
        const bangkokTz = 'Asia/Bangkok';
        if (startDate && endDate) {
            const start = fromZonedTime(startOfDay(parseISO(startDate)), bangkokTz);
            const end = fromZonedTime(endOfDay(parseISO(endDate)), bangkokTz);
            startDateTime = start.toISOString();
            endDateTime = end.toISOString();
        } else if (startDate) {
            const start = fromZonedTime(startOfDay(parseISO(startDate)), bangkokTz);
            const end = fromZonedTime(endOfDay(parseISO(startDate)), bangkokTz);
            startDateTime = start.toISOString();
            endDateTime = end.toISOString();
        } else {
            const now = new Date();
            const bangkokNow = toZonedTime(now, bangkokTz);
            const start = fromZonedTime(startOfDay(bangkokNow), bangkokTz);
            const end = fromZonedTime(endOfDay(bangkokNow), bangkokTz);
            startDateTime = start.toISOString();
            endDateTime = end.toISOString();
        }
        const url = `https://graph.microsoft.com/v1.0/users/${userEmail}/calendarView?startDateTime=${startDateTime}&endDateTime=${endDateTime}&$select=subject,body,bodyPreview,organizer,attendees,start,end,location,onlineMeeting,isAllDay`;
        const response = await fetch(url, {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Prefer': 'outlook.timezone="Asia/Bangkok"'
            }
        });
        if (!response.ok) {
            const errorText = await response.text();
            console.error('Graph API error response:', errorText);
            return { error: `ไม่สามารถดึงข้อมูลปฏิทินของ ${userEmail} ได้ครับ` };
        }
        const data = await response.json();


        if (data.value && startDate) { // ใช้ตัวแปร startDate ที่เป็น string 'YYYY-MM-DD'
            data.value = data.value.filter(event => {
                // กรณีเป็นกิจกรรมเต็มวัน (All Day)
                if (event.isAllDay) {
                    // เพิ่ม .substring(0, 10) เพื่อตัดเอาเฉพาะวันที่ 'YYYY-MM-DD' มาเปรียบเทียบ
                    return event.end.dateTime.substring(0, 10) > startDate;
                }

                else {
                    const queryStartBoundary = new Date(startDateTime); // 'startDateTime' คือเวลาเต็มรูปแบบ ISO
                    const eventEnd = new Date(event.end.dateTime);

                    return eventEnd > queryStartBoundary;
                }
            });
        }


        return data;

    } catch (error) {
        console.error('Graph API error:', error);
        return { error: error.message };
    }

}

export async function findAvailableTime({ attendees, durationInMinutes, startSearch, endSearch }) {
    // The full implementation of this function is long, so it's copied here directly.
    // ... (Paste the entire findAvailableTime function code from the original file here)
    try {
        console.log('Finding available time for:', { attendees, durationInMinutes, startSearch, endSearch });
        const bangkokTz = 'Asia/Bangkok';
        const userLookups = await Promise.all(
            attendees.map(name => findUserByShortName(name.trim()))
        );
        const resolvedUsers = [];
        const unresolvedNames = [];
        userLookups.forEach((users, index) => {
            if (users && users.length === 1) {
                resolvedUsers.push(users[0]);
            } else {
                unresolvedNames.push(attendees[index]);
            }
        });
        if (unresolvedNames.length > 0) {
            return { error: `ไม่พบผู้ใช้: ${unresolvedNames.join(', ')}` };
        }
        if (resolvedUsers.length === 0) {
            return { error: 'ไม่พบรายชื่อผู้เข้าร่วมที่ถูกต้อง' };
        }
        const calendarPromises = resolvedUsers.map(user =>
            getUserCalendar(user.userPrincipalName, startSearch, endSearch)
        );
        const calendarResults = await Promise.all(calendarPromises);
        let allBusySlots = [];
        for (const result of calendarResults) {
            if (result.value) {
                result.value.forEach(event => {
                    allBusySlots.push({
                        start: new Date(event.start.dateTime),
                        end: new Date(event.end.dateTime)
                    });
                });
            }
        }
        allBusySlots.sort((a, b) => a.start - b.start);
        const availableSlots = [];
        const workingHoursStart = 9;
        const workingHoursEnd = 18;
        let searchDate = fromZonedTime(startOfDay(parseISO(startSearch)), bangkokTz);
        const searchEndDate = fromZonedTime(endOfDay(parseISO(endSearch)), bangkokTz);
        while (searchDate <= searchEndDate && availableSlots.length < 5) {
            let potentialSlotStart = toZonedTime(searchDate, bangkokTz);
            potentialSlotStart.setHours(workingHoursStart, 0, 0, 0);
            const dayEnd = toZonedTime(searchDate, bangkokTz);
            dayEnd.setHours(workingHoursEnd, 0, 0, 0);
            const todayBusySlots = allBusySlots.filter(slot =>
                startOfDay(toZonedTime(slot.start, bangkokTz)).getTime() === startOfDay(searchDate).getTime()
            );
            for (const busySlot of todayBusySlots) {
                const gapMillis = busySlot.start - potentialSlotStart;
                const gapMinutes = Math.floor(gapMillis / (1000 * 60));
                if (gapMinutes >= durationInMinutes) {
                    availableSlots.push({
                        start: potentialSlotStart.toISOString(),
                        end: new Date(potentialSlotStart.getTime() + durationInMinutes * 60000).toISOString()
                    });
                    if (availableSlots.length >= 5) break;
                }
                potentialSlotStart = new Date(Math.max(potentialSlotStart, busySlot.end));
            }
            if (availableSlots.length < 5 && potentialSlotStart < dayEnd) {
                const finalGapMillis = dayEnd - potentialSlotStart;
                const finalGapMinutes = Math.floor(finalGapMillis / (1000 * 60));
                if (finalGapMinutes >= durationInMinutes) {
                    availableSlots.push({
                        start: potentialSlotStart.toISOString(),
                        end: new Date(potentialSlotStart.getTime() + durationInMinutes * 60000).toISOString()
                    });
                }
            }
            searchDate.setDate(searchDate.getDate() + 1);
        }
        return { availableSlots: availableSlots.slice(0, 5) };
    } catch (error) {
        console.error('findAvailableTime error:', error);
        return { error: error.message };
    }
}

export async function createCalendarEvent({
    subject,
    startDateTime,
    endDateTime,
    attendees,
    optionalAttendees = [],
    bodyContent,
    location,
    createMeeting = true,
    recurrence = null
}) {
    // The full implementation of this function is long, so it's copied here directly.
    // ... (Paste the entire createCalendarEvent function code from the original file here)
    try {
        const token = await getGraphToken();
        let attendeeObjects = [];
        const requiredLookups = await Promise.all(
            attendees.map(name => findUserByShortName(name.trim()))
        );
        let organizerEmail = null;
        if (requiredLookups.length > 0 && requiredLookups[0] && requiredLookups[0].length === 1) {
            organizerEmail = requiredLookups[0][0].userPrincipalName;
        } else {
            if (optionalAttendees.length > 0) {
                const optionalLookupsForOrganizer = await Promise.all(optionalAttendees.map(name => findUserByShortName(name.trim())));
                if (optionalLookupsForOrganizer.length > 0 && optionalLookupsForOrganizer[0] && optionalLookupsForOrganizer[0].length === 1) {
                    organizerEmail = optionalLookupsForOrganizer[0][0].userPrincipalName;
                }
            }
        }
        if (!organizerEmail) {
            return { error: "ไม่สามารถระบุผู้จัดงาน (organizer) ที่ถูกต้องได้ กรุณาระบุผู้เข้าร่วมอย่างน้อย 1 คน" };
        }
        requiredLookups.forEach((users) => {
            if (users && users.length === 1) {
                attendeeObjects.push({
                    emailAddress: { address: users[0].userPrincipalName, name: users[0].displayName },
                    type: "required"
                });
            }
        });
        if (optionalAttendees.length > 0) {
            const optionalLookups = await Promise.all(
                optionalAttendees.map(name => findUserByShortName(name.trim()))
            );
            optionalLookups.forEach((users) => {
                if (users && users.length === 1) {
                    attendeeObjects.push({
                        emailAddress: { address: users[0].userPrincipalName, name: users[0].displayName },
                        type: "optional"
                    });
                }
            });
        }
        const event = {
            subject: subject,
            body: { contentType: "HTML", content: bodyContent || "" },
            start: { dateTime: startDateTime, timeZone: "Asia/Bangkok" },
            end: { dateTime: endDateTime, timeZone: "Asia/Bangkok" },
            location: { displayName: location || "" },
            attendees: attendeeObjects,
        };
        if (recurrence) {
            event.recurrence = recurrence;
        }
        if (createMeeting) {
            event.isOnlineMeeting = true;
            event.onlineMeetingProvider = "teamsForBusiness";
        }
        const url = `https://graph.microsoft.com/v1.0/users/${organizerEmail}/events`;
        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify(event)
        });
        if (!response.ok) {
            const errorText = await response.text();
            console.error('Graph API create event error:', errorText);
            return { error: `HTTP ${response.status}: ${errorText}` };
        }
        const createdEvent = await response.json();
        return {
            success: true,
            subject: createdEvent.subject,
            startTime: createdEvent.start.dateTime,
            organizer: createdEvent.organizer.emailAddress.name,
            webLink: createdEvent.webLink,
            meetingCreated: createdEvent.isOnlineMeeting || false
        };
    } catch (error) {
        console.error('createCalendarEvent error:', error);
        return { error: error.message };
    }
}
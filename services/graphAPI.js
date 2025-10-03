
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
        const url = `https://graph.microsoft.com/v1.0/users/${userEmail}/calendarView?startDateTime=${startDateTime}&endDateTime=${endDateTime}&$select=subject,body,bodyPreview,organizer,attendees,start,end,location,onlineMeeting`;
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
        return await response.json();
    } catch (error) {
        console.error('Graph API error:', error);
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
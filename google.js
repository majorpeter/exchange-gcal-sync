const fs = require('fs').promises;
const path = require('path');
const process = require('process');
const { authenticate } = require('@google-cloud/local-auth');
const { google } = require('googleapis');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/calendar'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');

/**
 * based on Google's Node.js example code
 * @see https://developers.google.com/calendar/api/quickstart/nodejs
 */
class GoogleCalendar {
    #calendarId;
    #auth;

    constructor(calendarId) {
        this.#calendarId = calendarId;
        this.#auth = this.#authorize();
    }

    /**
     * get all calendar events between start and end datetimes
     * @param {Date} startDateTime start date and time for filtering
     * @param {Date} endDateTime end date and time for filtering
     * @note max 1000 entries are returned
     */
    async getCalendarEvents(startDateTime, endDateTime) {
        const calendar = google.calendar({ version: 'v3', auth: await this.#auth });
        const res = await calendar.events.list({
            calendarId: this.#calendarId,
            timeMin: startDateTime.toISOString(),
            timeMax: endDateTime.toISOString(),
            maxResults: 1000,   // max would be 2500
            singleEvents: true,
            orderBy: 'startTime',
        });
        return res.data.items;
    }

    /**
     * remove an entry from calendar
     * @param {string} id 'id' field in calendar object
     */
    async deleteCalendarEvent(id) {
        const calendar = google.calendar({ version: 'v3', auth: await this.#auth });
        await calendar.events.delete({
            calendarId: this.#calendarId,
            eventId: id
        });
    }

    /**
     * Lists the next 10 events on the user's primary calendar.
     */
    async listEvents() {
        const calendar = google.calendar({ version: 'v3', auth: await this.#auth });
        const res = await calendar.events.list({
            calendarId: this.#calendarId,
            timeMin: (new Date()).toISOString(),
            maxResults: 10,
            singleEvents: true,
            orderBy: 'startTime',
        });
        const events = res.data.items;
        if (!events || events.length === 0) {
            console.log('No upcoming events found.');
            return;
        }
        console.log('Upcoming 10 events:');
        events.map((event, i) => {
            const start = event.start.dateTime || event.start.date;
            console.log(`${start} - ${event.summary}`);
        });
    }

    /**
     * create new event
     * @param {calendar_v3.Schema$Event} event event details
     */
    async insertEvent(eventBody) {
        const calendar = google.calendar({ version: 'v3', auth: await this.#auth });
        const result = await calendar.events.insert({
            calendarId: this.#calendarId,
            requestBody: eventBody
        });
    }

    /**
     * Reads previously authorized credentials from the save file.
     *
     * @return {Promise<OAuth2Client|null>}
     */
    static async loadSavedCredentialsIfExist() {
        try {
            const content = await fs.readFile(TOKEN_PATH);
            const credentials = JSON.parse(content);
            return google.auth.fromJSON(credentials);
        } catch (err) {
            return null;
        }
    }

    /**
     * Serializes credentials to a file comptible with GoogleAUth.fromJSON.
     *
     * @param {OAuth2Client} client
     * @return {Promise<void>}
     */
    static async saveCredentials(client) {
        const content = await fs.readFile(CREDENTIALS_PATH);
        const keys = JSON.parse(content);
        const key = keys.installed || keys.web;
        const payload = JSON.stringify({
            type: 'authorized_user',
            client_id: key.client_id,
            client_secret: key.client_secret,
            refresh_token: client.credentials.refresh_token,
        });
        await fs.writeFile(TOKEN_PATH, payload);
    }

    /**
     * Load or request or authorization to call APIs.
     *
     */
    async #authorize() {
        let client = await GoogleCalendar.loadSavedCredentialsIfExist();
        if (client) {
            return client;
        }

        client = await authenticate({
            scopes: SCOPES,
            keyfilePath: CREDENTIALS_PATH,
        });
        if (client.credentials) {
            await GoogleCalendar.saveCredentials(client);
        }
        return client;
    }
}

module.exports = GoogleCalendar;

"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __classPrivateFieldSet = (this && this.__classPrivateFieldSet) || function (receiver, state, value, kind, f) {
    if (kind === "m") throw new TypeError("Private method is not writable");
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a setter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot write private member to an object whose class did not declare it");
    return (kind === "a" ? f.call(receiver, value) : f ? f.value = value : state.set(receiver, value)), value;
};
var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _GoogleCalendar_instances, _GoogleCalendar_calendarId, _GoogleCalendar_auth, _GoogleCalendar_authorize;
Object.defineProperty(exports, "__esModule", { value: true });
exports.GoogleCalendar = void 0;
const googleapis_1 = require("googleapis");
const local_auth_1 = require("@google-cloud/local-auth");
const fs = require('fs').promises;
const path = require('path');
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
    constructor(calendarId) {
        _GoogleCalendar_instances.add(this);
        _GoogleCalendar_calendarId.set(this, void 0);
        _GoogleCalendar_auth.set(this, void 0);
        __classPrivateFieldSet(this, _GoogleCalendar_calendarId, calendarId, "f");
        __classPrivateFieldSet(this, _GoogleCalendar_auth, __classPrivateFieldGet(this, _GoogleCalendar_instances, "m", _GoogleCalendar_authorize).call(this), "f");
    }
    /**
     * get all calendar events between start and end datetimes
     * @param {Date} startDateTime start date and time for filtering
     * @param {Date} endDateTime end date and time for filtering
     * @note max 1000 entries are returned
     */
    getCalendarEvents(startDateTime, endDateTime) {
        return __awaiter(this, void 0, void 0, function* () {
            const calendar = googleapis_1.google.calendar({ version: 'v3', auth: yield __classPrivateFieldGet(this, _GoogleCalendar_auth, "f") });
            const res = yield calendar.events.list({
                calendarId: __classPrivateFieldGet(this, _GoogleCalendar_calendarId, "f"),
                timeMin: startDateTime.toISOString(),
                timeMax: endDateTime.toISOString(),
                maxResults: 1000,
                singleEvents: true,
                orderBy: 'startTime',
            });
            return res.data.items;
        });
    }
    /**
     * remove an entry from calendar
     * @param {string} id 'id' field in calendar object
     */
    deleteCalendarEvent(id) {
        return __awaiter(this, void 0, void 0, function* () {
            const calendar = googleapis_1.google.calendar({ version: 'v3', auth: yield __classPrivateFieldGet(this, _GoogleCalendar_auth, "f") });
            yield calendar.events.delete({
                calendarId: __classPrivateFieldGet(this, _GoogleCalendar_calendarId, "f"),
                eventId: id
            });
        });
    }
    /**
     * Lists the next 10 events on the user's primary calendar.
     */
    listEvents() {
        return __awaiter(this, void 0, void 0, function* () {
            const calendar = googleapis_1.google.calendar({ version: 'v3', auth: yield __classPrivateFieldGet(this, _GoogleCalendar_auth, "f") });
            const res = yield calendar.events.list({
                calendarId: __classPrivateFieldGet(this, _GoogleCalendar_calendarId, "f"),
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
        });
    }
    /**
     * create new event
     * @param {calendar_v3.Schema$Event} eventBody event details
     */
    insertEvent(eventBody) {
        return __awaiter(this, void 0, void 0, function* () {
            const calendar = googleapis_1.google.calendar({ version: 'v3', auth: yield __classPrivateFieldGet(this, _GoogleCalendar_auth, "f") });
            const result = yield calendar.events.insert({
                calendarId: __classPrivateFieldGet(this, _GoogleCalendar_calendarId, "f"),
                requestBody: eventBody
            });
        });
    }
    /**
     * update already existing event
     * @param {string} eventId 'id' field in calendar object
     * @param {calendar_v3.Schema$Event} eventBody event details
     */
    patchEvent(eventId, eventBody) {
        return __awaiter(this, void 0, void 0, function* () {
            const calendar = googleapis_1.google.calendar({ version: 'v3', auth: yield __classPrivateFieldGet(this, _GoogleCalendar_auth, "f") });
            const result = yield calendar.events.patch({
                calendarId: __classPrivateFieldGet(this, _GoogleCalendar_calendarId, "f"),
                eventId: eventId,
                requestBody: eventBody
            });
        });
    }
    /**
     * Reads previously authorized credentials from the save file.
     */
    static loadSavedCredentialsIfExist() {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const content = yield fs.readFile(TOKEN_PATH);
                const credentials = JSON.parse(content);
                return googleapis_1.google.auth.fromJSON(credentials);
            }
            catch (err) {
                return null;
            }
        });
    }
    /**
     * Serializes credentials to a file comptible with GoogleAUth.fromJSON.
     * @param {OAuth2Client} client
     * @return {Promise<void>}
     */
    static saveCredentials(client) {
        return __awaiter(this, void 0, void 0, function* () {
            const content = yield fs.readFile(CREDENTIALS_PATH);
            const keys = JSON.parse(content);
            const key = keys.installed || keys.web;
            const payload = JSON.stringify({
                type: 'authorized_user',
                client_id: key.client_id,
                client_secret: key.client_secret,
                refresh_token: client.credentials.refresh_token,
            });
            yield fs.writeFile(TOKEN_PATH, payload);
        });
    }
}
exports.GoogleCalendar = GoogleCalendar;
_GoogleCalendar_calendarId = new WeakMap(), _GoogleCalendar_auth = new WeakMap(), _GoogleCalendar_instances = new WeakSet(), _GoogleCalendar_authorize = function _GoogleCalendar_authorize() {
    return __awaiter(this, void 0, void 0, function* () {
        let savedClient = yield GoogleCalendar.loadSavedCredentialsIfExist();
        if (savedClient) {
            return savedClient;
        }
        let authenticatedClient = yield (0, local_auth_1.authenticate)({
            scopes: SCOPES,
            keyfilePath: CREDENTIALS_PATH,
        });
        if (authenticatedClient.credentials) {
            yield GoogleCalendar.saveCredentials(authenticatedClient);
        }
        return authenticatedClient;
    });
};
//# sourceMappingURL=google.js.map
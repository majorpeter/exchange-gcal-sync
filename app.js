const fs = require('fs');
const ExchangeCalendar = require('./exchange');
const GoogleCalendar = require('./google');
const util = require('./util');

let config = JSON.parse(fs.readFileSync('config.json'));
let exch = new ExchangeCalendar(config.exchangeServerUrl, config.exchangeDomain, config.exchangeUsername, config.exchangePassword);
let gcal = new GoogleCalendar(config.googleCalendarId);

(async () => {
    let gEvents = await gcal.getCalendarEvents(util.startOfWeek(), util.endOfWeek());
    let exchEvents = await exch.getCalendarEvents(util.startOfWeek(), util.endOfWeek());

    let meetingtime = 0;
    exchEvents.value.forEach((event) => {
        let data = {
            subject: event.Subject,
            start: new Date(event.Start.DateTime + 'Z'),
            end: new Date(event.End.DateTime + 'Z'),
            attendees: event.Attendees.map(value => value.EmailAddress.Address),
            iCalUId: event.iCalUId,
            lastModified: event.LastModifiedDateTime,
        };

        let foundGEvent = gEvents.find(value => value.extendedProperties.private.sourceICalUId == event.iCalUId);
        if (foundGEvent) {
            foundGEvent.found = true;
            // TODO check for change and update
        } else {
            const gcalEventBody = {
                summary: data.subject,
                start: {dateTime: data.start.toISOString()},
                end: {dateTime: data.end.toISOString()},
                // cannot use 'iCalUID' for sync
                extendedProperties: {private: {
                    sourceICalUId: data.iCalUId,
                    sourceLastModified: data.lastModified
                }}
            };
            console.log('inserting event: %s', data.subject);
            //TODO gcal.insertEvent(gcalEventBody);
        }

        if (data.attendees.length > 0) {
            let length = (data.end - data.start) / 3600e3;
            // console.log('Meeting: ' + data.subject + ' (' + length.toFixed(1) + ' h)');
            meetingtime += length;
        }
    });

    // remove google calendar entries that are no longer in exchange
    gEvents.forEach((event) => {
        if (!event.found) {
            console.log('removing: %d', event.summary);
            gcal.deleteCalendarEvent(event.id);
        }
    });

    console.log('All meeting time for this week: ' + meetingtime.toFixed(2) + ' hrs');
    console.log('Meeting time percentage: ' + (meetingtime / 40 * 100).toFixed(2) + '%');
})();

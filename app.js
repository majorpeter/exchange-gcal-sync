const fs = require('fs');
const ExchangeCalendar = require('./exchange');
const GoogleCalendar = require('./google');

let config = JSON.parse(fs.readFileSync('config.json'));
let exch = new ExchangeCalendar(config.exchangeServerUrl, config.exchangeDomain, config.exchangeUsername, config.exchangePassword);
let gcal = new GoogleCalendar(config.googleCalendarId);

(async () => {
    let events = await exch.getCalendarEventsThisWeek();
    let meetingtime = 0;

    events.value.forEach((event) => {
        let data = {
            subject: event.Subject,
            start: new Date(event.Start.DateTime + 'Z'),
            end: new Date(event.End.DateTime + 'Z'),
            attendees: event.Attendees.map(value => value.EmailAddress.Address)
        };

        if (data.attendees.length > 0) {
            let length = (data.end - data.start) / 3600e3;
            console.log('Meeting: ' + data.subject + ' (' + length.toFixed(1) + ' h)');
            meetingtime += length;
        }
    });
    console.log();

    console.log('All meeting time for this week: ' + meetingtime.toFixed(2) + ' hrs');
    console.log('Meeting time percentage: ' + (meetingtime / 40).toFixed(2) + '%');

    gcal.listEvents();
})();

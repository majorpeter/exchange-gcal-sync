const fs = require('fs');
const ExchangeCalendar = require('./exchange');
const GoogleCalendar = require('./google');
const util = require('./util');

let config = JSON.parse(fs.readFileSync('config.json'));
let exch = new ExchangeCalendar(config.exchangeServerUrl, config.exchangeDomain, config.exchangeUsername, config.exchangePassword);
let gcal = new GoogleCalendar(config.googleCalendarId);

function convertExchangeResponseToGCal(response) {
    if (['Organizer', 'Accepted'].includes(response)) {
        return 'confirmed';
    }
    if (['TentativelyAccepted', 'NotResponded'].includes(response)) {
        return 'tentative';
    }
    throw Error('Unknown response status: ' + response);
}

async function syncEvents(dryRun, forceUpdate) {
    let gEvents = await gcal.getCalendarEvents(util.startOfWeek(), util.endOfNextWeek());
    let exchEvents = await exch.getCalendarEvents(util.startOfWeek(), util.endOfNextWeek());

    let meetingtime = 0;
    exchEvents.value.forEach((event) => {
        let data = {
            subject: event.Subject,
            body: event.Body.Content,
            location: event.Location.DisplayName,
            start: new Date(event.Start.DateTime + 'Z'),
            end: new Date(event.End.DateTime + 'Z'),
            attendees: event.Attendees.map(value => value.EmailAddress.Address),
            iCalUId: event.iCalUId,
            lastModified: event.LastModifiedDateTime,
            status: convertExchangeResponseToGCal(event.ResponseStatus.Response),
        };

        let foundGEvent = gEvents.find(value => value.extendedProperties.private.sourceICalUId == event.iCalUId);
        if (foundGEvent) {
            foundGEvent.found = true;
            // TODO check for change and update
        } else {
            const gcalEventBody = {
                summary: data.subject,
                description: data.body,
                location: data.location,
                start: {dateTime: data.start.toISOString()},
                end: {dateTime: data.end.toISOString()},
                status: data.status,
                // cannot use 'iCalUID' for sync
                extendedProperties: {private: {
                    sourceICalUId: data.iCalUId,
                    sourceLastModified: data.lastModified
                }}
            };
            console.log('inserting event: %s', data.subject);
            if (!dryRun) {
                gcal.insertEvent(gcalEventBody);
            }
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
            if (!dryRun) {
                gcal.deleteCalendarEvent(event.id);
            }
        }
    });

    console.log('All meeting time for this week: ' + meetingtime.toFixed(2) + ' hrs');
    console.log('Meeting time percentage: ' + (meetingtime / 40 * 100).toFixed(2) + '%');
}

const args = require('yargs')
  .scriptName('exchange-gcal-sync')
  .option('force', {
      type: 'boolean',
      alias: 'f',
      default: false,
      desc: 'Force update of entries even if unchanged'
  })
  .option('dry-run', {
      type: 'boolean',
      alias: 'n',
      default: false,
      desc: 'Do not change Google Calendar entries, just log instead.'
  })
  .help()
  .argv;
syncEvents(args.dryRun, args.force);

const fs = require('fs');
const GoogleCalendar = require('./google');
const util = require('./util');

import {Exchange} from './exchange';

const config : {
    exchangeServerUrl: string;
    exchangeDomain: string;
    exchangeUsername: string;
    exchangePassword: string;
    googleCalendarId: string;
} = JSON.parse(fs.readFileSync('config.json'));
let exch = new Exchange.Calendar(config.exchangeServerUrl, config.exchangeDomain, config.exchangeUsername, config.exchangePassword);
let gcal = new GoogleCalendar(config.googleCalendarId);

function convertExchangeResponseToGCal(response: Exchange.ResponseStatus): 'confirmed'|'tentative' {
    switch (response) {
    case "Organizer":
    case "Accepted":
        return 'confirmed';
    case "TentativelyAccepted":
    case "NotResponded":
        return 'tentative';
    default:
        throw Error('Unknown response status: ' + response);
    }
}

async function syncEvents(dryRun: boolean, forceUpdate: boolean) {
    let gEvents = await gcal.getCalendarEvents(util.startOfWeek(), util.endOfNextWeek());
    console.log('Found %d events in Google Calendar', gEvents.length);
    let exchEvents = await exch.getCalendarEvents(util.startOfWeek(), util.endOfNextWeek());
    console.log('Found %d events in Exchange Calendar', exchEvents.value.length);

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

        let foundGEvent = gEvents.find(value => value.extendedProperties.private.sourceICalUId == event.iCalUId);

        if (foundGEvent) {
            // flags 'not to delete' later
            foundGEvent.found = true;

            if (event.LastModifiedDateTime == foundGEvent.extendedProperties.private.sourceLastModified) {
                if (!forceUpdate) {
                    return;
                }
            }

            console.log('patching event: %s', data.subject);
            if (!dryRun) {
                gcal.patchEvent(foundGEvent.id, gcalEventBody);
            }
        } else {
            console.log('inserting event: %s', data.subject);
            if (!dryRun) {
                gcal.insertEvent(gcalEventBody);
            }
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

    console.log('Sync done');
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
  .option('period', {
      type: 'number',
      alias: 'p',
      desc: 'Update period in minutes',
      nargs: 1,
      default: null
  })
  .help()
  .strict()
  .argv;

syncEvents(args.dryRun, args.force);

if (args.period) {
    setInterval(() => {
        console.log('Periodic run after %d min', args.period);
        syncEvents(args.dryRun, args.force);
    }, args.period * 1000 * 60);
}

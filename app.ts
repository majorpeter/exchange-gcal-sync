const util = require('./util');

import { readFileSync } from 'fs';
import { calendar_v3 } from 'googleapis';
import path = require('path');
import { Exchange } from './exchange';
import { GoogleCalendar } from './google';

const config : {
    exchangeServerUrl: string;
    exchangeDomain: string;
    exchangeUsername: string;
    exchangePassword: string;
    googleCalendarId: string;
} = JSON.parse(readFileSync(path.join(__dirname, 'config.json')).toString());
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

function formatCalendarEventBody(event: Exchange.Event): string {
    let result = '';
    if (event.Attendees.length > 0) {
        result += '<strong>Attendees</strong>:\n';
        result += event.Attendees.map(attendee => {
            return attendee.EmailAddress.Name + ' &lt;' + attendee.EmailAddress.Address + '&gt;';
        }).join('; ') + '\n';
    }
    result += event.Body.Content;
    result += '<a href="' + event.WebLink + '">View in Exchange</a>';

    return result;
}

async function syncEvents(dryRun: boolean, forceUpdate: boolean) {
    interface gEvent extends calendar_v3.Schema$Event {
        found?: boolean;
        extendedProperties?: {
            private?: {
                /// used to uniquely identify the Event in Exchange
                sourceICalUId?: string;
                /// used to detectt changes in Exchange since last sync
                sourceLastModified?: string;
            };
        };
    };

    let gEvents: gEvent[] = await gcal.getCalendarEvents(util.startOfWeek(), util.endOfNextWeek());
    console.log('Found %d events in Google Calendar', gEvents.length);
    let exchEvents = await exch.getCalendarEvents(util.startOfWeek(), util.endOfNextWeek());
    console.log('Found %d events in Exchange Calendar', exchEvents.value.length);

    exchEvents.value.forEach((event: Exchange.Event) => {
        const gcalEventBody: gEvent = {
            summary: event.Subject,
            description: formatCalendarEventBody(event),
            location: event.Location.DisplayName,
            start: {dateTime: (new Date(event.Start.DateTime + 'Z')).toISOString()},
            end: {dateTime: (new Date(event.End.DateTime + 'Z')).toISOString()},
            status: convertExchangeResponseToGCal(event.ResponseStatus.Response),
            // cannot use 'iCalUID' for sync
            extendedProperties: {private: {
                sourceICalUId: event.iCalUId,
                sourceLastModified: event.LastModifiedDateTime
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

            console.log('patching event: %s', event.Subject);
            if (!dryRun) {
                gcal.patchEvent(foundGEvent.id, gcalEventBody);
            }
        } else {
            console.log('inserting event: %s', event.Subject);
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

if (args.dryRun) {
    console.log('Dry-run: Not writing to Google Calendar');
}
syncEvents(args.dryRun, args.force);

if (args.period) {
    setInterval(() => {
        console.log('Periodic run after %d min', args.period);
        syncEvents(args.dryRun, args.force);
    }, args.period * 1000 * 60);
}

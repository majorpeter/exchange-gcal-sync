import { readFileSync, writeFileSync } from 'fs';
import { calendar_v3 } from 'googleapis';
import path = require('path');
import { Exchange } from './exchange';
import { GoogleCalendar } from './google';
import { Util } from './util';

const config: {
    exchangeServerUrl: string;
    exchangeDomain: string;
    exchangeUsername: string;
    exchangePassword: string;
    googleCalendarId: string;
} = JSON.parse(readFileSync(path.join(__dirname, 'config.json')).toString());
let exch = new Exchange.Calendar(config.exchangeServerUrl, config.exchangeDomain, config.exchangeUsername, config.exchangePassword);
let gcal = new GoogleCalendar(config.googleCalendarId);

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

function convertExchangeResponseToGCal(response: Exchange.ResponseStatus): 'confirmed' | 'tentative' {
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

    if (event.WebLink) {
        result += '<a href="' + event.WebLink + '">View in Exchange</a>';
    }

    return result;
}

/**
 * creates a bogus event for syncing that contains statistics about the week
 * @param events {Exchange.Event[]} all events found in query
 * @returns {Exchange.Event} bogus event
 */
function generateStatisticsEvent(events: readonly Exchange.Event[]): Exchange.Event {
    const minimumStartDateTime = events.map(event => event.Start.DateTime).reduce((prev, curr) => (prev < curr) ? prev : curr);
    const startOfWeek: Date = Util.startOfWeek(new Date(minimumStartDateTime + 'Z'));

    const lastModifiedDateTime = events.map(event => event.LastModifiedDateTime).reduce((prev, curr) => (prev > curr) ? prev : curr);
    const endOfWeekDateTime = new Date(startOfWeek.valueOf() + 7 * 24 * 3600e3).toISOString();
    const weekMeetingsSorted = events.filter(event => (event.End.DateTime < endOfWeekDateTime) && event.Attendees.length > 0)
                                .sort((a, b) => (a.Start.DateTime < b.Start.DateTime) ? -1 : 1);

    let meetingHours = 0;
    let overLapCount = 0;
    weekMeetingsSorted.forEach((event, index, array) => {
        const start = new Date(event.Start.DateTime + 'Z');
        const end = new Date(event.End.DateTime + 'Z');
        meetingHours += (+end - +start) / 3600e3;

        for (let i = 0; i < index; i++) {
            if (array[i].End.DateTime > event.Start.DateTime) {
                overLapCount++;

                const ostart = new Date(array[i].Start.DateTime + 'Z');
                const oend = new Date(array[i].End.DateTime + 'Z');

                meetingHours -= (Math.min(+oend, +end) - +start) / 3600e3;
            }
        }
    });
    const ratio = (meetingHours / 40 * 100).toFixed(2) + '%';

    return {
        Subject: `Exchange Calendar Statistics (${ratio})`,
        Body: {
            ContentType: 'HTML', Content:
                `Meetings this week: <b>${weekMeetingsSorted.length}</b>\n` +
                `Meeting hours this week: <b>${meetingHours.toFixed(1)} h (${ratio})</b>\n` +
                `Overlapping meetings this week: <b>${overLapCount}</b>`
        },
        Attendees: [],
        Start: {
            DateTime: new Date(startOfWeek.valueOf() + 8 * 3600e3).toISOString().slice(0, -1),
            TimeZone: 'UTC',
        },
        End: {
            DateTime: new Date(startOfWeek.valueOf() + 8 * 3600e3 + 600e3).toISOString().slice(0, -1),
            TimeZone: 'UTC',
        },
        LastModifiedDateTime: lastModifiedDateTime,
        Location: { DisplayName: '' },
        ResponseStatus: { Response: 'Organizer' },
        iCalUId: 'Statistics',
    };
}

async function syncEvents(dryRun: boolean, forceUpdate: boolean, stats: boolean, save: boolean, load: boolean): Promise<boolean> {
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

    let gEvents: gEvent[];
    let exchEvents: Exchange.EventList;

    if (!load) {
        try {
            gEvents = await gcal.getCalendarEvents(Util.startOfWeek(), Util.endOfNextWeek());
            console.log('Found %d events in Google Calendar', gEvents.length);
        } catch (e) {
            console.log('Google Calendar fetch failed:');
            console.log(e);
            return false;
        }

        try {
            exchEvents = await exch.getCalendarEvents(Util.startOfWeek(), Util.endOfNextWeek());
            console.log('Found %d events in Exchange Calendar', exchEvents.value.length);
        } catch (e) {
            console.log('Exchange Calendar fetch failed:');
            console.log(e);
            return false;
        }
    } else {
        gEvents = JSON.parse(readFileSync('gEvents.json').toString());
        exchEvents = JSON.parse(readFileSync('exchEvents.json').toString());
    }

    if (save) {
        writeFileSync('gEvents.json', JSON.stringify(gEvents));
        writeFileSync('exchEvents.json', JSON.stringify(exchEvents));
    }

    if (stats && (exchEvents.value.length > 0)) {
        exchEvents.value.push(generateStatisticsEvent(exchEvents.value));
    }

    for (const event of exchEvents.value) {
        const gcalEventBody: gEvent = {
            summary: event.Subject,
            description: formatCalendarEventBody(event),
            location: event.Location.DisplayName,
            start: { dateTime: (new Date(event.Start.DateTime + 'Z')).toISOString() },
            end: { dateTime: (new Date(event.End.DateTime + 'Z')).toISOString() },
            status: convertExchangeResponseToGCal(event.ResponseStatus.Response),
            // cannot use 'iCalUID' for sync
            extendedProperties: {
                private: {
                    sourceICalUId: event.iCalUId,
                    sourceLastModified: event.LastModifiedDateTime
                }
            }
        };

        let foundGEvent = gEvents.find(value => {
            try {
                return value.extendedProperties.private.sourceICalUId == event.iCalUId;
            } catch (e) {
                // TypeError if keys not present in event object
                if (e instanceof TypeError) {
                    return false;
                }
            }
        });

        if (foundGEvent) {
            // flags 'not to delete' later
            foundGEvent.found = true;

            if (event.LastModifiedDateTime == foundGEvent.extendedProperties.private.sourceLastModified) {
                if (!forceUpdate) {
                    continue;
                }
            }

            console.log('patching event: %s', event.Subject);
            if (!dryRun) {
                await gcal.patchEvent(foundGEvent.id, gcalEventBody);
                await sleep(1000);  // rate limiting
            }
        } else {
            console.log('inserting event: %s', event.Subject);
            if (!dryRun) {
                await gcal.insertEvent(gcalEventBody);
                await sleep(1000);  // rate limiting
            }
        }
    }

    // remove google calendar entries that are no longer in exchange
    for (const event of gEvents) {
        if (!event.found) {
            console.log('removing: %s', event.summary);
            if (!dryRun) {
                await gcal.deleteCalendarEvent(event.id);
                await sleep(1000);  // rate limiting
            }
        }
    }

    console.log('Sync done');
    return true;
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
    .option('stats', {
        type: 'boolean',
        default: false,
        desc: 'Include a calendar entry with statistics',
    })
    .option('save', {
        type: 'boolean',
        default: false,
        desc: 'Save remote data to JSON files (useful for debugging)'
    })
    .option('load', {
        type: 'boolean',
        default: false,
        desc: 'Load "remote" data from JSON files (when debugging)'
    })
    .help()
    .strict()
    .argv;

if (args.save && args.load) {
    throw Error('Cannot load and save at the same time.');
}
if (args.dryRun) {
    console.log('Dry-run: Not writing to Google Calendar');
}

syncEvents(args.dryRun, args.force, args.stats, args.save, args.load);

if (args.period) {
    setInterval(() => {
        console.log('Periodic run after %d min', args.period);
        syncEvents(args.dryRun, args.force, args.stats, false, false);
    }, args.period * 1000 * 60);
}

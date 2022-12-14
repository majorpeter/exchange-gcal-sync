const httpntlm = require('httpntlm');

export namespace Exchange {

interface DateTime {
    DateTime: string;
    TimeZone: string;
};

interface Email {
    Name: string;
    Address: string;
};

interface Attendee {
    Type: "Required" | string;
    Status: {
        Response: string;
        Time: string;
    };
    EmailAddress: Email;
};

export type ResponseStatus = "Organizer" | "Accepted" | "TentativelyAccepted" | "NotResponded";

export interface Event {
    "@odata.etag"?: string;
    "@odata.id"?: string;
    Attendees: Attendee[];
    Body: {
        ContentType: "HTML" | string;
        Content: string;
    };
    BodyPreview?: string;
    End: DateTime;
    iCalUId: string;
    IsAllDay: boolean;
    IsReminderOn?: boolean;
    LastModifiedDateTime: string;
    Location: {
        DisplayName: string;
        Address?: {
            Type: string;
        };
        Coordinates?: any;
    };
    Organizer?: {
        EmailAddress: Email;
    };
    ReminderMinutesBeforeStart?: number;
    ResponseStatus: {
        Response: ResponseStatus;
        Time?: string;
    };
    Sensitivity?: "Normal" | "Private";
    ShowAs?: "Busy" | "Tentative" | "Oof";
    Start: DateTime;
    Subject: string;
    WebLink?: string;
};

export interface EventList {
    "@odata.context": string;
    value: Event[];
}

export class Calendar {
    serverurl: string;
    domain: string;
    username: string;
    password: string;

    constructor(serverurl: string, domain: string, username: string, password: string) {
        this.serverurl = serverurl;
        this.domain = domain;
        this.username = username;
        this.password = password;
    }

    static dateToStr(date: Date): string {
        return new Date(date.valueOf() - date.getTimezoneOffset()*60000).toISOString()
    }

    /**
     * get all calendar events between start and end datetimes
     * @param {Date} startDateTime start date and time for filtering
     * @param {Date} endDateTime end date and time for filtering
     * @returns {Promise<EventList>} calendar entries in Exchange's native format
     * @note max 1000 entries are returned
     * @see https://docs.microsoft.com/hu-hu/archive/blogs/exchangedev/throttling-coming-to-outlook-api-and-microsoft-graph
     */
    async getCalendarEvents(startDateTime: Date, endDateTime: Date): Promise<EventList> {
        return new Promise((resolve, reject) => {
            httpntlm.get({
                url: this.serverurl + 'api/v2.0/users/me/calendarview?StartDateTime=' + Calendar.dateToStr(startDateTime) + '&EndDateTime=' + Calendar.dateToStr(endDateTime) + '&$top=1000',
                username: this.username,
                password: this.password,
                domain: this.domain
            }, function (err, res) {
                if (err) {
                    reject(err);
                    return;
                }
                if (res.statusCode != 200) {
                    reject(res);
                    return;
                }
                resolve(JSON.parse(res.body));
            });
        });
    }
}

}

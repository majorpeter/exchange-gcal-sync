import { Exchange } from "./exchange";
import { GoogleCalendarColor } from "./google";

export namespace Util {
    export function startOfWeek(date?: Date): Date {
        if (typeof(date) == 'undefined') {
            date = new Date();
        }

        date.setHours(0, 0, 0, 0);
        let day = date.getDay(), diff = date.getDate() - day + (day == 0 ? -6:1); // adjust when day is sunday
        date.setDate(diff);
        return date;
    }

    export function daysFromDate(days: number, date: Date): Date {
        let d = new Date(date);
        d.setDate(d.getDate() + days);
        return d;
    }

    export function endOfWeek(): Date {
        let date = this.startOfWeek();
        date.setDate(date.getDate() + 7);
        return date;
    }

    export function endOfNextWeek(): Date {
        let date = this.startOfWeek();
        date.setDate(date.getDate() + 14);
        return date;
    }

    export function startOfDay(): Date {
        let date = new Date();
        date.setHours(0, 0, 0, 0);
        return date;
    }

    export function endOfDay(): Date {
        let date = new Date();
        date.setHours(23, 59, 59, 0);
        return date;
    }

    export function convertExchResponseToGCal(response: Exchange.ResponseStatus): 'confirmed' | 'tentative' {
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

    export function convertExchShowAsToGCalColor(showAs: string): GoogleCalendarColor | null {
        switch (showAs) {
        case 'Oof':
            return GoogleCalendarColor.Purple;
        case 'Tentative':
            return GoogleCalendarColor.Turquoise;
        }
        return null;
    }

    export function convertExchUtcDateTimeToGCal(dateTime: string, isAllDay: boolean): string {
        if (!isAllDay) {
            return (new Date(dateTime + 'Z')).toISOString();
        } else {
            return dateTime; //.split('T')[0];
        }
    }
}

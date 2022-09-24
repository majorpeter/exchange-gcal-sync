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
}

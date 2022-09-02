var httpntlm = require('httpntlm');

class ExchangeCalendar {
    constructor(serverurl, domain, username, password) {
        this.serverurl = serverurl;
        this.domain = domain;
        this.username = username;
        this.password = password;
    }

    static dateToStr(date) {
        return new Date(date.valueOf() - date.getTimezoneOffset()*60000).toISOString()
    }

    /**
     * get all calendar events between start and end datetimes
     * @param {Date} startDateTime start date and time for filtering
     * @param {Date} endDateTime end date and time for filtering
     * @returns calendar entries in Exchange's native format
     * @note max 1000 entries are returned
     * @see https://docs.microsoft.com/hu-hu/archive/blogs/exchangedev/throttling-coming-to-outlook-api-and-microsoft-graph
     */
    async getCalendarEvents(startDateTime, endDateTime) {
        return new Promise((resolve, reject) => {
            httpntlm.get({
                url: this.serverurl + 'api/v2.0/users/me/calendarview?StartDateTime=' + ExchangeCalendar.dateToStr(startDateTime) + '&EndDateTime=' + ExchangeCalendar.dateToStr(endDateTime) + '&$top=1000',
                username: this.username,
                password: this.password,
                domain: this.domain
            }, function (err, res) {
                if (res.statusCode != 200) {
                    reject(res);
                }
                resolve(JSON.parse(res.body));
            });
        });
    }
}

module.exports = ExchangeCalendar;

module.exports = {
    startOfWeek: function() {
        let date = new Date();
        date.setHours(0, 0, 0, 0);
        let day = date.getDay(), diff = date.getDate() - day + (day == 0 ? -6:1); // adjust when day is sunday
        date.setDate(diff);
        return date;
    },

    endOfWeek: function() {
        let date = this.startOfWeek();
        date.setDate(date.getDate() + 7);
        return date;
    },

    startOfDay: function() {
        let date = new Date();
        date.setHours(0, 0, 0, 0);
        return date;
    },

    endOfDay: function() {
        let date = new Date();
        date.setHours(23, 59, 59, 0);
        return date;
    }
}

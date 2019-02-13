// 1) show public holidays when duration is given
// 2) show flexible holidays when duration is given
// 3) Avail flexible holidays (with validation of max 3 flexible holidays)





const { TimexProperty } = require('@microsoft/recognizers-text-data-types-timex-expression');
const { ActivityTypes, CardFactory, ActionTypes } = require('botbuilder')
const holidays = require('./resources/holidays.json');

const FLEXIBLE_HOLIDAY_TYPE = "flexible holidays";
const DATE_TIME = "datetime";
const REQUEST_TYPES = "request_types";
const PUBLIC_HOLIDAY = "Public";
const FLEXIBLE_HOLIDAY = "flexible"
class HolidayCalendar {

    async listHolidays(context, entities) {
        console.log(entities);

        if (entities[DATE_TIME]) {
            var timexDate = new TimexProperty(entities[DATE_TIME][0].timex.toString());
            if (entities[REQUEST_TYPES] && entities[REQUEST_TYPES][0].includes(FLEXIBLE_HOLIDAY_TYPE)) {
                var flexibleHolidays = holidays.filter(
                    function (calendarHoliday) {

                        // date to timex ==> TimexProperty.fromDate(calendarHoliday.date)..   only comparison left
                        return calendarHoliday.type === FLEXIBLE_HOLIDAY && (new Date() < new Date(calendarHoliday.date));
                    }
                );
                if (flexibleHolidays.length > 0) {
                    this.holidayCard = this.createHeroCard(flexibleHolidays);
                } else {
                    return context.sendActivity("No flexible holidays.");
                }
            } else {
                var publicHolidays = holidays.filter(
                    function (calendarHoliday) {
                        // date to timex ==> TimexProperty.fromDate(calendarHoliday.date)..   only comparison left
                        return calendarHoliday.type === PUBLIC_HOLIDAY && (new Date() < new Date(calendarHoliday.date));
                    }
                );

                if (publicHolidays.length > 0) {

                    this.holidayCard = this.createAdaptiveCard(publicHolidays);
                } else {
                    return context.sendActivity("No public holidays.");
                }
            }

        } else {
            if (entities[REQUEST_TYPES] && entities[REQUEST_TYPES][0].includes(FLEXIBLE_HOLIDAY_TYPE)) {
                var flexibleHolidays = holidays.filter(
                    function (calendarHoliday) {
                        return calendarHoliday.type === FLEXIBLE_HOLIDAY && (new Date() < new Date(calendarHoliday.date));
                    }
                );
                if (flexibleHolidays.length > 0) {
                    this.holidayCard = this.createHeroCard(flexibleHolidays);
                } else {
                    return context.sendActivity("No flexible holidays.");
                }
            } else {

                var publicHolidays = holidays.filter(
                    function (calendarHoliday) {
                        return calendarHoliday.type === PUBLIC_HOLIDAY && (new Date() < new Date(calendarHoliday.date));
                    }
                );

                if (publicHolidays.length > 0) {

                    this.holidayCard = this.createAdaptiveCard(publicHolidays);
                } else {
                    return context.sendActivity("No public holidays.");
                }
            }
        }

        const reply = {
            type: ActivityTypes.Message,
            attachments: [this.holidayCard]
        };
        return context.sendActivity(reply);
    }

    createHeroCard(flexibleHolidays) {
        let flexibleHolidayButtons = [];
        for (var i = 0; i < flexibleHolidays.length; i++) {
            const button = {
                type: ActionTypes.PostBack,
                title: flexibleHolidays[i].day + ", " + flexibleHolidays[i].date + " - " + flexibleHolidays[i].name,
                value: 'Button clicked'
            }
            flexibleHolidayButtons.push(button);
        }


        const flexibleHolidayHeroCard = CardFactory.heroCard(
            'Flexible Holidays',
            undefined,
            flexibleHolidayButtons,
            {
                text: 'You may avail for the flexible leave by selecting it.'
            }
        );

        return flexibleHolidayHeroCard;
    }


    createAdaptiveCard(publicHolidays) {
        var card = {
            "type": "AdaptiveCard",
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.0",
            "body": []
        };
        
        var textBlocks = [];
        for (var i = 0; i < publicHolidays.length; i++) {
            let textBlock = {
                "type": "TextBlock",
                "color": "dark",
                "wrap": true,
                "text": publicHolidays[i].day + ", " + publicHolidays[i].date + " - " + publicHolidays[i].name
            }
            textBlocks.push(textBlock);
        }
        card.body = textBlocks;

        return CardFactory.adaptiveCard(card);
    }
}

module.exports.HolidayCalendar = HolidayCalendar;
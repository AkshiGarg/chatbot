// 1) View leave records by duration


const { CardFactory, ActivityTypes } = require('botbuilder');


const fs = require('fs');
const USER_LEAVE_RECORDS_FILE = './resources/user_leave_record.json';
const Recognizers = require('@microsoft/recognizers-text-suite');


// entities defined in LUIS
const DATE_TIME = "datetime";
const REQUEST_TYPES = "request_types";

class LeaveRequestManager {

    constructor(userStateAccessor) {
        this.userStateAccessor = userStateAccessor;
    }

    async viewSubmittedRequests(userProfile, context, entities) {
        const records = fs.readFileSync(USER_LEAVE_RECORDS_FILE);
        const leaveRecords = JSON.parse(records);
        var userLeaveRecord = leaveRecords.find(leaveRecord => leaveRecord.employeeId === userProfile.id);

        if (!userLeaveRecord) {
            return await context.sendActivity("No record found for employee with id: " + userProfile.id);
        } else {
            var leaveRequests = userLeaveRecord.leaveRequests.filter(
                function (leaveRequest) {
                    if(entities[DATE_TIME]) {
                        LeaveRequestManager.filterByDate(entities[DATE_TIME], leaveRequest.date);
                    } else {
                        return entities[REQUEST_TYPES][0].includes(leaveRequest.type)
                            && (new Date() < new Date(leaveRequest.date));
                    }
                }
            );
            if (leaveRequests.length === 0) {
                return await context.sendActivity("No upcoming leave requests found for employee: " + userProfile.id);
            } else {

                this.leaveCard = this.createAdaptiveCard(leaveRequests);
            }
        }

        const reply = {
            type: ActivityTypes.Message,
            text: "You have submitted following requests: ",
            attachments: [this.leaveCard]
        };
        return context.sendActivity(reply);
    }

    createAdaptiveCard(leaveRequests) {
        var card = {
            "type": "AdaptiveCard",
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.0",
            "body": []
        };

        var textBlocks = [];
        for (var i = 0; i < leaveRequests.length; i++) {
            let textBlock = {
                "type": "TextBlock",
                "color": "dark",
                "wrap": true,
                "text": leaveRequests[i].date + " ( " + leaveRequests[i].reason + " )"
            }
            textBlocks.push(textBlock);
        }
        card.body = textBlocks;

        return CardFactory.adaptiveCard(card);
    }

    static filterByDate(requestRange, submittedLeaveDate) {
        try {
            const results = Recognizers.recognizeDateTime(requestRange, Recognizers.Culture.English);
            const now = new Date();
            const earliest = now.getTime();
            let output;
            results.forEach(function (result) {
                // result.resolution is a dictionary, where the "values" entry contains the processed input.
                result.resolution['values'].forEach(function (resolution) {
                    // The processed input contains a "value" entry if it is a date-time value, or "start" and
                    // "end" entries if it is a date-time range.
                    const datevalue = resolution['value'] || resolution['start'];
                    // If only time is given, assume it's for today.
                    const datetime = new Date(datevalue);
                    if ([0, 6].includes(datetime.getDay())) {
                        output = { success: false, message: "The date you have mentioned falls on weekend." };
                        return;
                    }

                    if (datetime && earliest < datetime.getTime()) {
                        output = { success: true, date: result, startDate: datetime.toLocaleDateString() };
                        return;
                    }
                });
            });
            return output || { success: false, message: "I'm sorry, please enter an upcoming date." };
        } catch (error) {
            return {
                success: false,
                message: "I'm sorry, I could not interpret that as an appropriate date. Please enter an upcoming date."
            };
        }
    }
}

module.exports.LeaveRequestManager = LeaveRequestManager;

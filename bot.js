// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityTypes } = require('botbuilder');
const { LuisRecognizer } = require('botbuilder-ai');
const { WelcomeUser } = require('./welcome_user');
const { HolidayCalendar } = require('./holiday_calendar');
const { LeaveRequestManager } = require('./leave_request_manager');
const { TimexProperty } = require('@microsoft/recognizers-text-data-types-timex-expression');
const Recognizers = require('@microsoft/recognizers-text-suite');

// State Accessor Properties
const CONVERSATION_STATE_ACCESSOR = 'conversationData';
const USER_STATE_ACCESSOR = 'userData';
const USER_LEAVE_RECORDS_FILE = './resources/user_leave_record.json';

// Luis Intents
const GREETING_INTENT = "Greeting";
const NONE_INTENT = "None";
const HOLIDAY_INTENT = "Upcoming Holidays";
const INTRODUCTION_INTENT = "Introduction";
const LEAVE_REQUESTS = "leave_requests";
const APPLY_ACTION = "apply";
const SHOW_ACTION = "show";
const Action_Types = "action_types";

const DATE_TIME = "datetime";


var fs = require('fs');

const detail = {
    none: 'none',
    date: 'date',
    reason: 'reason',
    comment: 'comment',
    confirm: 'confirm'
};
class NagarroLeaveManagerBot {
    /**
     *
     * @param {TurnContext} on turn context object.
     */

    constructor(application, luisPredictionOptions, conversationState, userState) {
        this.luisRecognizer = new LuisRecognizer(
            application,
            luisPredictionOptions,
            true
        );
        this.conversationStateAccessor = conversationState.createProperty(CONVERSATION_STATE_ACCESSOR);
        this.userStateAccessor = userState.createProperty(USER_STATE_ACCESSOR);
        this.conversationState = conversationState;
        this.userState = userState;
    }

    async onTurn(turnContext) {
        const activityType = turnContext.activity.type;
        this.welcomeUser = new WelcomeUser();
        switch (activityType) {

            // handle activites of type message.
            case ActivityTypes.Message:

                const userProfile = await this.userStateAccessor.get(turnContext, {});

                const conversationFlow = await this.conversationStateAccessor.get(turnContext, {
                    luisResultForAskedQuestion: [],
                    promptedForEmployeeId: false,
                    promptedForLeaveRequestDetails: detail.none
                });

                if (conversationFlow.promptedForEmployeeId) {
                    if (conversationFlow.promptedForLeaveRequestDetails === detail.none) {
                        userProfile.id = userProfile.id || turnContext.activity.text;
                        await this.userStateAccessor.set(turnContext, userProfile);
                        await this.userState.saveChanges(turnContext);
                        const result = conversationFlow.luisResultForAskedQuestion[0] || await this.luisRecognizer.recognize(turnContext);
                        const topIntent = result.luisResult.topScoringIntent.intent;
                        conversationFlow.luisResultForAskedQuestion.length = [];
                        await this.conversationStateAccessor.set(turnContext, conversationFlow);
                        await this.conversationState.saveChanges(turnContext);

                        if (topIntent === GREETING_INTENT || topIntent === INTRODUCTION_INTENT) {
                            await this.welcomeUser.giveIntroduction(turnContext);
                        } else if (topIntent === HOLIDAY_INTENT) {
                            const holidayCalendar = new HolidayCalendar();
                            await turnContext.sendActivity(holidayCalendar.listHolidays(turnContext, result.entities));
                        }
                        else if (topIntent === LEAVE_REQUESTS) {
                            const leaveRequestManager = new LeaveRequestManager(this.userStateAccessor);
                            if (result.entities[Action_Types]) {
                                if (result.entities[Action_Types][0].includes(APPLY_ACTION)) {
                                    return await this.applyForLeave(userProfile, result.entities, turnContext, conversationFlow);
                                } else if (result.entities[Action_Types][0].includes(SHOW_ACTION)) {
                                    return turnContext.sendActivity(leaveRequestManager.viewSubmittedRequests(userProfile, turnContext, result.entities));
                                }
                            } else {
                                await turnContext.sendActivity("I didn't understand your query.");
                            }
                        } else if (topIntent === NONE_INTENT) {
                            await turnContext.sendActivity("I didn't understand your query.");
                        }
                    } else {
                        const input = turnContext.activity.text;
                        let result;
                        switch (conversationFlow.promptedForLeaveRequestDetails) {
                            case detail.reason:
                                result = this.validateDate(input);
                                if (result.success) {
                                    userProfile.leaveDate = result.startDate;
                                    await turnContext.sendActivity("What is the reason for applying the leave?");
                                    conversationFlow.promptedForLeaveRequestDetails = detail.comment;
                                    await this.conversationStateAccessor.set(turnContext, conversationFlow);
                                    await this.conversationState.saveChanges(turnContext);
                                    await this.userStateAccessor.set(turnContext, userProfile);
                                    await this.userState.saveChanges(turnContext);
                                } else {
                                    await turnContext.sendActivity(
                                        result.message || "I'm sorry, I didn't understand that.");
                                }
                                break;
                            case detail.comment:
                                userProfile.reason = input;
                                await turnContext.sendActivity("Any other comment regarding this leave application?");
                                conversationFlow.promptedForLeaveRequestDetails = detail.confirm;
                                await this.conversationStateAccessor.set(turnContext, conversationFlow);
                                await this.conversationState.saveChanges(turnContext);
                                await this.userStateAccessor.set(turnContext, userProfile);
                                await this.userState.saveChanges(turnContext);
                                break;
                            case detail.confirm:
                                userProfile.comment = input;
                                await turnContext.sendActivity("Please verify your details");
                                await turnContext.sendActivity("Date: " + userProfile.leaveDate + "\nReason: " + userProfile.reason + "\nComment: " + userProfile.comment);
                                await turnContext.sendActivity("Do you confirm (Y/N)?")
                                conversationFlow.promptedForLeaveRequestDetails = detail.submitted;
                                await this.conversationStateAccessor.set(turnContext, conversationFlow);
                                await this.conversationState.saveChanges(turnContext);
                                await this.userStateAccessor.set(turnContext, userProfile);
                                await this.userState.saveChanges(turnContext);
                                break;
                            case detail.submitted:
                                if (input.toLowerCase() === 'y') {
                                    var jsonString = fs.readFileSync('./resources/user_leave_record.json');
                                    var leaveRecords = JSON.parse(jsonString);
                                    for (let i = 0; i < leaveRecords.length; i++) {
                                        if (leaveRecords[i].employeeId === userProfile.id) {
                                            let new_leave_request = {
                                                "reason": userProfile.reason,
                                                "type": "leave",
                                                "date": userProfile.leaveDate,
                                                "comments": userProfile.comment
                                            }
                                            leaveRecords[i].leaveRequests.push(new_leave_request);
                                            leaveRecords[i].leavesTaken += 1;
                                            break;
                                        }
                                    }
                                    var updatedData = JSON.stringify(leaveRecords);
                                    fs.writeFileSync('./resources/user_leave_record.json', updatedData);
                                    conversationFlow.promptedForLeaveRequestDetails = detail.none;

                                    await turnContext.sendActivity("Leave record updated");

                                } else if (input.toLowerCase() === 'n') {
                                    conversationFlow.promptedForLeaveRequestDetails = detail.none;
                                    await turnContext.sendActivity("Cancelling your request.")
                                } else {
                                    // prompt again
                                    await turnContext.sendActivity("Please enter y or n.")
                                }
                                await this.conversationStateAccessor.set(turnContext, conversationFlow);
                                await this.conversationState.saveChanges(turnContext);

                        }
                    }
                } else {
                    // fetch the luis recognizer result of 1st question asked by user, before asking details.
                    const result = await this.luisRecognizer.recognize(turnContext);
                    conversationFlow.luisResultForAskedQuestion.push(result);
                    conversationFlow.promptedForEmployeeId = true;
                    await turnContext.sendActivity('Please provide your employee id.');
                    await this.conversationStateAccessor.set(turnContext, conversationFlow);
                    await this.conversationState.saveChanges(turnContext);
                }
                break;

            // Handle activities of type ConversationUpdate
            case ActivityTypes.ConversationUpdate:
                await this.welcomeUser.welcomeUser(turnContext);
                break;
        }
    }

    async dateValidator(promptContext) {
        if (!promptContext.recognized.succeeded) {
            await promptContext.context.sendActivity(
                "I'm sorry, I do not understand. Please enter the upcoming date."
            );
            return false;
        }
    }
    async applyForLeave(userProfile, entities, turnContext, conversationFlow) {
        const records = fs.readFileSync(USER_LEAVE_RECORDS_FILE);
        const leaveRecords = JSON.parse(records);
        var userLeaveRecord = leaveRecords.find(leaveRecord => leaveRecord.employeeId === userProfile.id);
        if (!userLeaveRecord) {
            return await turnContext.sendActivity("No record found for employee with id: " + userProfile.id);
        } else if (userLeaveRecord.leavesTaken === 27) {
            return await turnContext.sendActivity("You have taken all your leaves. You can not apply for more.");
        } else {
            if (entities[DATE_TIME]) {
                await turnContext.sendActivity("date already meantioned" + new TimexProperty(entities[DATE_TIME][0].timex.toString()))
            } else {
                return await turnContext.sendActivity(this.askForDate(conversationFlow, turnContext));
            }
        }
    }

    async askForDate(conversationFlow, turnContext) {
        await turnContext.sendActivity("When do you want to take leave(s)?");
        conversationFlow.promptedForLeaveRequestDetails = detail.reason;
        await this.conversationStateAccessor.set(turnContext, conversationFlow);
        await this.conversationState.saveChanges(turnContext);
    }

    validateDate(input) {
        try {
            const results = Recognizers.recognizeDateTime(input, Recognizers.Culture.English);
            const now = new Date();
            const earliest = now.getTime();
            let output;
            results.forEach(function (result) {
                result.resolution['values'].forEach(function (resolution) {
                    const datevalue = resolution['value'] || resolution['start'];
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

module.exports.NagarroLeaveManagerBot = NagarroLeaveManagerBot;

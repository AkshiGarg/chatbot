
const { ActivityTypes, CardFactory } = require('botbuilder')

const IntroductionCard = require('./resources/introduction_card.json');
class WelcomeUser {
    async welcomeUser(turnContext) {
        if (turnContext.activity.membersAdded.length !== 0) {
            for (var member in turnContext.activity.membersAdded) {
                if (turnContext.activity.membersAdded[member].id !== turnContext.activity.recipient.id) {
                    await this.giveIntroduction(turnContext);
                }
            }
        }
    }

    async giveIntroduction(turnContext) {
        const reply = {
            text: "Hi, I am Cali !!",
            attachments: [CardFactory.adaptiveCard(IntroductionCard)]
        };
        return await turnContext.sendActivity(reply);
    }
}

module.exports.WelcomeUser = WelcomeUser;
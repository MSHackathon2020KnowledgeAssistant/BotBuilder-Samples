// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TurnContext,
    MessageFactory,
    TeamsActivityHandler,
    CardFactory,
} = require('botbuilder');

class TeamsConversationBot extends TeamsActivityHandler {
    constructor() {
        super();
        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);
            const text = context.activity.text.trim().toLocaleLowerCase();
            if (text.includes('analyze')) {
                await this.analyzeMeetingAsync(context);
            } else {
                await this.sendWelcomeCardAsync(context);
            }
        });

        this.onMembersAddedActivity(async (context, next) => {
            context.activity.membersAdded.forEach(async (teamMember) => {
                if (teamMember.id !== context.activity.recipient.id) {
                    await context.sendActivity(`Welcome to the team ${teamMember.givenName} ${teamMember.surname}`);
                }
            });
            await next();
        });
    }

    async analyzeMeetingAsync(context) {
        const card = CardFactory.adaptiveCard({
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.0",
            "body": [
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "OCR",
                            "weight": "bolder",
                            "size": "medium",
                            "color": "accent"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Optical Character Recognition",
                            "color": "good"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Optical character recognition or optical character reader (OCR) is the electronic or mechanical conversion of images of typed, handwritten or printed text into machine-encoded text.",
                            "wrap": true
                        },
                    ]
                },
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "OCR",
                            "weight": "bolder",
                            "size": "medium",
                            "color": "accent"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Office for Civil Rights",
                            "color": "good"
                        },
                        {
                            "type": "TextBlock",
                            "text": "The Office for Civil Rights (OCR) is a sub-agency of the U.S. Department of Education that is primarily focused on enforcing civil rights laws prohibiting schools from engaging in discrimination on the basis of race, color, national origin, sex, disability, age, or membership in patriotic youth organizations.",
                            "wrap": true
                        },
                    ]
                },
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Form Recognizer",
                            "weight": "bolder",
                            "size": "medium",
                            "color": "accent"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Azure Form Recognizer",
                            "color": "good"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Form Recognizer is part of Azure Cognitive Services, backed by Azure infrastructure and enterprise-grade security, availability, compliance, and manageability.",
                            "wrap": true
                        },
                    ]
                },
            ],
            "actions": [
                {
                    "type": "Action.OpenUrl",
                    "title": "See More",
                    "url": "http://23.96.31.42:8080/?meetingId=23"
                }
            ]
        });

        const message = MessageFactory.attachment(card, "Here are the top keywords from your previous meeting.");
        await context.sendActivity(message);
    }

    async sendWelcomeCardAsync(context) {
        await context.sendActivity(MessageFactory.text("Please type `analyze` to analyze the previous meeting."));
    }
}

module.exports.TeamsConversationBot = TeamsConversationBot;

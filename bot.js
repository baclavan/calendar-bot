const { ActivityHandler, MessageFactory, TurnContext } = require('botbuilder');

class CalendarBot extends ActivityHandler {
  constructor() {
    super();
    // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
    this.onConversationUpdate(async (context, next) => {
      const conversationReference = TurnContext.getConversationReference(context.activity);
      console.log(conversationReference)
      await next();
    });

    this.onMessage(async (context, next) => {
      const replyText = `Echo: ${ context.activity.text }`;
      await context.sendActivity(MessageFactory.text(replyText, replyText));
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      const welcomeText = 'Hello and welcome!';
      for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
          if (membersAdded[cnt].id !== context.activity.recipient.id) {
              await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
          }
      }
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });
  }
}

module.exports.CalendarBot = CalendarBot;

const { ActivityHandler, MessageFactory, TurnContext } = require('botbuilder');

class CalendarBot extends ActivityHandler {
  constructor(conversationReferences) {
    super();

    this.conversationReferences = conversationReferences;

    // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
    this.onConversationUpdate(async (context, next) => {
      this.addConversationReference(context.activity);

      await next();
    });

    this.onMessage(async (context, next) => {
      this.addConversationReference(context.activity);

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

  addConversationReference(activity) {
    const conversationReference = TurnContext.getConversationReference(activity);
    this.conversationReferences[conversationReference.conversation.id] = conversationReference;
  }
}

module.exports.CalendarBot = CalendarBot;

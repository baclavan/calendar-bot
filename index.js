const path = require('path');
const dotenv = require('dotenv');
const restify = require('restify');
const { BotFrameworkAdapter } = require('botbuilder');
const { CalendarBot } = require('./bot');

dotenv.config({ path: path.join(__dirname, '.env') });

const conversationReferences = {};
const myBot = new CalendarBot(conversationReferences);
const server = restify.createServer();

const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword
});

// Catch-all for errors.
const onTurnErrorHandler = async (context, error) => {
  console.error(`\n [onTurnError] unhandled error: ${ error }`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    'OnTurnError Trace',
    `${ error }`,
    'https://www.botframework.com/schemas/error',
    'TurnError'
  );

  // Send a message to the user
  await context.sendActivity('The bot encountered an error or bug.');
  await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Set the onTurnError for the singleton BotFrameworkAdapter.
adapter.onTurnError = onTurnErrorHandler;

server.use(restify.plugins.bodyParser());

// Listen for incoming requests.
server.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    // Route to main dialog.
    await myBot.run(context);
  });
});

// Listen for Upgrade requests for Streaming.
server.on('upgrade', (req, socket, head) => {
  // Create an adapter scoped to this WebSocket connection to allow storing session data.
  const streamingAdapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
  });
  // Set onTurnError for the BotFrameworkAdapter created for each connection.
  streamingAdapter.onTurnError = onTurnErrorHandler;

  streamingAdapter.useWebSocket(req, socket, head, async (context) => {
    // After connecting via WebSocket, run this logic for every request sent over
    // the WebSocket connection.
    await myBot.run(context);
  });
});

// Demo send proactive messages to users.
server.get('/proactive', async (req, res) => {
  res.setHeader('Content-Type', 'text/html');
  res.writeHead(200);
  res.write('<html><body>');
  res.write('<h1>Proactive messages</h1>');
  res.write('<form method="POST"><select name="id">');
  for (let id in conversationReferences) {
    res.write('<option value="'+ id +'">' + id + '</option>')
  }
  res.write('</select><br><textarea name="content"></textarea><br><button type="submit">Send</button></form>')
  res.write('</body></html>');
  res.end();
});

server.post('/proactive', async (req, res) => {
  let conversationReference = conversationReferences[req.body.id]

  await adapter.continueConversation(conversationReference, async turnContext => {
    await turnContext.sendActivity(req.body.content);
  });

  res.setHeader('Content-Type', 'text/html');
  res.writeHead(200);
  res.write('<html><body><h1>Notification have been sent.</h1></body></html>');
  res.end();
});

server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\n${ server.name } listening to ${ server.url }`);
});

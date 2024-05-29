import * as restify from "restify";
import * as fs from "fs";
import send from "send";
import { BotFrameworkAdapter } from "botbuilder";
import { TeamsBot } from "./MyBot";
import msal from '@azure/msal-node';
const fetch = require("cross-fetch");

require('dotenv').config();

const msalClient = new msal.ConfidentialClientApplication({
  auth: {
    clientId: process.env.MicrosoftAppId,
    clientSecret: process.env.MicrosoftAppPassword,
    authority: process.env.Authority,
  },
});

//Create HTTP server.
const server = restify.createServer({
  key: process.env.SSL_KEY_FILE ? fs.readFileSync(process.env.SSL_KEY_FILE) : undefined,
  certificate: process.env.SSL_CRT_FILE ? fs.readFileSync(process.env.SSL_CRT_FILE) : undefined,
  formatters: {
    "text/html": (req, res, body) => {
      return body;
    },
  },
});

server.use(restify.plugins.bodyParser({
  mapParams: true
}));


/****** TABs ********/
server.get(
  "/static/*",
  restify.plugins.serveStatic({
    directory: __dirname,
  })
);

server.listen(process.env.port || process.env.PORT || 3000, function () {
  console.log(`\n${server.name} listening to ${server.url}`);
});

// Adding tabs to our app. This will setup routes to various views
// Setup home page
server.get("/", (req, res, next) => {
  send(req, __dirname + "/views/configure.html").pipe(res);
});

server.get("/configure", (req, res, next) => {
  send(req, __dirname + "/views/configure.html").pipe(res);
});

// Setup the static tab
server.get("/tab", (req, res, next) => {
  send(req, __dirname + "/views/basic-tab.html").pipe(res);
});

server.post("/user-token/me", async (req, res) => {
  const token = req.body.token;
  if (!token) {
    return res.json({});
  }
  const scopes = ["https://graph.microsoft.com/User.Read"];
  const authResponse = await msalClient.acquireTokenOnBehalfOf({
    authority: `https://login.microsoftonline.com/${req.body.tid}`,
    oboAssertion: token,
    scopes: scopes,
    skipCache: false
  });
  const accessToken = authResponse.accessToken;
  const user = await fetch("https://graph.microsoft.com/v1.0/me/",
    {
      method: 'GET',
      headers: {
        "accept": "application/json",
        "authorization": "bearer " + accessToken
      }
    }
  ).then(response => {
    return response.json();
  });
  res.json(user);
});

server.post("/user-token/teams/:teamId/channels/:channelId/members", async (req, res) => {
  const teamId = req.params.teamId;
  const channelId = req.params.channelId;
  const token = req.body.token;
  if (!token) {
    return res.json({});
  }
  const scopes = ["https://graph.microsoft.com/ChannelMember.Read.All"];
  const authResponse = await msalClient.acquireTokenOnBehalfOf({
    authority: `https://login.microsoftonline.com/${req.body.tid}`,
    oboAssertion: token,
    scopes: scopes,
    skipCache: false
  });
  const accessToken = authResponse.accessToken;
  const members = await getChannelMembers(teamId, channelId, accessToken);
  res.json(members);
});

server.post("/app-token/teams/:teamId/channels/:channelId/members", async (req, res) => {
  const teamId = req.params.teamId;
  const channelId = req.params.channelId;
  const scopes = ["https://graph.microsoft.com/.default"];
  const authResponse = await msalClient.acquireTokenByClientCredential({
    scopes: scopes,
  });
  const accessToken = authResponse.accessToken;
  console.log(accessToken);
  const members = await getChannelMembers(teamId, channelId, accessToken);
  res.json(members);
});


const getChannelMembers = async (teamId: string, channelId: string, token: string) => {
  return await fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/members`,
    {
      method: 'GET',
      headers: {
        "accept": "application/json",
        "authorization": "bearer " + token
      }
    }
  ).then(response => {
    return response.json();
  });
}

/****** BOTs ********/
const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword,
});


adapter.onTurnError = async (context, error) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);
  await context.sendActivity(`Oops. Something went wrong! ${error}`);
};

const bot = new TeamsBot();
server.post("/api/messages", async (req, res) => {
  await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

server.get('/api/notify', async (req, res) => {
  for (const conversationReference of Object.values(bot.getAllConversations())) {
    await adapter.continueConversation(conversationReference, async (context) => {
      await context.sendActivity('proactive hello');
    });
  }
  res.send('Proactive messages have been sent.')
});
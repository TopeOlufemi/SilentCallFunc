const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const qs = require('querystring');
const fetch = require('node-fetch');

module.exports = async function (context, req) {
  const userEmail = req.query.userEmail;
  if (!userEmail) {
    context.res = { status: 400, body: "Missing 'userEmail' query parameter." };
    return;
  }

  try {
    const token = await getToken();
    const graph = Client.init({ authProvider: done => done(null, token) });

    const user = await graph.api(`/users/${userEmail}`).get();
    const chat = await graph.api('/chats').post({
      chatType: 'oneOnOne',
      members: [
        {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          roles: ['owner'],
          user@odata.bind: `https://graph.microsoft.com/v1.0/users/${user.id}`
        }
      ]
    });

    const now = new Date().toLocaleTimeString();
    await graph.api(`/chats/${chat.id}/messages`).post({
      body: { content: `🔕 You missed a call from SilentCallBot at ${now}.` }
    });

    context.res = { status: 200, body: `Silent call sent to ${userEmail}` };
  } catch (err) {
    context.log.error(err);
    context.res = { status: 500, body: err.message };
  }
};

async function getToken() {
  const body = qs.stringify({
    grant_type: 'client_credentials',
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default'
  });

  const res = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });

  const json = await res.json();
  return json.access_token;
}


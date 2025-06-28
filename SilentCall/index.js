const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;

async function getToken() {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append("client_id", clientId);
  params.append("scope", "https://graph.microsoft.com/.default");
  params.append("client_secret", clientSecret);
  params.append("grant_type", "client_credentials");

  const res = await fetch(url, {
    method: "POST",
    body: params,
  });

  const data = await res.json();
  if (!res.ok) {
    throw new Error(`Token error: ${JSON.stringify(data)}`);
  }

  return data.access_token;
}

module.exports = async function (context, req) {
  context.log("Starting SilentCall function");

  const userEmail = req.query.userEmail;
  context.log("Received userEmail:", userEmail);

  if (!userEmail) {
    context.res = {
      status: 400,
      body: "Missing 'userEmail' query parameter.",
    };
    return;
  }

  try {
    const token = await getToken();
    context.log("Token acquired");

    const client = Client.init({
      authProvider: (done) => done(null, token),
    });

    const user = await client.api(`/users/${userEmail}`).get();
    context.log("User retrieved:", user.id);

    const chat = await client.api('/chats').post({
      chatType: 'oneOnOne',
      members: [{
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.micr



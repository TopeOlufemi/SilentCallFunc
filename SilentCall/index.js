const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

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

  const response = await fetch(url, { method: "POST", body: params });
  const data = await response.json();

  if (!response.ok) {
    throw new Error(`Token error: ${JSON.stringify(data)}`);
  }

  return data.access_token;
}

module.exports = async function (context, req) {
  context.log("💡 SilentCall function triggered");

  const userEmail = req.query.userEmail;
  context.log("📩 userEmail parameter:", userEmail);

  if (!userEmail) {
    context.res = {
      status: 400,
      body: "❌ Missing 'userEmail' query parameter.",
    };
    return;
  }

  try {
    context.log("🔐 Getting Microsoft Graph token...");
    const token = await getToken();
    context.log("✅ Token acquired");

    const client = Client.init({
      authProvider: (done) => done(null, token),
    });

    context.log(`🔎 Looking up user: ${userEmail}`);
    const user = await client.api(`/users/${userEmail}`).get();
    context.log("👤 User found:", user.id);

    context.log("💬 Creating chat");
    const chat = await client.api("/chats").post({
      chatType: "oneOnOne",
      members: [
        {
          "@odata.type": "#microsoft.graph.aadUserConversationMember",
          roles: ["owner"],
          "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${user.id}`,
        },
      ],
    });

    context.log("✉️ Sending message to chat");
    await client.api(`/chats/${chat.id}/messages`).post({
      body: {
        content: "🔕 SilentCallBot: You have a missed call.",
      },
    });

    context.res = {
      status: 200,
      body: `✅ Silent call registered for ${userEmail}`,
    };
  } catch (err) {
    context.log.error("❌ Error occurred:", err.message);
    context.res = {
      status: 500,
      body: `⚠️ Error: ${err.message}`,
    };
  }
};




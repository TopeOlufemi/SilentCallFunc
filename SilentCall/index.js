/**
 * SilentCall Azure Function  ─ JavaScript (Node 20+)
 * --------------------------------------------------
 * • Reads TENANT_ID, CLIENT_ID, CLIENT_SECRET from app settings
 * • Accepts GET or POST:
 *      └ query/body param  userEmail   (UPN or email address)
 * • Fetches a client-credentials token on every call (simple & always fresh)
 * • Makes a harmless Graph request   GET /users/{userEmail}
 *   ↳ replace with any Graph action you need (sendMail, chat message, etc.)
 * • Responds 200 on success, 4xx/5xx on error with helpful text
 */

const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const qs = require('querystring');
const fetch = require('node-fetch');

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET
} = process.env;

/* ────────────────────────────────────────────────── */
/* Helper: fetch app-only Microsoft Graph token      */
async function getAppToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = qs.stringify({
    grant_type: 'client_credentials',
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default'
  });

  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });

  const json = await res.json();
  if (!res.ok) {
    throw new Error(`Token fetch failed: ${json.error_description || res.statusText}`);
  }
  return json.access_token;
}

/* ────────────────────────────────────────────────── */
/* Main function entry                                */
module.exports = async function (context, req) {
  context.log('🔔 SilentCall function triggered');

  /* 1️⃣  Grab userEmail from query or body */
  const userEmail =
    (req.query && req.query.userEmail) ||
    (req.body && req.body.userEmail);

  if (!userEmail) {
    context.log.warn("❌ Missing 'userEmail'");
    context.res = { status: 400, body: "Provide 'userEmail' as query or JSON body." };
    return;
  }

  try {
    /* 2️⃣  Get fresh app token */
    context.log('🔐 Fetching Graph token…');
    const token = await getAppToken();

    /* 3️⃣  Init Graph client */
    const graph = Client.init({ authProvider: done => done(null, token) });

    /* 4️⃣  Simple Graph call (replace with your own logic) */
    context.log(`🔎 Querying Graph for user ${userEmail}`);
    const user = await graph.api(`/users/${userEmail}`).get();

    context.log('✅ Graph lookup ok:', user.id);

    /* 5️⃣  Example: send a Teams chat or email here… */

    context.res = {
      status: 200,
      body: `Silent call processed for ${userEmail}`
    };
  } catch (err) {
    context.log.error('🔥 Error:', err.message);
    context.res = {
      status: 500,
      body: `Internal error: ${err.message}`
    };
  }
};


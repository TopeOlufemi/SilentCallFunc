const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

module.exports = async function (context, req) {
    context.log("🔔 SilentCall function triggered");

    const userEmail = req.query.userEmail || (req.body && req.body.userEmail);

    if (!userEmail) {
        context.log("❌ Missing 'userEmail'");
        context.res = {
            status: 400,
            body: "Missing 'userEmail' query parameter."
        };
        return;
    }

    context.log(`📨 Simulating silent call for: ${userEmail}`);

    try {
        // Access token from environment settings (Application Settings)
        const accessToken = process.env.GRAPH_TOKEN;

        if (!accessToken) {
            context.log("❌ Missing GRAPH_TOKEN environment variable");
            context.res = {
                status: 500,
                body: "GRAPH_TOKEN not configured."
            };
            return;
        }

        const client = Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });

        // Fake Microsoft Graph call to simulate missed call — adjust as needed
        await client.api(`/users/${userEmail}/messages`).get();

        context.res = {
            status: 200,
            body: `Silent call registered for ${userEmail}.`
        };

    } catch (error) {
        context.log.error("🔥 Error executing Graph request:", error);
        context.res = {
            status: 500,
            body: "Internal Server Error"
        };
    }
};

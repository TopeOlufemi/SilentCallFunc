module.exports = async function (context, req) {
    context.log('SilentCall function processed a request.');

    const userEmail = req.body?.userEmail || req.query?.userEmail;

    if (!userEmail) {
        context.res = {
            status: 400,
            body: "Missing userEmail in request"
        };
        return;
    }

    context.res = {
        status: 200,
        body: `Silent call simulated for user ${userEmail}`
    };
};

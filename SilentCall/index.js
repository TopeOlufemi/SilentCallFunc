module.exports = async function (context, req) {
    const userEmail = req.query.userEmail;

    if (!userEmail) {
        context.res = {
            status: 400,
            body: "Missing \'userEmail\' query parameter."
        };
        return;
    }

    context.log(`Silent call initiated for ${userEmail}`);

    context.res = {
        status: 200,
        body: `Silent call registered for ${userEmail}.`
    };
};
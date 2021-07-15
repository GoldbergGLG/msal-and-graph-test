const { ConfidentialClientApplication } = require("@azure/msal-node");

async function getToken() {
  // initialize "Confidential" Client Application (as opposed to "Public" ie. browser applications)
  const tokenResponse = await new ConfidentialClientApplication({
    auth: {
      authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
      clientId: process.env.CLIENT_ID,
      clientSecret: process.env.CLIENT_SECRET, // Only for Confidential Client Applications
    },
    system: {
      loggerOptions: {
        loggerCallback(loglevel, message, containsPii) {
          // console.log(message);
        },
        piiLoggingEnabled: false,
        // logLevel: msal.LogLevel.Verbose,
      },
    },
  }).acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return tokenResponse.accessToken;
}

module.exports = { getToken };

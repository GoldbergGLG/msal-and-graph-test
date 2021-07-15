require("dotenv").config();

// Code adapted from the following:
// https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/samples/msal-node-samples/auth-code
// https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-node-migration

// NOTE: can comment in the 2 log items in system prop of the config object of msal in the getToken()

console.log("TENANT_ID:", process.env.TENANT_ID);
console.log("CLIENT_ID:", process.env.CLIENT_ID);
console.log("CLIENT_SECRET:", process.env.CLIENT_SECRET);
console.log("MAILBOX:", process.env.MAILBOX);

const { getToken } = require("./auth");
require("isomorphic-fetch");
const { Client } = require("@microsoft/microsoft-graph-client");

const createDraft = async (client, mailbox) => {
  return await client.api(`/users/${mailbox}/messages`).post({
    subject: "my test from test app",
    body: {
      contentType: "Text",
      content: "The test app works.",
    },
    toRecipients: [
      {
        emailAddress: {
          address: "dgoldberg@glgroup.com",
        },
      },
    ],
  });
};

async function main() {
  const accessToken = await getToken();

  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });

  console.log(await createDraft(client, process.env.MAILBOX));
}

main();

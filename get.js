require('dotenv').config()

// Code adapted from the following:
// https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/samples/msal-node-samples/auth-code
// https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-node-migration

// NOTE: can comment in the 2 log items in system prop of the config object of msal in the getToken()

console.log("TENANT_ID:", process.env.TENANT_ID);
console.log("CLIENT_ID:", process.env.CLIENT_ID);
console.log("CLIENT_SECRET:", process.env.CLIENT_SECRET);
console.log("MAILBOX:", process.env.MAILBOX);

const { getToken } = require("./auth")
require("isomorphic-fetch");
const { Client } = require("@microsoft/microsoft-graph-client");

const getEmails = async (accessToken, mailbox) => {
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });

  let messages = await client
    .api(`/users/${mailbox}/messages`)
    .get();
  return messages;
};

async function main() {
  const accessToken = await getToken();

  console.log(await getEmails(accessToken, process.env.MAILBOX));
}

main();

// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

const getAccessToken = async function () {
    // Create a config variable that store credentials from config.json
    const config = require(__dirname + "/../config/config.json");

    const tenant_id = process.env.tenantId;
    if (!tenant_id) {
        console.log("Tenant id not found");
    }
    const client_secret = process.env.clientSecret;
    if (!client_secret) {
        console.log("client secret not found");
    }
    const client_id = process.env.CLIENTID;
    if (!client_id) {
        console.log("client id not found");
    }

    // Use MSAL.js for authentication
    const msal = require("@azure/msal-node");

    const msalConfig = {
        auth: {
            clientId: client_id, // config.clientId,
            authority: `${config.authorityUrl}${tenant_id}`,    // {config.tenantId}
        }
    };

    // Check for the MasterUser Authentication
    if (config.authenticationMode.toLowerCase() === "masteruser") {
        const clientApplication = new msal.PublicClientApplication(msalConfig);

        const usernamePasswordRequest = {
            scopes: [config.scopeBase],
            username: config.pbiUsername,
            password: config.pbiPassword
        };

        return clientApplication.acquireTokenByUsernamePassword(usernamePasswordRequest);
    };

    // Service Principal auth is the recommended by Microsoft to achieve App Owns Data Power BI embedding
    if (config.authenticationMode.toLowerCase() === "serviceprincipal") {
        msalConfig.auth.clientSecret =  client_secret;  // config.clientSecret
        const clientApplication = new msal.ConfidentialClientApplication(msalConfig);

        const clientCredentialRequest = {
            scopes: [config.scopeBase],
        };

        // console.log(`Contents of NodeJs.Require Config ${JSON.stringify(config)}`);
        // console.log(`MSAL Client Application (msal.ConfidentialClientApplication): ${JSON.stringify(clientApplication)}`);
        // console.log("Calling clientApplication.acquireTokenByClientCredential");

        return clientApplication.acquireTokenByClientCredential(clientCredentialRequest);
    }
}

module.exports.getAccessToken = getAccessToken;
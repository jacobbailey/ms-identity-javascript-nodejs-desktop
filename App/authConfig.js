/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

const { LogLevel } = require("@azure/msal-node");

/**
 * Configuration object to be passed to MSAL instance on creation.
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/configuration.md
 */
const AAD_ENDPOINT_HOST = "Enter_the_Cloud_Instance_Id_Here"; // include the trailing slash
const REDIRECT_URI = "Enter_the_Redirect_Uri_Here";

const msalConfig = {
  auth: {
    clientId: "Enter_the_Application_Id_Here",
    authority: `${AAD_ENDPOINT_HOST}Enter_the_Tenant_Info_Here`,
  },
  system: {
    loggerOptions: {
      loggerCallback(loglevel, message, containsPii) {
        console.log(message);
      },
      piiLoggingEnabled: false,
      logLevel: LogLevel.Verbose,
    },
  },
};

/**
 * Add here the endpoints and scopes when obtaining an access token for protected web APIs. For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */
const GRAPH_ENDPOINT_HOST = "Enter_the_Graph_Endpoint_Here"; // include the trailing slash

const protectedResources = {
  graphMe: {
    endpoint: `${GRAPH_ENDPOINT_HOST}v1.0/me`,
    scopes: ["User.Read"],
  },
  graphMessages: {
    endpoint: `${GRAPH_ENDPOINT_HOST}v1.0/me/messages`,
    scopes: ["Mail.Read"],
  },
};


module.exports = {
  msalConfig: msalConfig,
  protectedResources: protectedResources,
  REDIRECT_URI: REDIRECT_URI,
};

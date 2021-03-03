/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
const express = require('express');
const msal = require('@azure/msal-node');

const SERVER_PORT = process.env.PORT || 3000;

// Before running the sample, you will need to replace the values in the config,
// including the clientSecret
const config = {
    auth: {
        clientId: '26c3029e-6a32-424d-b557-36430eec07c4',
        authority:
            'https://login.microsoftonline.com/23a51087-bf44-49a2-ae57-2005668fec39',
        clientSecret: 'd_2Rb~eKBV24347oADL-~fqWjXuVC_442F',
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Verbose,
        },
    },
};

// Create msal application object
const pca = new msal.ConfidentialClientApplication(config);

// Create Express App and Routes
const app = express();

app.get('/', (req, res) => {
    const authCodeUrlParameters = {
        scopes: ['user.read'],
        redirectUri: 'http://localhost:3000/redirect',
    };

    // get url to sign user in and consent to scopes needed for application
    pca.getAuthCodeUrl(authCodeUrlParameters)
        .then((response) => {
            res.redirect(response);
        })
        .catch((error) => console.log(JSON.stringify(error)));
});

app.get('/redirect', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ['user.read'],
        redirectUri: 'http://localhost:3000/redirect',
    };

    pca.acquireTokenByCode(tokenRequest)
        .then((response) => {
            console.log('\nResponse: \n:', response);
            res.sendStatus(200);
        })
        .catch((error) => {
            console.log(error);
            res.status(500).send(error);
        });
});

app.listen(SERVER_PORT, () =>
    console.log(
        `Msal Node Auth Code Sample app listening on port ${SERVER_PORT}!`
    )
);

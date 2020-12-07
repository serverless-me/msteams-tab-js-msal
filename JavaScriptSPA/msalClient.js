
let msalConfig = {
    auth: {
        clientId: "<APP_ID>",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://<NGROK_DOMAIN>.ngrok.io/auth-end.html",
        navigateToLoginRequestUrl: false
    },
    cache: {
        cacheLocation: "localStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    }
};

// MSAL request object to use over and over
let tokenRequest = {
    scopes: ["User.Read", "Mail.Read", "Files.Read"],
    extraQueryParameters: {domain_hint: 'organizations'}
}

// Keep this MSAL client around to manage state across SPA "pages"
let msalClient = new Msal.UserAgentApplication(msalConfig);
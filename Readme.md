# Documentation on Authentication and Authorization for Teams Tabs

How to use MSAL in a Teams Tab is described in this [blog post by Bob German](https://bob1german.com/2020/08/31/calling-microsoft-graph-from-your-teams-application-part3/#3-teams-pop-up-with-msal-20). The blog post is linked to this [sample using REACT](https://github.com/pnp/teams-dev-samples/tree/master/samples/tab-aad-msal2)

The sample in this repository is based on the same approach, but uses only Node.js and JavaScript.

# Authentication & Authorization Flow

This sample is using only JavaScript for the implementation of the authentication popup:

1. If the user is already authorized, there will be no popup and the token will be acquired silently:

```
    if (msalClient.getAccount()) {
        microsoftTeams.getContext(function (context, error) {
            tokenRequest.loginHint = context.loginHint;
            msalClient.acquireTokenSilent(tokenRequest)
            .then(response => {
                graphApp.accessToken = response.accessToken;
                graphCallback();
            })
            .catch(error => {
                console.log(error);
                console.log("silent token acquisition fails. acquiring token using popup");
                    
                teamsAuth.signIn(graphCallback);
            });;
        });
    }
```


2. a popup is created using Microsoft Teams SDK (auth-start.html)

```
    signIn: function(graphCallback){        
        microsoftTeams.authentication.authenticate({
            url: window.location.origin + "/auth-start.html",
            width: 600,
            height: 535,
            successCallback: function (response) {
                console.log("Successfull callback");
                graphApp.accessToken = response.accessToken;
                graphCallback();
            },
            failureCallback: function (reason) {
                console.log("callback failure");
            }
        });
    }
```

3. within the popup there is a redirect to Azure AD login page using MSAL 

```
    if (msalClient.getAccount()) {
        tokenRequest.loginHint = msalClient.getAccount().userName;
        msalClient.acquireTokenRedirect(tokenRequest);
    }
    else{
        msalClient.loginRedirect(tokenRequest);
    }
```

4. after login the user is redirected to auth-end.html (specified in Azure App Registration) and the successfull login is confirmed using Microsoft Teams SDK

```
    if (msalClient.getAccount()) {
        tokenRequest.loginHint = msalClient.getAccount().userName;
        msalClient.acquireTokenSilent(tokenRequest)
        .then(tokenResponse => {
            microsoftTeams.authentication.notifySuccess(tokenResponse);
        })
        .catch(error => {
            microsoftTeams.authentication.notifyFailure(error);
        });
    }
```

# Prerequisites

## Azure App Registration

1. Authentication

![Authentication](/docs/Az-Authentication.jpg "Authentication")

2. API Permissions

![ApiPermissions](/docs/Az-ApiPermissions.jpg "API Permissions")

3. Expose API

![Expose API](/docs/Az-ExposeApi.jpg "Expose API")

## Replace the placeholders with your ngrok domain and App ID

1. In the manifest.json
2. In the msalClient.js

## Run the sample with NGROK

```
ngrok http -host-header=rewrite 3978
```

# References

- [Microsoft Teams authentication flow for tabs](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-flow-tab)

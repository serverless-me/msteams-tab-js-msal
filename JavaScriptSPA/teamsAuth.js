microsoftTeams.initialize();

let teamsAuth = {
    token: "",
    // https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
    sso: function() {
        microsoftTeams.authentication.getAuthToken({
            successCallback: function(result) { 
                console.log("Success: " + result); 
                var decoded = jwt_decode(result);
                updateProfile(decoded);
            },
            failureCallback: function(error) { 
                console.log("Failure: " + error); 
            }
        });
    },    
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
    }, 
    getToken: function(graphCallback) {        
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
        else{
            teamsAuth.signIn(graphCallback);
        }
    }
};
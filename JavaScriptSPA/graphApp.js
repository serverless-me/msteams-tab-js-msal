

let graphApp = {
  accessToken: '',
  graphMailEndpoint: "https://graph.microsoft.com/v1.0/me/messages",
  graphFilesEndpoint: "https://graph.microsoft.com/v1.0/me/drive/root/children",  
  
  // Helper function to call MS Graph API endpoint 
  // using authorization bearer token scheme
  callMSGraph: function (endpoint) {
    const headers = new Headers();
    const bearer = `Bearer ${this.accessToken}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    console.log('request made to Graph API at: ' + new Date().toString());
    
    fetch(endpoint, options)
      .then(response => response.json())
      .then(response => updateUI(response, endpoint))
      .catch(error => console.log(error))
  },

  readMail: function () {    
    if(graphApp.accessToken != ''){
      graphApp.callMSGraph(graphApp.graphMailEndpoint);
      mailButton.classList.add("d-none");
      filesButton.classList.remove("d-none");
    }
    else{
      teamsAuth.getToken(graphApp.readMail);
    }
  },

  getFiles: function () {
    if(graphApp.accessToken != ''){
      graphApp.callMSGraph(graphApp.graphFilesEndpoint);
      filesButton.classList.add("d-none");
      mailButton.classList.remove("d-none");
    }
    else{
      teamsAuth.getToken(graphApp.getFiles);
    }
  }

}
// Select DOM elements to work with
const welcomeDiv = document.getElementById("welcomeMessage");
const signInButton = document.getElementById("signIn");
const signOutButton = document.getElementById('signOut');
const cardDiv = document.getElementById("card-div");
const mailButton = document.getElementById("readMail");
const filesButton = document.getElementById("getFiles");
const profileDiv = document.getElementById("profile-div");



function updateProfile(data) {
  console.log('OpenID responded at: ' + new Date().toString());
    const name = document.createElement('p');
    name.innerHTML = "<strong>Name: </strong>" + data.name;
    const upn = document.createElement('p');
    upn.innerHTML = "<strong>UPN: </strong>" + data.upn;
    const tenant = document.createElement('p');
    tenant.innerHTML = "<strong>Tenant: </strong>" + data.tid;
    profileDiv.appendChild(name);
    profileDiv.appendChild(upn);
    profileDiv.appendChild(tenant);    
    mailButton.classList.remove("d-none");
    filesButton.classList.remove("d-none");
 }
 
function updateUI(data, endpoint) {
  if(endpoint == graphApp.graphMailEndpoint){
    if (data.value.length < 1) {
      alert("Your mailbox is empty!")
    } else {
      const tabList = document.getElementById("list-tab");
      tabList.innerHTML = ''; // clear tabList at each readMail call
      const tabContent = document.getElementById("nav-tabContent");

      data.value.map((d, i) => {
        // Keeping it simple
        if (i < 10) {
          const listItem = document.createElement("a");
          listItem.setAttribute("class", "list-group-item list-group-item-action")
          listItem.setAttribute("id", "list" + i + "list")
          listItem.setAttribute("data-toggle", "list")
          listItem.setAttribute("href", "#mail" + i)
          listItem.setAttribute("role", "tab")
          listItem.setAttribute("aria-controls", i)
          listItem.innerHTML = d.subject;
          tabList.appendChild(listItem)
  
          const contentItem = document.createElement("div");
          contentItem.setAttribute("class", "tab-pane fade")
          contentItem.setAttribute("id", "mail" + i)
          contentItem.setAttribute("role", "tabpanel")
          contentItem.setAttribute("aria-labelledby", "list" + i + "list")
          contentItem.innerHTML = "<strong> from: " + d.from.emailAddress.address + "</strong><br><br>" + d.bodyPreview + "...";
          tabContent.appendChild(contentItem);
        }
      });
    }
  }
  else if(endpoint == graphApp.graphFilesEndpoint){    
    const tabList = document.getElementById("list-tab");
    tabList.innerHTML = ''; // clear tabList at each readMail call
    const tabContent = document.getElementById("nav-tabContent");

    data.value.map((d, i) => {
      // Keeping it simple
      if (i < 10) {
        const listItem = document.createElement("a");
        listItem.setAttribute("class", "list-group-item list-group-item-action")
        listItem.setAttribute("id", "list" + i + "list")
        listItem.setAttribute("data-toggle", "list")
        listItem.setAttribute("href", "#file" + i)
        listItem.setAttribute("role", "tab")
        listItem.setAttribute("aria-controls", i)
        listItem.innerHTML = d.name;
        tabList.appendChild(listItem)

        const contentItem = document.createElement("div");
        contentItem.setAttribute("class", "tab-pane fade")
        contentItem.setAttribute("id", "file" + i)
        contentItem.setAttribute("role", "tabpanel")
        contentItem.setAttribute("aria-labelledby", "list" + i + "list")
        contentItem.innerHTML = "<strong> Name: " + d.name + "</strong><br><br>Modified: " + d.fileSystemInfo.lastModifiedDateTime; 
        tabContent.appendChild(contentItem);
      }
    });
  }   
}
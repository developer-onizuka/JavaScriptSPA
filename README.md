# JavaScriptSPA

![diagram-01-auth-code-flow.png](https://learn.microsoft.com/ja-jp/azure/active-directory/develop/media/tutorial-v2-javascript-auth-code/diagram-01-auth-code-flow.png)

# 1. Install npm and some libraries
```
sudo apt install -y npm
npm init -y
npm install @azure/msal-browser
npm install express
npm install morgan
npm install yargs
```

# 2. Create server.js
```
const express = require('express');
const morgan = require('morgan');
const path = require('path');

const DEFAULT_PORT = process.env.PORT || 3000;

// initialize express.
const app = express();

// Initialize variables.
let port = DEFAULT_PORT;

// Configure morgan module to log all requests.
app.use(morgan('dev'));

// Setup app folders.
app.use(express.static('app'));

// Set up a route for index.html
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname + '/index.html'));
});

// Start the server.
app.listen(port);
console.log(`Listening on port ${port}...`);
```

# 3. Create UI
Create two files, **index.html and ui.js**.
```
mkdir app
cd app
vi index.html
```
```
<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, shrink-to-fit=no">
  <title>Microsoft identity platform</title>
  <link rel="SHORTCUT ICON" href="./favicon.svg" type="image/x-icon">

   <!-- msal.min.js can be used in the place of msal.js; included msal.js to make debug easy -->
  <script src="https://alcdn.msauth.net/browser/2.30.0/js/msal-browser.js"
    integrity="sha384-o4ufwq3oKqc7IoCcR08YtZXmgOljhTggRwxP2CLbSqeXGtitAxwYaUln/05nJjit"
    crossorigin="anonymous"></script>
  
  <!-- adding Bootstrap 4 for UI components  -->
  <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css"
    integrity="sha384-Vkoo8x4CGsO3+Hhxv8T/Q5PaXtkKtu6ug5TOeNV6gBiFeWPGFN9MuhOf23Q9Ifjh" crossorigin="anonymous">
  <link rel="SHORTCUT ICON" href="https://c.s-microsoft.com/favicon.ico?v2" type="image/x-icon">
</head>

<body>
  <nav class="navbar navbar-expand-lg navbar-dark bg-primary">
    <a class="navbar-brand" href="/">Microsoft identity platform</a>
    <div class="btn-group ml-auto dropleft">
      <button type="button" id="SignIn" class="btn btn-secondary" onclick="signIn()">
        Sign In
      </button>
    </div>
  </nav>
  <br>
  <h5 class="card-header text-center">Vanilla JavaScript SPA calling MS Graph API with MSAL.js</h5>
  <br>
  <div class="row" style="margin:auto">
    <div id="card-div" class="col-md-3" style="display:none">
      <div class="card text-center">
        <div class="card-body">
          <h5 class="card-title" id="WelcomeMessage">Please sign-in to see your profile and read your mails</h5>
          <div id="profile-div"></div>
          <br>
          <br>
          <button class="btn btn-primary" id="seeProfile" onclick="seeProfile()">See Profile</button>
          <br>
          <br>
          <button class="btn btn-primary" id="readMail" onclick="readMail()">Read Mails</button>
        </div>
      </div>
    </div>
    <br>
    <br>
    <div class="col-md-4">
      <div class="list-group" id="list-tab" role="tablist">
      </div>
    </div>
    <div class="col-md-5">
      <div class="tab-content" id="nav-tabContent">
      </div>
    </div>
  </div>
  <br>
  <br>

  <!-- importing bootstrap.js and supporting js libraries -->
  <script src="https://code.jquery.com/jquery-3.4.1.slim.min.js"
    integrity="sha384-J6qa4849blE2+poT4WnyKhv5vZF5SrPo0iEjwBvKU7imGFAV0wwj1yYfoRSJoZ+n"
    crossorigin="anonymous"></script>
  <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js"
    integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo"
    crossorigin="anonymous"></script>
  <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/js/bootstrap.min.js"
    integrity="sha384-wfSDF2E50Y2D1uUdj0O3uMBJnjuUD4Ih7YwaYd1iqfktj0Uod8GCExl3Og8ifwB6"
    crossorigin="anonymous"></script>

  <!-- importing app scripts (load order is important) -->
  <script type="text/javascript" src="./authConfig.js"></script>
  <script type="text/javascript" src="./graphConfig.js"></script>
  <script type="text/javascript" src="./ui.js"></script>

  <!-- <script type="text/javascript" src="./authRedirect.js"></script>   -->
  <!-- uncomment the above line and comment the line below if you would like to use the redirect flow -->
  <script type="text/javascript" src="./authPopup.js"></script>
  <script type="text/javascript" src="./graph.js"></script>
</body>

</html>
```
```
vi ui.js
```
```
// Select DOM elements to work with
const welcomeDiv = document.getElementById("WelcomeMessage");
const signInButton = document.getElementById("SignIn");
const cardDiv = document.getElementById("card-div");
const mailButton = document.getElementById("readMail");
const profileButton = document.getElementById("seeProfile");
const profileDiv = document.getElementById("profile-div");

function showWelcomeMessage(username) {
    // Reconfiguring DOM elements
    cardDiv.style.display = 'initial';
    welcomeDiv.innerHTML = `Welcome ${username}`;
    signInButton.setAttribute("onclick", "signOut();");
    signInButton.setAttribute('class', "btn btn-success")
    signInButton.innerHTML = "Sign Out";
}

function updateUI(data, endpoint) {
    console.log('Graph API responded at: ' + new Date().toString());

    if (endpoint === graphConfig.graphMeEndpoint) {
        profileDiv.innerHTML = ''
        const title = document.createElement('p');
        title.innerHTML = "<strong>Title: </strong>" + data.jobTitle;
        const email = document.createElement('p');
        email.innerHTML = "<strong>Mail: </strong>" + data.mail;
        const phone = document.createElement('p');
        phone.innerHTML = "<strong>Phone: </strong>" + data.businessPhones[0];
        const address = document.createElement('p');
        address.innerHTML = "<strong>Location: </strong>" + data.officeLocation;
        profileDiv.appendChild(title);
        profileDiv.appendChild(email);
        profileDiv.appendChild(phone);
        profileDiv.appendChild(address);

    } else if (endpoint === graphConfig.graphMailEndpoint) {
        if (!data.value) {
            alert("You do not have a mailbox!")
        } else if (data.value.length < 1) {
            alert("Your mailbox is empty!")
        } else {
            const tabContent = document.getElementById("nav-tabContent");
            const tabList = document.getElementById("list-tab");
            tabList.innerHTML = ''; // clear tabList at each readMail call

            data.value.map((d, i) => {
                // Keeping it simple
                if (i < 10) {
                    const listItem = document.createElement("a");
                    listItem.setAttribute("class", "list-group-item list-group-item-action")
                    listItem.setAttribute("id", "list" + i + "list")
                    listItem.setAttribute("data-toggle", "list")
                    listItem.setAttribute("href", "#list" + i)
                    listItem.setAttribute("role", "tab")
                    listItem.setAttribute("aria-controls", i)
                    listItem.innerHTML = d.subject;
                    tabList.appendChild(listItem)

                    const contentItem = document.createElement("div");
                    contentItem.setAttribute("class", "tab-pane fade")
                    contentItem.setAttribute("id", "list" + i)
                    contentItem.setAttribute("role", "tabpanel")
                    contentItem.setAttribute("aria-labelledby", "list" + i + "list")
                    contentItem.innerHTML = "<strong> from: " + d.from.emailAddress.address + "</strong><br><br>" + d.bodyPreview + "...";
                    tabContent.appendChild(contentItem);
                }
            });
        }
    }
}
```

# 4. App registrations
> https://learn.microsoft.com/ja-jp/azure/active-directory/develop/scenario-spa-app-registration

Enter "**http://localhost:3000**" in the Redirect URL.



// <button onclick="loginClicked()">Login</button>

var applicationConfig = {
    clientID: '7bfe8490-c99a-44f7-b16e-ea3df3051c3f',
    authority: 'https://login.microsoftonline.com/nikolame.onmicrosoft.com'
};

var userAgentApplication = new Msal.UserAgentApplication(applicationConfig.clientID, applicationConfig.authority, tokenReceivedCallback);

var graphScopes = ["user.read", "calendars.read"];

function getAccessToken()
{
    userAgentApplication.acquireTokenSilent(graphScopes).then(function (accessToken) {
        //AcquireTokenSilent Success
        console.log(accessToken);
        getMe(accessToken)
    }, function (error) {
        userAgentApplication.acquireTokenPopup(graphScopes).then(accessToken => {
            getMe(accessToken);
        }, error => {
            _$login_div.innerHTML = `
            <button onclick="loginClicked()">Login</button>
            `;
        })
    })
}

getAccessToken();


var _$login_div = document.getElementById('login_div');
var _$meetings_div = document.getElementById('meetings_div');


//callback function for redirect flows
function tokenReceivedCallback(errorDesc, token, error, tokenType) {
    // if (token) {
    //     console.log(tokenType + " " + token)
    //     if (tokenType == 'id_token')
    //     {
            
    //     }
    //     else if (tokenType == 'access_token')
    //     {
    //         // getMe(token);
    //     }
        
    // }
    // else {
    //     console.log(error + ":" + errorDesc);
    //     _$login_div.innerHTML = `
    //         <button onclick="loginClicked()">Login</button>
    //         `;
    // }
}

function getGraphJson(token, endpoint)
{
    var headers = new Headers();
    var bearer = "Bearer " + token;
    headers.append("Authorization", bearer);
    var options = {
        method: "GET",
        headers: headers
    };
    var graphEndpoint = "https://graph.microsoft.com/v1.0" + endpoint;

    return fetch(graphEndpoint, options).then(response => {
        if (response.ok)
        {
            return response.json();
        }
        else
        {
            throw new Error('Bad response from graph!');
        }
    }).then(json => {
        return json;
    }) 
}

function getMe(token)
{
    getGraphJson(token, '/me')
        .then(function (data){
            _$login_div.innerHTML = `
            <div id='displayname'></div>
            `;

            _$login_div.querySelector('#displayname').innerHTML = data.displayName

            getAllMeetings(token);
        });
}

function getAllMeetings(token)
{
    var now = new Date();
    var nextWeek = new Date();
    nextWeek.setDate(nextWeek.getDate() + 7);
    getGraphJson(token, '/me/calendarView?startdatetime=' + now.toISOString() + '&enddatetime=' + nextWeek.toISOString()).then(response => {
        let content = 'Meetings:<ul>';
        response.value.forEach(meeting => {
            let body = meeting.body.content;
            let teamsRegex = /(https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[A-Za-z0-9\-\._~:\/\?#\[\]@!$&%'\(\)\*\+,;\=]*)"/;
            let matches = teamsRegex.exec(body);
            let html = ''
            if (matches && matches.length > 0)
            {
                let teamsUrl = matches[0];
                html = '<a target="_blank" href="' + teamsUrl + '">' + meeting.subject + '</a>';
            }
            else
            {
                html = meeting.subject;
            }

            content += '<li>' + html + '</li>'
        })
        content += '</ul>';

        _$meetings_div.innerHTML = content;
    })
}

function loginClicked()
{
    userAgentApplication.loginPopup(graphScopes).then(idToken => {
        getAccessToken();
    });
}
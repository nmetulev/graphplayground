// <button onclick="loginClicked()">Login</button>

var applicationConfig = {
    clientID:   '83c70f65-0adc-49f6-b6c5-97f601a4334f',
    // clientID:   '7bfe8490-c99a-44f7-b16e-ea3df3051c3f',
    authority:  'https://login.microsoftonline.com/common',
    resource:   'https://graph.microsoft.com'
    // authority: 'https://login.microsoftonline.com/nikolame.onmicrosoft.com'
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

async function initWindows()
{
    if (window.Windows)
    {
        let webCore = window.Windows.Security.Authentication.Web.Core;
        let wap = await webCore.WebAuthenticationCoreManager.findAccountProviderAsync('https://login.microsoft.com', applicationConfig.authority);
        let wtr = new webCore.WebTokenRequest(wap, '', applicationConfig.clientID);
        wtr.properties.insert('resource', applicationConfig.resource);

        let wtrr = await webCore.WebAuthenticationCoreManager.requestTokenAsync(wtr);
        if (wtrr.responseStatus == webCore.WebTokenRequestStatus.success)
        {
            let account = wttr.responseData[0].webAccount;
        }
    }

}

let token = 'eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFDNXVuYTBFVUZnVElGOEVsYXh0V2pUYWctSXFQVGE0TlZCdFc5V3JibjhfWWZmTW81ckcxLW1fUWozWTJhNlhNLWJHUFItQW5HT3FIMEFkSFBCd0Z4eUdualFXUlI2Ymp4bUFNcE5sX3lSYnlBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoibmJDd1cxMXczWGtCLXhVYVh3S1JTTGpNSEdRIiwia2lkIjoibmJDd1cxMXczWGtCLXhVYVh3S1JTTGpNSEdRIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9kNmQwNzk5My00NmYwLTQwMGQtODYzYi1mM2U2ZDZkMGFiYTIvIiwiaWF0IjoxNTQ0NTYxNzY3LCJuYmYiOjE1NDQ1NjE3NjcsImV4cCI6MTU0NDU2NTY2NywiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IjQyUmdZTkJPakhUbStzRnVHMnRadk9YQnBxVTJVOFhmUjVTRnZINlF1TWJ2WlBncGkxSUEiLCJhbXIiOlsicHdkIl0sImFwcF9kaXNwbGF5bmFtZSI6IkdyYXBoVGVzdEFwcCIsImFwcGlkIjoiN2JmZTg0OTAtYzk5YS00NGY3LWIxNmUtZWEzZGYzMDUxYzNmIiwiYXBwaWRhY3IiOiIwIiwiZmFtaWx5X25hbWUiOiJUZXN0ZXIiLCJnaXZlbl9uYW1lIjoiVG9vbGtpdCIsImlwYWRkciI6IjEzMS4xMDcuMTc0LjE1MSIsIm5hbWUiOiJUb29sa2l0IFRlc3RlciIsIm9pZCI6ImQzZmM2OGI5LThiODEtNDlmMi04YzgwLWQwMDgyMjhlZDMxZCIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzMDAwMEFDNTdDRjEyIiwic2NwIjoiQ2FsZW5kYXJzLlJlYWQgb3BlbmlkIHByb2ZpbGUgVXNlci5SZWFkIGVtYWlsIiwic3ViIjoiLXItYjNkZWNaYXZubGV0UFF6OE5KbUFZOW1Cdm5IOUVpQ3VHYWo3UlFXNCIsInRpZCI6ImQ2ZDA3OTkzLTQ2ZjAtNDAwZC04NjNiLWYzZTZkNmQwYWJhMiIsInVuaXF1ZV9uYW1lIjoidGVzdGVyQG5pa29sYW1lLm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6InRlc3RlckBuaWtvbGFtZS5vbm1pY3Jvc29mdC5jb20iLCJ1dGkiOiJlblJfbDRBVGxrYTVGaDdjemdCd0FBIiwidmVyIjoiMS4wIiwieG1zX3N0Ijp7InN1YiI6InJPZ0lzbXRDUXljSGZwci1OQkEtbVByUG1FYWVFX0F2TjI3TGVVVHo1WjQifSwieG1zX3RjZHQiOjE1MzEyNDMyNDd9.p4kifw1jxyvbXynb1LmIkagQxrR5DGUDVd3s8XEs8rmUh99yVkefKvGBUlET1oSXhHnG0wHKROFJhHNyBOZZMD1yvZ3ASDuTYmvDmhgkM5ZZyJoCAEkHHyOIFQMcQp-K7da2GHeON_VMjxrRA8fwKC870YcHcBepjubmO4kIujsrbMgbTQ7XWt1qdZyVOq1A9qUJYrC-5i6cQW16vwk_oVKkmVaxIuuTqLOBoHFmQg2avBeIrmxBfBGh5JCmsASsZAvU72qLQo8exki3_86Eo_qPT6QE60VllHzKnIRsCDWFGJeUmD6rwn3GKl58QTfRuh_unnUxMar6E6hE9MWrHQ';

if (token)
{
    getMe(token)
}
else
{
    getAccessToken();

}


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

async function getMe(token)
{
    let data = await getGraphJson(token, '/me')

    _$login_div.innerHTML = `
    <div id='displayname'></div>
    `;

    _$login_div.querySelector('#displayname').innerHTML = data.displayName

    getAllMeetings(token);
}

async function getAllMeetings(token)
{
    var now = new Date();
    var nextWeek = new Date();
    nextWeek.setDate(nextWeek.getDate() + 7);

    let response = await getGraphJson(token, '/me/calendarView?startdatetime=' + now.toISOString() + '&enddatetime=' + nextWeek.toISOString())
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
}

async function loginClicked()
{
    try {
        let idToken = await userAgentApplication.loginPopup(graphScopes);
        getAccessToken();
    } catch (error) {
        
    }
}
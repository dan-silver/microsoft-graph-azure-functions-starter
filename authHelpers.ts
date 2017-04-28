import { MicrosoftAppRefreshToken, MicrosoftAppSecret, MicrosoftAppID } from './secrets'
import { Client } from '@microsoft/microsoft-graph-client'

import fetch from 'node-fetch';
import {Headers} from 'node-fetch';

export async function getAccessToken() {
    let body = {
        client_id: MicrosoftAppID,
        scope: "openid profile User.ReadWrite User.ReadBasic.All Mail.ReadWrite Mail.Send Mail.Send.Shared Calendars.ReadWrite Calendars.ReadWrite.Shared Contacts.ReadWrite Contacts.ReadWrite.Shared MailboxSettings.ReadWrite Files.ReadWrite Files.ReadWrite.All Notes.Create Notes.ReadWrite.All People.Read Sites.ReadWrite.All Tasks.ReadWrite",
        redirect_uri: "http://localhost:7071",
        grant_type: "refresh_token",
        client_secret: MicrosoftAppSecret,
        refresh_token: MicrosoftAppRefreshToken
    };

    let options:any = {
        method: "POST",
        headers: {
        'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
        }
    };

    const searchParams = Object.keys(body).map((key) => {
        return encodeURIComponent(key) + '=' + encodeURIComponent(body[key]);
    }).join('&');

    options.body = searchParams;

    return fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", options).then((rawResponse) => {
        return rawResponse.json()
    }).then((json) => {
        return json["access_token"]
    })
}

export async function GraphClient() {
    let accessToken = await getAccessToken();

    return Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
}
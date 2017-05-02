import { MicrosoftAppSecret, MicrosoftAppID, TenantDomain } from './secrets'
import { Client } from '@microsoft/microsoft-graph-client'

import fetch from 'node-fetch';

export async function getAccessToken() {
    let body = {
        client_id: MicrosoftAppID,
        grant_type: 'client_credentials',
        resource: 'https://graph.microsoft.com',
        client_secret: MicrosoftAppSecret
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

    return fetch(`https://login.windows.net/${TenantDomain}/oauth2/token`, options).then((rawResponse) => {
        return rawResponse.json()
    }).then((json) => {
        console.log(json);
        return json["access_token"];
    }).catch((e) => {
        debugger;
        console.log(e);
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
"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const secrets_1 = require("./secrets");
const microsoft_graph_client_1 = require("@microsoft/microsoft-graph-client");
const node_fetch_1 = require("node-fetch");
function getAccessToken() {
    return __awaiter(this, void 0, void 0, function* () {
        let body = {
            client_id: secrets_1.MicrosoftAppID,
            scope: "openid profile User.ReadWrite User.ReadBasic.All Mail.ReadWrite Mail.Send Mail.Send.Shared Calendars.ReadWrite Calendars.ReadWrite.Shared Contacts.ReadWrite Contacts.ReadWrite.Shared MailboxSettings.ReadWrite Files.ReadWrite Files.ReadWrite.All Notes.Create Notes.ReadWrite.All People.Read Sites.ReadWrite.All Tasks.ReadWrite",
            redirect_uri: "http://localhost:7071",
            grant_type: "refresh_token",
            client_secret: secrets_1.MicrosoftAppSecret,
            refresh_token: secrets_1.MicrosoftAppRefreshToken
        };
        let options = {
            method: "POST",
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
            }
        };
        const searchParams = Object.keys(body).map((key) => {
            return encodeURIComponent(key) + '=' + encodeURIComponent(body[key]);
        }).join('&');
        options.body = searchParams;
        return node_fetch_1.default("https://login.microsoftonline.com/common/oauth2/v2.0/token", options).then((rawResponse) => {
            return rawResponse.json();
        }).then((json) => {
            return json["access_token"];
        });
    });
}
exports.getAccessToken = getAccessToken;
function GraphClient() {
    return __awaiter(this, void 0, void 0, function* () {
        let accessToken = yield getAccessToken();
        return microsoft_graph_client_1.Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });
    });
}
exports.GraphClient = GraphClient;

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
            grant_type: 'client_credentials',
            resource: 'https://graph.microsoft.com',
            client_secret: secrets_1.MicrosoftAppSecret
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
        return node_fetch_1.default(`https://login.windows.net/${secrets_1.TenantDomain}/oauth2/token`, options).then((rawResponse) => {
            return rawResponse.json();
        }).then((json) => {
            console.log(json);
            return json["access_token"];
        }).catch((e) => {
            debugger;
            console.log(e);
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
//# sourceMappingURL=authHelpers.js.map
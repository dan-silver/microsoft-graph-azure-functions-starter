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
const authHelpers_1 = require("./authHelpers");
function getUsersWithExtensions() {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield authHelpers_1.GraphClient();
        return client
            .api("/users")
            .version("beta")
            .expand("extensions")
            .get()
            .then((res) => {
            return res.value;
        });
    });
}
exports.getUsersWithExtensions = getUsersWithExtensions;
function sendMailReport(users) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield authHelpers_1.GraphClient();
        let emailString = '';
        for (let user of users) {
            if (user.extensions)
                emailString += "<tr><td>" + user.displayName + "</td><td>" + user.extensions[0].numEvents + " events next week</td></tr>";
        }
        let message = {
            subject: "Report on employee calendars",
            toRecipients: [{
                    emailAddress: {
                        address: "dansil@microsoft.com"
                    }
                }],
            body: {
                content: `
                <table>

                ${emailString}
                </table>
            `,
                contentType: "html"
            }
        };
        return yield client
            .api("/users/admin@MOD789932.onmicrosoft.com/sendMail")
            .post({ message })
            .then((res) => {
            console.log("Mail sent!");
        }).catch((error) => {
            debugger;
        });
    });
}
exports.sendMailReport = sendMailReport;
function sortUsersOnNumCalEvents(a, b) {
    if (!a.extensions || !b.extensions)
        return 0;
    if (a.extensions[0].numEvents > b.extensions[0].numEvents)
        return -1;
    if (a.extensions[0].numEvents < b.extensions[0].numEvents)
        return 1;
    return 0;
}
exports.sortUsersOnNumCalEvents = sortUsersOnNumCalEvents;
function removeAllExtensionsOnUser(user) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield authHelpers_1.GraphClient();
        return client
            .api(`/users/${user.id}/extensions`)
            .version(`beta`)
            .get()
            .catch((e) => {
            debugger;
        }).then((res) => {
            let extensionIds = res['value'].map((extension) => extension.id);
            let extensionRemovals = [];
            for (let id of extensionIds) {
                extensionRemovals.push(client
                    .api(`/users/${user.id}/extensions/${id}`)
                    .version(`beta`)
                    .delete()
                    .catch((e) => {
                    debugger;
                }));
            }
            return Promise.all(extensionRemovals);
        });
    });
}
exports.removeAllExtensionsOnUser = removeAllExtensionsOnUser;
function queryNumberCalendarEvents(user) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield authHelpers_1.GraphClient();
        let today = new Date();
        let inOneMonth = new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000);
        return client
            .api(`/users/${user.mail}/calendarview/`)
            .query({
            startdatetime: today.toISOString(),
            enddatetime: inOneMonth.toISOString()
        })
            .get()
            .then((res) => {
            return res.value.length;
        })
            .catch((e) => {
            console.log(e);
            debugger;
            return -1;
        });
    });
}
exports.queryNumberCalendarEvents = queryNumberCalendarEvents;
function saveUserExtension(user, calendarEventsCount) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield authHelpers_1.GraphClient();
        let extensionData = {
            extensionName: "numCalendarEvents",
            numEvents: calendarEventsCount
        };
        return client
            .api(`/users/${user.id}/extensions`)
            .version(`beta`)
            .post(extensionData)
            .catch((e) => {
            debugger;
        });
    });
}
exports.saveUserExtension = saveUserExtension;
//# sourceMappingURL=graph-helpers.js.map
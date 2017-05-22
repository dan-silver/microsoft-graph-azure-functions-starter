
import { GraphClient } from "./authHelpers";
import { User, Message } from "@microsoft/microsoft-graph-types/microsoft-graph";

export async function getUsersWithExtensions(): Promise<User[]> {
    const client = await GraphClient();

    return client
        .api("/users")
        .version("beta")
        .expand("extensions")
        .get()
        .then((res) => {
            return res.value;
        })
}


export async function sendMailReport(users) {
    const client = await GraphClient();

    let emailString = ''

    for (let user of users) {
        if (user.extensions)
            emailString += "<tr><td>" + user.displayName + "</td><td>" + user.extensions[0].numEvents + " events next week</td></tr>"
    }

    let message:Message = {
        subject: "Report on employee calendars",
        toRecipients: [{
            emailAddress: {
                address: "example@example.com"
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
    }
    return await client
        .api("/users/admin@MOD789932.onmicrosoft.com/sendMail")
        .post({message})
        .then((res) => {
            console.log("Mail sent!")
        }).catch((error) => {
            debugger;
        });
}



export function sortUsersOnNumCalEvents(a,b) {
  if (!a.extensions || !b.extensions) return 0;
  if (a.extensions[0].numEvents > b.extensions[0].numEvents)
    return -1;
  if (a.extensions[0].numEvents < b.extensions[0].numEvents)
    return 1;
  return 0;
}



export async function removeAllExtensionsOnUser(user:User) {
    const client = await GraphClient();

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
                extensionRemovals.push(
                        client
                        .api(`/users/${user.id}/extensions/${id}`)
                        .version(`beta`)
                        .delete()
                        .catch((e) => {
                            debugger;
                        }));
            }
            return Promise.all(extensionRemovals);
        })
}



export async function queryNumberCalendarEvents(user:User) {
    const client = await GraphClient();

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
        })
}


export async function saveUserExtension(user:User, calendarEventsCount:number) {
    const client = await GraphClient();

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
        })
}

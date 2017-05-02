import { GraphClient } from '../authHelpers';
import { User } from '@microsoft/microsoft-graph-types' 

export async function main (context, req) {
    context.log("Starting Azure function!");

    let emails = await getEmails();
    let response = {
        status: 200,
        body: {
            emails
        }
    };
    return response;
};

async function getEmails() {
    const client = await GraphClient();

    return client
        .api("/users")
        .get()
        .then((res) => {
            let users:User[] = res.value;
            return users.map((user) => user.mail)
        });
}
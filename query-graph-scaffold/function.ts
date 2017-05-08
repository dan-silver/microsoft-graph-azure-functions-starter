import { GraphClient } from '../authHelpers';
import { User } from '@microsoft/microsoft-graph-types' 

export async function main (context?, req?) {
    if (context) context.log("Starting Azure function!");

    let users = await getAllUsers();
    let response = {
        status: 200,
        body: {
            users
        }
    };
    return response;
};

async function getAllUsers() {
    const client = await GraphClient();

    return client
        .api("/users")
        .get()
        .then((res) => {
            let users:User[] = res.value;
            return users;
        });
}

// to test this locally in Visual Studio code, uncomment the following line
// when running in Azure functions, they'll call main() for you

// main();
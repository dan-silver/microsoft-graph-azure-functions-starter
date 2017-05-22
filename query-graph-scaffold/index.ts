import { GraphClient } from '../authHelpers';
import { User } from '@microsoft/microsoft-graph-types' 


export async function index(context?, req?) {
    if (context) context.log("Starting Azure function!");

    let users = await getAllUsers(context);
    let response = {
        status: 200,
        body: {"result":`Found ${users.length} users`}
    };
    return response;
};

async function getAllUsers(context) {
    const client = await GraphClient(context);
    return client
        .api("/users")
        .get()
        .then((res) => {
            context.log(`Found ${res.value.length} users`);
            let users:User[] = res.value;
            return users;
        });
}

// to test this locally in Visual Studio code, uncomment the following line
// when running in Azure functions, they'll call main() for you

// index();
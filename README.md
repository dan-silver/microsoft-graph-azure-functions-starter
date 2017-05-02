## Microsoft Graph and Azure Functions TypeScript Starter Project
This project contains a starting point for developers to build on the Microsoft Graph using Azure Functions.

### Setup
Clone this repo with `git clone git@github.com:dan-silver/microsoft-graph-azure-functions-starter.git`
* Install Node.js
* ```npm install``` to install project dependencies
* ``` npm run build ``` to start TypeScript compiler and watch for changes (leave this open in another console)

#### App registration

* Register your application at https://apps.dev.microsoft.com/
* Add `http://localhost:3000` as a redirect URL under Platforms -> web
* Add Application permissions (Directory.ReadWrite.All)
* Update `secrets.ts` with your application id
* Update `secrets.ts` with an app secret (click 'generate new password')
* Update `secrets.ts` with your tenant domain like `MOD507192.onmicrosoft.com`

#### Allow your app running as a service to access Graph

* Visit `https://login.microsoftonline.com/common/adminconsent?client_id=YOUR_APP_ID&state=12345&redirect_uri=http://localhost:3000` and grant the app access. Replace `YOUR_APP_ID`

#### Deploy to Azure Functions

* Create an Azure function app at https://portal.azure.com and create a new 'Function App'
![create function app](screenshots/create-function-app.png)
* Set up deployment using Git (…)
* Push a deployment (git …)
* Check the logs (go to site: )
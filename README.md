## Microsoft Graph and Azure Functions TypeScript Starter Project
This project contains a starting point for developers to deploy Microsoft Graph service apps in Azure Functions.

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
* Configure your app to allow Git deployments
   * Under `Platform features` select `Deployment Options` and choose 'Local Git Repository as the source.
* Add your Function as a git remote
   * Under General Settings -> Properties, select your Git Url
   * `git remote add azure YOUR_GIT_URL`
* Push a deployment
  * (make sure TypeScript recompiled after you updated your secrets file with npm run build)
  * `git add .`, `git commit -m 'Initial commit'`, `git push azure master` 
* Check the logs and your function should be running!

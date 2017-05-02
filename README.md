## Microsoft Graph and Azure Functions Node.js Starter Project

This project contains a starting point for developers to build on the Microsoft Graph using Azure Functions.

### Setup
* Create an Azure function app at https://portal.azure.com and create a new 'Function App'


![create function app](screenshots/create-function-app.png)

* Register your application at https://apps.dev.microsoft.com/
* Set up deployment using Git (…)
* Clone this repo with `git clone git@github.com:dan-silver/microsoft-graph-azure-functions-starter.git`
  * Install Node.js
  * ```npm install``` to install project dependencies
  * ``` tsc -w ``` to start TypeScript compiler and watch for changes
* Push a deployment (git …)
* Check the logs (go to site: )

### Install
* Obtain an App ID, App Secret and refresh token and update appsettings.json (Windows) for local testing or Azure app settings.


### Testing locally with Azure functions client (Windows only)
https://www.npmjs.com/package/azure-functions-cli

* Enter secrets in appsettings.json (copy from example at `appsettings.example.json`). This is an easy way to setup environmental variables without modifying system settings.
* ```npm i -g azure-functions-cli```
* ``` func host start```
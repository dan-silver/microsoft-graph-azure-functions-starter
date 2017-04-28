## Microsoft Graph and Azure Functions Node.js Starter Project

This project contains a starting point for developers to build on the Microsoft Graph using Azure Functions.

### Setup
* Fork repo this repo to your GitHub account

### Register your app
* Visit https://apps.dev.microsoft.com/ and register a new application
 * Enter 


### Get started in Azure
* Visit https://portal.azure.com and create a new 'Function App'



### Install
* Install Node.js
* ```npm install``` to install Node dependencies declared in package.json
* ``` tsc -w ``` to start TypeScript compiler and watch for changes
* Obtain an App ID, App Secret and refresh token and update appsettings.json (Windows) for local testing or Azure app settings.


### Testing locally with Azure functions client (Windows only)
https://www.npmjs.com/package/azure-functions-cli

* Enter secrets in appsettings.json (copy from example at `appsettings.example.json`). This is an easy way to setup environmental variables without modifying system settings.
* ```npm i -g azure-functions-cli```
* ``` func host start```


### Deploying to Azure
* Create a GitHub repository for your project and push code
* Authorize Azure to access your GitHub account if you haven't already
* Create an Azure Functions app from source control
# Summary

This sample is a action based messaging extension created using the Teams Yeoman Generator as featured in this video. It authenticates against Microsoft Graph via SSO and on-behalf flow from the frontend and recieves documents to be posted as adaptive card to the message compose box of a channel.
The essential thing is that it is configurable and persists t's configuration to an Azure App Configuration resource.

![Configuration of Teams Messaging extension ...](https://mmsharepoint.files.wordpress.com/2021/05/08fetchtask_configurefirst.png?w=946)
For further details see the author's [blog post](https://mmsharepoint.wordpress.com/)

## Prerequisites

* [Office 365 tenant](https://dev.office.com/sharepoint/docs/spfx/set-up-your-development-environment)
* [Node.js](https://nodejs.org) version 10.14.1 or higher
* [ngrok](https://ngrok.com) or similar tunneling application is required for local testing
* [Azure App Configuration](https://docs.microsoft.com/en-us/azure/azure-app-configuration/quickstart-aspnet-core-app?tabs=core5x#create-an-app-configuration-store) setup like mentioned here

    ```bash
    # determine node version
    node --version
    ```
## Version history

Version|Date|Author|Comments
-------|----|----|--------
1.0|May 14, 2021|[Markus Moeller](https://twitter.com/moeller2_0)|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## Minimal Path to Awesome
- Clone the repository

    ```bash
    git clone https://github.com/PnP/teams-dev-samples.git
    ```

- In a console, navigate to `samples/msgext-customer-search`

    ```bash
    cd samples/msgext-customer-search
    ```

- Install modules

    ```bash
    npm install
    ```
- Run ngrok and note down the given url

    ```bash
    gulp start-ngrok
    ```
- Since messaging extensions utilize the Azure Bot Framework, you will need to register a new bot. 
[These instructions](https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/create-a-bot-for-teams#register-your-web-service-with-the-bot-framework) provide options for registering with or without an Azure subscription. 
  - Be sure to enable the Microsoft Teams bot channel so your solution can communicate with Microsoft Teams
  - For local testing, set the messaging endpoint to the https URL returned by ngrok plus "/api/messages"
  - Note the bot's Application ID and password (also called the Client Secret) assigned to your bot during the registration process. In the Azure portal this is under the Bot Registration settings; in the legacy portal it's in the Settings tab. Click Manage to go to Azure AD to obtain the Client Secret. You may need to create a new Application Secret in order to have an opportunity to copy it out of the Azure portal. 
- Update the `.env` configuration for the bot to use the Microsoft App Id and App Password (aka Client Secret) from the previous step.
- You will furthermore need to register an app in Azure AD [also described here](https://mmsharepoint.wordpress.com/2020/07/03/a-microsoft-teams-messaging-extension-with-authentication-and-access-to-microsoft-graph-i/)
  - with client secret
  - with **delegated** permissions email, offline_access, openid, profile, Sites.Read.All
  - With exposed Api "access_as_user" and App ID Uri api://<NGrok-Url>/<App ID>
  - With the client IDs for Teams App and Teams Web App 1fec8e78-bce4-4aaf-ab1b-5451cc387264 and 5e3ce6c0-2b1f-4285-8d4b-75ee78787346
- Also add the app ID and its secret to .env as GRAPH_APP_ID= and GRAPH_APP_SECRET=
- Add your configured App Configuration Http endpoint to .env as AZURE_CONFIG_CONNECTION_STRING
- Register another app and secret and insert it to your .env as AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET and grant access to the Azure App Configuration if you need to debug locally
- Generate and upload the application package
  ```bash
  gulp manifest
  ```
  Upload the resulting zip file into Teams [using these instructions](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/apps-upload).

- Build the app
``` bash
npm i -g gulp gulp-cli
gulp build
```


## Configuration

Configuration is stored in the `.env` file. 

## Debug and test locally

To debug and test the solution locally you use the `serve` Gulp task. This will first build the app and then start a local web server on port 3007, where you can test your Tabs, Bots or other extensions. Also this command will rebuild the App if you change any file in the `/src` directory.

``` bash
gulp serve
```

To debug the code you can append the argument `debug` to the `serve` command as follows. This allows you to step through your code using your preferred code editor.

``` bash
gulp serve --debug
```

To step through code in Visual Studio Code the following snippet in the `./.vscode/launch.json` file is added. So you can easily attach to the node process after running the `gulp serve --debug` command.

``` json
{
    "type": "node",
            "request": "attach",
            "name": "Attach Node",
            "port": 9229,
            "sourceMaps": true,
            "outFiles": [
                "${workspaceRoot}/dist/**/*.js"
            ]
},
```

### Additional build options

You can use the following flags for the `serve`, `ngrok-serve` and build commands:

* `--no-linting` - skips the linting of Typescript during build to improve build times
* `--debug` - builds in debug mode

## Deployment

The solution can be deployed to Azure using any deployment method.

* For Azure Devops see [How to deploy a Yo Teams generated project to Azure through Azure DevOps](https://www.wictorwilen.se/blog/deploying-yo-teams-and-node-apps/)
* For Docker containers, see the included `Dockerfile`

## Logging

To enable logging for the solution you need to add `msteams` to the `DEBUG` environment variable. See the [debug package](https://www.npmjs.com/package/debug) for more information. By default this setting is turned on in the `.env` file.

Example for Windows command line:

``` bash
SET DEBUG=msteams
```

If you are using Microsoft Azure to host your Microsoft Teams app, then you can add `DEBUG` as an Application Setting with the value of `msteams`.

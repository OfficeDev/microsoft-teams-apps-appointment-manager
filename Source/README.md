# Setup and Debugging

## Overview
The Virtual Consult app for Microsoft Teams requires the following setup steps:

1. [Register bot](#regBot)
1. [Create app registration](#createApp)
1. [Clone the application](#cloneApp)
1. [Update app manifests/packages](#updateManifest)
1. [Provision Cosmos DB](#createCosmos)
1. [Provision Application Insights](#createAppInsights)
1. [Update appsettings.json](#updateSettings)
1. [Run/debug the app](#runDebug)
1. [Important Notes](#notes)

## <a name="regBot"></a>Register bot
You can also register your web service by creating a Bot Channels Registration resource in the Azure portal.

1. In the Azure portal, under Azure services, select **Create a resource**.

1. In the search box enter "bot". And in the drop-down list, select **Bot Channels Registration**.

1. Select the **Create** button.

1. In the **Bot Channel Registration** blade, provide the requested information about your bot.

1. Leave the **Messaging endpoint** box empty for now, you will enter the required URL after deploying the bot. 

1. Click **Microsoft App ID and password** and then **Create New**.

1. Click **Create App ID in the App Registration Portal** link.

1. In the displayed App registration window, click the **New registration** tab in the upper left.

1. Enter the name of the bot application you are registering, we used VirtualConsult (you need to select your own unique name).

1. For the **Supported account types** select Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox).

1. Click the **Register** button. Once completed, Azure displays the **Overview** page for the application.

1. Copy and save to a file the **Application (client) ID** value. You’ll need it later when updating your Teams application manifest and application settings.

1. In the left panel, click **Certificate and secrets** under **Manage**.

    a. In the **Client secrets** section, click **New client secret**.
    
    b. Add a description to identify this secret from others you might need to create for this app.

    c. **Set Expires** to your selection.

    d. Click **Add**.

    e. Copy the **client secret** and save it to a file. You’ll need it later when updating your Teams application manifest and application settings.

1. Go back to the **Bot Channel Registration** window and copy the **App ID** and the **Client secret** in the **Microsoft App ID** and **Password** boxes, respectively.

1. Click **OK**.

1. Finally, click **Create**.

Once your bot channels registration is created, you'll need to enable the Teams channel.

- In the Azure portal, under Azure services, select the Bot Channel Registration you just created.

- In the left panel, click **Channels**.

- Click the **Microsoft Teams** icon, then choose **Save**.

The Bot Framework portal is optimized for registering bots in Microsoft Azure. Here are some things to know:

- The Microsoft Teams channel for bots registered on Azure is free. Messages sent over the Teams channel will NOT count towards the consumed messages for the bot.

- If you register your bot using Microsoft Azure, your bot code doesn't need to be hosted on Microsoft Azure.

- If you do register a bot using Microsoft Azure portal, you must have a Microsoft Azure account. You can create one for free. To verify your identity when you create an Azure account, you must provide a credit card, but it won't be charged; it's always free to create and use bots with Microsoft Teams.

## <a name="createApp"></a>Create app registration
In addition to the bot's app registration (created in the previous section), the app needs an app registration so the front-end can securely communicate with the back-end and with the Microsoft Graph. 

This Teams application uses [Resource-Specific Consent](https://docs.microsoft.com/en-us/microsoftteams/platform/graph-api/rsc/resource-specific-consent) (RSC), which minimizes the  permissions of the application to only the Teams/Channels where the application is installed. RSC is a Microsoft Teams and Graph API integration that enables your app to use API endpoints to manage specific teams within an organization. The RSC permissions model enables team owners to grant consent for an application to access and/or modify a team's data. The app registration with RSC is slightly more complex, so follow carefully.

1. In the Azure portal, expand the left menu and select **Azure Active Directory** > **Enterprise applications** > **User settings**.

1. Enable, disable, or limit user consent with the control labeled Users can consent to apps accessing company data for the groups they own (This capability is enabled by default).

1. Return to the main Azure Active Directory page and select **App registrations** from the menu.

1. Select **New registration** and on the register an application page, set following values:

    - Set **name** to your app name.
    - Choose the **supported account types** (any account type will work) ¹
    - Leave **Redirect URI** empty.
    - Choose **Register**.

1. On the overview page, copy and save the **Application (client) ID**. You’ll need it later when updating your Teams application manifest and application settings.

1. Choose **Expose an API** under **Manage** from the left nav bar.

1. Select the **Set** link to generate the **Application ID URI** in the form of `api://{AppID}`. Insert your fully qualified domain name (with a forward slash "/" appended to the end) between the double forward slashes and the GUID. The entire ID should have the form of: `api://fully-qualified-domain-name.com/{AppID}`

    - Ex: `api://subdomain.example.com/00000000-0000-0000-0000-000000000000`

    > The fully qualified domain name is the human readable domain name from which your app is served. If you are using a tunneling service such as ngrok, you will need to update this value whenever your ngrok subdomain changes.

1. Select the **Add a scope** button in the **Scopes defined by this API** section.

1. **Add a scope** with the following information:

    - **Scope name**: access_as_user
    - **Who can consent?**: Admins and users
    - **Admin consent display name**: Teams can access the user’s profile.
    - **Admin consent description**: Allows Teams to call the app’s web APIs as the current user.
    - **User consent display name**: Teams can access the user profile and make requests on the user's behalf.
    - **User consent description**: Teams can access the user profile and make requests on the user's behalf.
    - **State**: Enabled

1. Select the **Add scope** button to save.

    - The domain part of the Scope name displayed just below the text field should automatically match the Application ID URI set in the previous step, with `/access_as_user` appended to the end:

        - `api://subdomain.example.com/00000000-0000-0000-0000-000000000000/access_as_user`

1. In the **Authorized client applications** section, identify the applications that you want to authorize for your app’s web application. Select Add a client application. Enter each of the following client IDs and select the authorized scope you created in the previous step:

    - 1fec8e78-bce4-4aaf-ab1b-5451cc387264 (Teams mobile/desktop application)
    - 5e3ce6c0-2b1f-4285-8d4b-75ee78787346 (Teams web application)

1. Navigate to **API Permissions**. Select **Add a permission** > **Microsoft Graph** > **Delegated permissions**, then add the following permissions:

    - BookingsAppointment.ReadWrite.All
    - Calendars.Read
    - Calendars.Read.Shared
    - email
    - offline_access
    - openid
    - profile
    - User.ReadBasic.All

1. Select **Add a permission** > **My APIs** and select your application and add the delegated  **access_as_user** permission we created earlier:

1. Although the application is not requesting elevated permission that would require admin consent, the single sign-on (SSO) experience in Teams will be better if a global admin for the organization consents for users in advance. This can be done in two ways:

    - From this **API Permissions** page for the application in the app registration portal
    - Using the URI https://login.microsoftonline.com/{tenant-id}/adminconsent?client_id={client-id} 

    > If an app hasn't been granted IT admin consent, users will have to provide consent the first time they use an app.

1. Navigate to **Authentication** and set a redirect URI:

    - Select **Add a platform**.
    - Select **Web**.
    - Enter the **Redirect URI** for your app. This will be the page where a successful implicit grant flow will redirect the user. This will be same fully qualified domain name that you entered in step 7 followed by the API route where a authentication response should be sent. This application uses the **/auth** route, so the full redirect URI would be: https://subdomain.example.com/auth

1. Next, enable implicit grant by checking the following boxes:
    - ID Token
    - Access Token

1. Choose **Certificates & secrets** under **Manage** from the left nav bar.

1. In the left panel, click **Certificate and secrets** under **Manage**.

    a. In the **Client secrets** section, click **New client secret**.
    
    b. Add a description to identify this secret from others you might need to create for this app.

    c. **Set Expires** to your selection.

    d. Click **Add**.

    e. Copy the **client secret** and save it to a file. You’ll need it later when updating your Teams application manifest and application settings.

## <a name="cloneApp"></a>Clone the application

1. Clone the repository locally using the `git clone <REPO_URI>` command.

1. Restore client-side packages by running `npm install` in the **Source** folder of the repository.

1. Restore .NET Core packages by running `dotnet restore` in the **Source** folder of the repository. 

## <a name="updateManifest"></a>Update app manifests/packages

This application includes two Microsoft Teams applications. One for administrators to configure the application and one for users/agents that will triage consultation requests. The manifests for these two applications are located in the **Manifests** folder in the root of the repository. 

1. Open the **Manifests/AdminApp/manifest.json** file and update the following sections:

    - **id**: any GUID.
    - **developer**: update this entire section specific to your organization.
    - **staticTabs**: update this section so the **contentUrl** points to your web hosting location or tunneling service such as ngrok.
    - **validDomains**: update this section with the host domain of the application. If using tunneling service such as ngrok, use that full domain (ex: *subdomain*.ngrok.io)
    - **webApplicationInfo**: update this section with the **Application (client) ID** and **Application ID URI** you created in the [Create app registration](#createApp) section.


1. Open the **Manifests/AgentApp/manifest.json** file and update the following sections:

    - **id**: any GUID.
    - **developer**: update this entire section specific to your organization.
    - **configurableTabs**: update this section so the **configurationUrl** points to your web hosting location or tunneling service such as ngrok.
    - **staticTabs**: update this section so the **contentUrl** points to your web hosting location or tunneling service such as ngrok.
    - **bots**: update the **botId** in this section with the **Application (client) ID** you created in the [Register bot](#regBot) section.
    - **composeExtensions**: update the **botId** in this section with the **Application (client) ID** you created in the [Register bot](#regBot) section.
    - **validDomains**: update this section with the host domain of the application. If using tunneling service such as ngrok, use that full domain (ex: *subdomain*.ngrok.io)
    - **webApplicationInfo**: update this section with the **Application (client) ID** and **Application ID URI** you created in the [Create app registration](#createApp) section.

1. After updating these manifests, you can package the two apps separately by creating a .zip file for each app with the following contents:
    - manifest.json
    - color.png
    - outline.png

## <a name="createCosmos"></a>Provision Cosmos DB
See [infrastructure guide](../Infrastructure/README.md)

## <a name="createAppInsights"></a>Provision Application Insights
See [infrastructure guide](../Infrastructure/README.md)

## <a name="updateSettings"></a>Update appsettings.json
1. Create a copy of the **appsettings.json.template** file in the **Source** folder of the repo and rename it to **appsettings.json**.

1. Open the **appsettings.json** file and make the following updates:

    - **ApplicationInsights/InstrumentationKey**: the **Instrumentation Key** from the [Provision Application Insights](#createAppInsights) section.
    - **Bot/Id**: the **Application (client) ID** from the [Register bot](#regBot) section
    - **Bot/Password**: the **client secret** from the [Register bot](#regBot) section
    - **CosmosDb/ConnectionString**: the **Connection String** from the [Provision Cosmos DB](#createCosmos) section
    - **CosmosDb/DatabaseName**: the **Database Name** from the [Provision Cosmos DB](#createCosmos) section
    - **AzureAD/TenantId**: 
    - **AzureAD/AppId**: the **Application (client) ID** from the [Create app registration](#createApp) section
    - **AzureAD/AppPassword**: the **client secret** from the [Create app registration](#createApp) section
    - **AzureAD/HostDomain**: the **fully qualified domain name** part of the **Application ID URI** from the [Create app registration](#createApp) section (ex: subdomain.example.com)

## <a name="runDebug"></a>Run/debug the app

1. Build the front-end project by running `npm run build-dev` from the root of the Source folder. If you plan on making changes for development, you can instead run `npm run watch` to automatically rebuild when files change.

1. Run the .NET solution.
    - If using Visual Studio, you can open the `.sln` file in the root and start debugging as normal (F5).
    - If running from the command line, run `dotnet run` from the `Source` folder.

1. To debug the front-end project in VS Code, set breakpoints and use one of the provided launch configurations:

    > Note: These configurations assume that you've opened the `Source` folder in VS Code. If you've opened the root of the repository or a different folder, open [`launch.json`](.vscode/launch.json) and modify `webRoot` accordingly in the front-end launch configurations.

    - `Launch Chrome` will launch a new instance of Chrome and attach the debugger. Then you can navigate to `teams.microsoft.com` and use the app as usual.
    - `Attach to Chrome` will attach the debugger to an already running instance of Chrome. You must first launch Chrome with remote debugging enabled:

        ```
        chrome.exe --remote-debugging-port=9222
        ```

        If you have other instances of Chrome running, you may have to set the `user-data-dir` flag to ensure Chrome launches a new instance:

        ```
        chrome.exe --remote-debugging-port=9222 --user-data-dir=/path/to/some/temp/directory
        ```

        Then you can navigate to `teams.microsoft.com` and use the app as usual.
    
    If you want to use Edge (Chromium) instead, install the [Debugger for Microsoft Edge extension in VS Code](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) and set up a launch configuration for `edge` instead of `chrome`. (This *should* work but has not been tested.)

TODO: steps for creating a team/channels and sideloading the apps

## <a name="notes"></a>Important Notes

- You need to install the user/agent application into Microsoft Teams before teams/channels will show up in the admin application for routing. This is because Resource-Specific Consent must occur for the admin app to query these secure resources.

- If you plan to debug the bot, you will need to tunnel port 5000 to the internet. NGROK is a popular method for accomplishing this, but any tunneling tool that will get localhost on the internet will work. You can read more about using NGROK in [Run and debug your Microsoft Teams app](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/debug)

- If you make changes to the client-side files, you should rebuild the client-side project by re-running `npm run build-dev` from the `Source` folder. Alternatively, run `npm run watch` to automatically rebuild when files change.
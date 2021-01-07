# Debug and Run Application Locally
This guide is for debugging the application and running the application locally. Deployment of resources like database, bot registration, Azure Active Directory application registration needs to be done before running the steps below. See the [Deployment Guide](deployment-guide) doc for details.  

# <a name="updateSettings"></a>Update appsettings.json
1. Create a copy of the **appsettings.json.template** file in the **Source** folder of the repo and rename it to **appsettings.json**.

1. Open the **appsettings.json** file and make the following updates:

    - **ApplicationInsights/InstrumentationKey**: the **Instrumentation Key** from the [Deploy to your Azure subscription section in Deployment Guide](deployment-guide) doc
    - **Bot/Id**: the **Application (client) ID** 
    - **Bot/Password**: the **client secret** 
    - **CosmosDb/ConnectionString**: the **Connection String** 
    - **CosmosDb/DatabaseName**: the **Database Name** 
    - **AzureAD/TenantId**: your **Tenant ID** from the user-sign in section at [Register Azure Active Directory Applications section in Deployment Guide](deployment-guide) doc
    - **AzureAD/AppId**: the **Application (client) ID** from the user-sign in section at [Register Azure Active Directory Applications section in Deployment Guide](deployment-guide) doc
    - **AzureAD/AppPassword**: the **client secret** from the user-sign in section at [Register Azure Active Directory Applications section in Deployment Guide](deployment-guide) doc
    - **AzureAD/HostDomain**: the **fully qualified domain name** part of the **Application ID URI** from the [Register Azure Active Directory Applicationssection in Deployment Guide](deployment-guide) doc (ex: subdomain.example.com) 


# <a name="runDebug"></a>Run or debug the application

Note: You must already deploy the resources on Azure before following these steps. See the [Deployment Guide](deployment-guide) doc for details.

1. To debug locally, build the front-end project by running `npm run build-dev` or `npm run build` from the root of the Source folder. If you plan on making changes for development, you can instead run `npm run watch` to automatically rebuild when files change.

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
1.  If you want to debug the bot, you will need to tunnel port 5000 to the internet. NGROK is a popular method for accomplishing this, but any tunneling tool that will get localhost on the internet will work. You can read more about using NGROK in [Run and debug your Microsoft Teams app](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/debug)

1. If you make changes to the client-side files, you should rebuild the client-side project by re-running `npm run build-dev` from the `Source` folder. Alternatively, run `npm run watch` to automatically rebuild when files change.
    If you want to use Edge (Chromium) instead, install the [Debugger for Microsoft Edge extension in VS Code](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) and set up a launch configuration for `edge` instead of `chrome`. (This *should* work but has not been tested.)

## Important notes: 
1. To debug database locally, use the local cosmos emulator.
1. To debug bot locally, ngrok URL should be used instead of the bot endpoint.
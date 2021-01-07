# Troubleshooting

## 1. Forgetting the botId or appDomain

If you forgot the your botId and appDomain values from the end of the deployment. You can find them at the ["App registrations" blade](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps)

botId: This is the Microsoft Application ID for the bot. It can be found in the "MicrosoftAppId" field of your configuration e.g. 5630f8a2-c2a0-4cda-bdfa-c2fa87654321.

appDomain: This is the base domain for the Appointment Manager app. It is the value in the "AzureAd:ApplicationIdURI" field of your configuration without the "api://" e.g. appName.azurefd.net.

## 2. Mixing up authentication application Id and bot application Id

Auth application Id - In the application list, click on the app you previously created for user sign-in.

Bot application Id - In the bot channels registration page, click on Settings(under the Bot management section), Microsoft App ID

## 3. Channel mappings not showing
 You need to install the user/staff member application into Microsoft Teams first. Then run the application and the channels will show up in the admin application for routing. This is because Resource-Specific Consent must occur for the admin app to query these secure resources.

# Didn't find your problem here?
Please report the issue [here](<INSERT THE LINK TO THE GITHUB ISSUES PAGE>)

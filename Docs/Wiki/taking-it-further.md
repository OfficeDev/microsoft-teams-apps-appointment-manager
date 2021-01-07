# Taking it Further

## Authentication
**Adding authentication checks to task modules** : Add basic authentication check for administrator. For a more advanced scenario, [Active Directory security groups](https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/groups-self-service-management#self-service-group-management-scenarios) can be set up for better flexibility and management. On invoking task module, a Graph call is made to check security group membership.

**Adding authentication to messenging extension** : Authenticate users via Messenging Extension 

The messaging extension action can be configured with [dynamic input](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/messaging-extension-v3/create-extensions?tabs=typescript#dynamic-input-using-a-web-view) using webview and task module "fetchTask continue" pointing to authentication Uri.
A bot is needed to verify auth sucess and can be done by configuring [task module deeplink](https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/what-are-task-modules#task-module-deep-link-syntax) where a completionbotId is specified. The result is sent as a task/submit invoke back to the bot. The bot then post a message back to chat telling verification complete.

For complete flow, see deep link url [scenario here](https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/what-are-task-modules#overview-of-invoking-and-dismissing-task-modules).

## Change how appointment requests is made
Appointment Manager takes request from a web front end. You may change this behaviour to take requests from other services(JIRA, Emails, other web front end etc) as well - by consuming the Appointment Manager backend API.

## Add more language support
Please refer to the localization guide to add language support. 
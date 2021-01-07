# Known Issues

## 1. Admin consent required
 Granting consent(user or admin) has not implemented in the Appointment Manager app template. As a workaround, admin consent is configured and subsequently provided on authentication setup. Refer deployment guide section "Set up Authentication". Note that none of the Microsoft Graph scopes used by Appointment Manager app template needs admin consent.

User consent could be implemented using the [auth flow supported in Teams](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/authentication/authentication).

## 2. Bookings staff members must be added manually
Adding staff member in Teams will not add them to Bookings. You will have to add them over at the Microsoft 365 portal. Please refer to [this doc](https://docs.microsoft.com/en-us/microsoft-365/bookings/add-staff?view=o365-worldwide) to add staff member individually.

Note: Microsoft Graph Bookings API only supported delegated permissions hence multiple users cannot be added via API.


## 3. No security for admin app (anyone can use)
After installing the admin app, it will be available to all members of the Teams channel. Security check is not implemented. Additional info on how to implement can be found in the **Taking It Further** documentation.

## 4. Front Door / custom domain required
There is a limitation with Teams tab SSO that requires that domain to not use azurewebsites.net. As an alternative, either use a custom domain or Azure Front Door. See [tab SSO docs](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso).

## 5. 1:1 welcome message localization
Upon application installation, a welcome message is sent. This message cannot be localized. This is because Microsoft Teams ConversationUpdate activities don't include the user's locale. Alternatively, the locale can be set and changed after a user has interacted with the application. 

## 6. Only Administrator is allowed to create and modify Bookings
There is a limitation only Bookings administrators can create or modify Bookings. 
To grant a staff member administrator, refer [this doc](https://docs.microsoft.com/en-us/microsoft-365/bookings/add-staff?view=o365-worldwide#add-staff) to change a member role to administrator.
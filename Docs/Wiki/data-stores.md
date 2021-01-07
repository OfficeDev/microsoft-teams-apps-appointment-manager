The app uses the following data stores:
1. Azure Cosmos DB with SQL API
1. Application Insights

All of these resources are created in your Azure subscription. None are hosted directly by Microsoft.

## Azure Cosmos DB

Azure Cosmos DB holds the bulk of the app's data, split into various containers.

### Channels

The `Channels` container stores the channels that have access to the Appointment Manager app.

| Value | Description
| ----- | -----------
| `tenantId` | The tenant ID of the associated team. |
| `serviceUrl` | The service URL used to post proactive messages. |
| `teamId` | The team ID of the associated team. |
| `teamAadObjectId` | The object ID of the associated team's AAD group. |
| `teamName` | The display name of the associated team. |
| `channelId` | The Teams ID of the channel. |
| `channelName` | The display name of the channel. |
| `createdById` | The ID of the user who created the item (not used). |
| `createdByName` | The name of the user who created the item (not used). |
| `createdDateTime` | The timestamp when the channel was created in the database. |
| `updatedById` | The ID of the user who last updated the item (not used). |
| `updatedDateTime` | The timestamp when the channel was last updated in the database. |
| `id` | A randomly generated GUID. |

### Channel Mappings

The `ChannelMappings` container stores the mappings from categories to Teams channels and Bookings services that the admin configures.

| Value | Description
| ----- | -----------
| `channelId` | The Teams ID of the channel. |
| `category` | The name of the category. |
| `bookingsBusiness.id` | The ID of the mapped Bookings business. |
| `bookingsBusiness.displayName` | The display name of the mapped Bookings business. |
| `bookingsService.id` | The ID of the mapped Bookings service. |
| `bookingsService.displayName` | The display name of the mapped Bookings service. |
| `supervisors[i].id` | The object ID of a supervisor for the category.  |
| `supervisors[i].displayName` | The display name of a supervisor for the category. |
| `createdById` | The ID of the user who created the item. |
| `createdByName` | The name of the user who created the item. |
| `createdDateTime` | The timestamp when the channel was created in the database. |
| `updatedById` | The ID of the user who last updated the item. |
| `updatedDateTime` | The timestamp when the channel was last updated in the database. |
| `id` | A randomly generated GUID. |

### Staff Members

The `Agents` container stores the staff members who are in teams that have Appointment Manager installed.

| Value | Description
| ----- | -----------
| `userPrincipalName` | The user principal name of the user. |
| `aadObjectId` | The object ID of the user. |
| `teamsId` | The Teams ID of the user. |
| `serviceUrl` | The service URL used to send proactive messages. |
| `bookingsStaffMemberId` | The ID of the Bookings Staff Member corresponding to the user. |
| `isAdmin` | Whether or not the user is an admin (not used). |
| `createdById` | The ID of the user who created the item (not used). |
| `createdByName` | The name of the user who created the item (not used). |
| `createdDateTime` | The timestamp when the staff member was created in the database. |
| `updatedById` | The ID of the user who last updated the item (not used). |
| `updatedDateTime` | The timestamp when the staff member was last updated in the database. |
| `id` | A randomly generated GUID. |

### Appointment Requests

The `ConsultRequests` container stores the appointment requests, including activity history, notes, and attachments.

| Value | Description
| ----- | -----------
| `customerName` | The name of the customer who requested the appointment. |
| `customerPhone` | The phone number of the customer who requested the appointment. |
| `customerEmail` | The email address of the customer who requested the appointment. |
| `query` | The question or request details from the customer. |
| `preferredTimes` | The preferred times specified by the customer for the appointment. |
| `friendlyId` | A randomly generated short identifier. Used when displaying appointment requests. |
| `category` | The category chosen by the customer for the appointment. Used to route requests to different Teams channels and Bookings services. |
| `status` | The status of the appointment request. Possible values are `Unassigned`, `Assigned`, `ReassignRequested`, and `Completed` |
| `assignedToId` | The object ID of the assigned staff member. |
| `assignedToName` | The name of the assigned staff member. |
| `assignedTimeBlock` | The scheduled time for the appointment request. |
| `bookingsBusinessId` | The ID of the Bookings business associated with the appointment request. |
| `bookingsServiceId` | The ID of the Bookings service associated with the appointment request. |
| `bookingsAppointmentId` | The ID of the Bookings appointment created for the appointment request.  |
| `joinUri` | The URL to join the Teams meeting. |
| `activities` | See below. |
| `activityId` | The Bot Framework activity ID of the message posted to the Teams channel. Used to update the message in the future. |
| `conversationId` | The Bot Framework conversation ID associated with the message posted to the Teams channel. Used to update the message in the future. |
| `notes` | See below. |
| `attachments` | See below. |
| `createdById` | The ID of the user who created the item (not used). |
| `createdByName` | The name of the user who created the item (not used). |
| `createdDateTime` | The timestamp when the appointment request was created in the database. |
| `updatedById` | The ID of the user who last updated the appointment request. |
| `updatedDateTime` | The timestamp when the appointment request was last updated in the database. |
| `id` | A randomly generated GUID. |

#### Appointment Request Activities

The `activities` property holds a list of activities performed on the appointment request, such as assignment and reassignment.

| Value | Description
| ----- | -----------
| `type` | The type of activity that occurred. Possible values are `Assigned`, `ReassignRequested`, and `Completed`. |
| `activityForUserId` | The ID of the user for whom the activity was performed. For assignment, this is the user who was assigned to the request. |
| `activityForUserName` | The name of the user corresponding to `activityForUserId`.  |
| `comment` | The comment left by the user when performing the activity. |
| `createdById` | The ID of the user who performed the activity. For assignment, this is the user who assigned the request. |
| `createdByName` | The name of the user who performed the activity. |
| `createdDateTime` | The timestamp when the activity occurred. |
| `updatedById` | The ID of the user who last updated the activity (not used). |
| `updatedDateTime` | The timestamp when the activity was last updated in the database (not used). |
| `id` | A randomly generated GUID. |

#### Appointment Request Notes

The `notes` property holds a list of notes added to the appointment request.

| Value | Description
| ----- | -----------
| `text` | The notes added by the staff member. |
| `createdById` | The ID of the user who created the item. |
| `createdByName` | The name of the user who created the item. |
| `createdDateTime` | The timestamp when the note was created in the database. |
| `updatedById` | The ID of the user who last updated the item (not used). |
| `updatedDateTime` | The timestamp when the note was last updated in the database (not used). |
| `id` | A randomly generated GUID. |

#### Appointment Request Attachments

The `attachments` property holds a list of attachments added to the appointment request.

| Value | Description
| ----- | -----------
| `filename` | The name of the file sent by the customer. |
| `uri` | The URI of the file sent by the customer. |
| `title` | The attachment title given by the staff member. |
| `id` | A randomly generated GUID. |

## Application Insights

See [Telemetry](telemetry) for details.
# teams-activity-notifications-function

This repository contains the source code artifacts for sending Teams activity feed notifications to users using Azure Event Grid, Azure Functions, Teams App, and Microsoft Graph API. Follow this [article](article) for more details.

There are two folders in this repository:
- **TeamsApp**: This contains the Teams app manifest and images that can be packaged and uploaded to Teams app catalog in a tenant.
- **AzureFunction**: This contains the code for event grid triggered Azure Function that uses Microsoft Graph API to send activity notifications to Teams users.
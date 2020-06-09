---
page_type: sample
languages:
- csharp
products:
- office-teams
description: Scrums for channels helps you get status updates from your team in channel scope
urlFragment: microsoft-teams-apps-scrumsforchannels
---

# Scrums for Channels

| [Documentation](https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/Home) | [Deployment guide](https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/Deployment-Guide) | [Architecture](https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/Solution-Overview) |
| ---- | ---- | ---- |

Scrums for Channels is a simple scrum assistant application that enables users to run and schedule stand-up meetings and provide an easy way to share daily updates. The bot works in team channels and all members who have been added to a scrum team can participate in the scrum.
The app works great for teams that have members participating from varied geographical locations, different time zones or even fully remote teams. 

Using the Scrums for Channels app in Microsoft Teams, users will be able to:
 -  Schedule a scrum at a given time based on a time zone 
 -  Select the team members who will be part of the scrum
 -  Run scrums in a channel
 -  Configure multiple scrums to run in different or same channels
 -  If the setting is enabled, Export scrum details for the past 30 days in xlsx file

 **App workflow**

-	Tony is a Project lead in Contoso Technologies and leads multiple teams that are spread across multiple locations and time zones. He uses the Scrums for Channels app in Microsoft Teams to setup scrums and enable all his team members to share their daily work updates in an easy and concise manner
-	Once the app is installed, he opens the settings screen using the command ‘Open settings’ which will allow him to configure a scrum
-	He will select the required team members, the start time of the scrum, the time zone in which the scrum will be initiated and the channel for scrum updates 
-	He can choose to configure multiple scrums if required by selecting different team members
-	The app auto starts a scrum as per the schedule and sends an adaptive card with scrum details in the channel as configured and notify all the associated team members. 
-	The adaptive card will display scrum summary details like no of people who contributed to the scrum, the status of the scrum and no of team members that are blocked, if any. 
-	All the team members will be able to update their individual details and see what others in the team have posted through action buttons on the adaptive card. On every update, the card is refreshed to reflect the current scrum status
-	Anyone in the team can end a scrum at a given point in time. Once a scrum ends, the card will be refreshed reflecting the scrum status as Closed
-	The app will initiate the scrum again the next day at the scheduled time. If the previous scrum is still active, it will mark that as Closed before starting a new scrum


Here are some screenshots of a user interacting with Scrums for Channels :

**Configure scrums**

![Scrums for Channels settings task module screen](https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/Images/SettingsScreen.png)

**Provide your updates when a scrum is active**

![Scrum status screen adaptive card with @mentions](https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/Images/ScrumStatus.png)


**View details updated by you and others**

![Scrum details task module screen](https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/Images/ScrumDetails.png)

## Legal notice

This app template is provided under the [MIT License](https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/blob/master/LICENSE) terms.  In addition to these terms, by using this app template you agree to the following:

-	You are responsible for complying with all applicable privacy and security regulations related to use, collection and handling of any personal data by your app.  This includes complying with all internal privacy and security policies of your organization if your app is developed to be sideloaded internally within your organization.

-	Where applicable, you may be responsible for data related incidents or data subject requests for data collected through your app.

-	Any trademarks or registered trademarks of Microsoft in the United States and/or other countries and logos included in this repository are the property of Microsoft, and the license for this project does not grant you rights to use any Microsoft names, logos or trademarks outside of this repository.  Microsoft’s general trademark guidelines can be found [here](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general.aspx).

-	Use of this template does not guarantee acceptance of your app to the Teams app store.  To make this app available in the Teams app store, you will have to comply with the [submission and validation process](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/publish), and all associated requirements such as including your own privacy statement and terms of use for your app.

## Getting started

Begin with the [Solution overview](https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/Solution-overview) to read about what the app does and how it works.

When you're ready to try out Scrums for Channels, or to use it in your own organization, follow the steps in the [Deployment guide](https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/Deployment-Guide).

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

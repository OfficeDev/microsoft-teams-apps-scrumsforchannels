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

Scrums for Channels is a simple scrum assistant bot that enables users to run and schedule stand-up meetings and provide an easy way to share daily updates. The bot works in team channels and all members who have been added to a scrum team can participate in the scrum.
The bot works great for teams that have members participating from varied geographical locations, different time zones or even fully remote teams. 

Using the Scrums for Channels app in Microsoft Teams, users will be able to:
 -  Schedule a scrum at a given time based on a time zone 
 -  Select the team members who will be part of the scrum
 -  Run scrums in a channel
 -  Configure multiple scrums to run in different or same channels
 -  Export scrum details for the past 30 days in xlsx file

A typical scrum workflow using the app will be:
 - The bot will auto start the scrum at the specified time
 - It will post an adaptive card in the channel with buttons to share status updates, view details input by other scrum members and to end the scrum. The card will also display other details like the status of the scrum (active or closed), the number of people who have contributed to the scrum and  the number of people who have marked their status as blocked
 - Users can choose to share their updates, view details updated by other team members and end the scrum

Here are some screenshots of a user interacting with Scrums for Channels :

**Configure scrums**

[[https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/Images/SettingsScreen.png|Settings screen]]

**Provide your updates when a scrum is active**

[[https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/Images/ScrumStatus.png|Scrum status screen adaptive card with @mentions]]


**View details updated by you and others**

[[https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/Images/ScrumDetails.png|Scrum details]]

## Legal notice

This app template is provided under the [MIT License](https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/blob/master/LICENSE) terms.  In addition to these terms, by using this app template you agree to the following:

-	You are responsible for complying with all applicable privacy and security regulations related to use, collection and handling of any personal data by your app.  This includes complying with all internal privacy and security policies of your organization if your app is developed to be sideloaded internally within your organization.

-	Where applicable, you may be responsible for data related incidents or data subject requests for data collected through your app.

-	Any trademarks or registered trademarks of Microsoft in the United States and/or other countries and logos included in this repository are the property of Microsoft, and the license for this project does not grant you rights to use any Microsoft names, logos or trademarks outside of this repository.  Microsoft’s general trademark guidelines can be found [here](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general.aspx).

-	Use of this template does not guarantee acceptance of your app to the Teams app store.  To make this app available in the Teams app store, you will have to comply with the [submission and validation process](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/publish), and all associated requirements such as including your own privacy statement and terms of use for your app.

## Getting started

Begin with the [Solution overview](https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/Solution-overview) to read about what the app does and how it works.

When you're ready to try out Scrums for Channels, or to use it in your own organization, follow the steps in the [Deployment guide](https://github.com/OfficeDev/microsoft-teams-apps-scrumsforchannels/wiki/DeployementGuide).

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

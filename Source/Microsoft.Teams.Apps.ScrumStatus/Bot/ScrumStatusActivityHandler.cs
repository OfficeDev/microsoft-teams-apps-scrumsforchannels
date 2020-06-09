// <copyright file="ScrumStatusActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.ScrumStatus.Cards;
    using Microsoft.Teams.Apps.ScrumStatus.Common;
    using Microsoft.Teams.Apps.ScrumStatus.Helpers;
    using Microsoft.Teams.Apps.ScrumStatus.Models;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// The ScrumStatusActivityHandler is responsible for reacting to incoming events from Teams sent from BotFramework.
    /// </summary>
    public sealed class ScrumStatusActivityHandler : TeamsActivityHandler
    {
        /// <summary>
        /// Represents expiry of token in minutes.
        /// </summary>
        private const int JwtExpiryMinutes = 60;

        /// <summary>
        /// Represents the Application base Uri.
        /// </summary>
        private readonly string appBaseUri;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<ScrumStatusActivityHandler> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Instance of Application Insights Telemetry client.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Storage helper for working with scrum data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumStorageProvider scrumStorageProvider;

        /// <summary>
        /// Storage helper for working with scrum status data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumStatusStorageProvider scrumStatusStorageProvider;

        /// <summary>
        /// Storage helper for working with scrum configuration data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumConfigurationStorageProvider scrumConfigurationStorageProvider;

        /// <summary>
        /// Generating custom JWT token and retrieving access token for user.
        /// </summary>
        private readonly ITokenHelper tokenHelper;

        /// <summary>
        /// Instance of class that handles card create/update helper methods.
        /// </summary>
        private readonly CardHelper cardHelper;

        /// <summary>
        /// Instance of class that handles Bot activity helper methods.
        /// </summary>
        private readonly ActivityHelper activityHelper;

        /// <summary>
        /// Instance of class that handles scrum helper methods.
        /// </summary>
        private readonly ScrumHelper scrumHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrumStatusActivityHandler"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        /// <param name="scrumStorageProvider">Scrum storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="scrumStatusStorageProvider">Scrum status storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="scrumConfigurationStorageProvider">Scrum configuration storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="tokenHelper">Generating custom JWT token and retrieving access token for user.</param>
        /// <param name="cardHelper">Instance of class that handles card create/update helper methods.</param>
        /// <param name="activityHelper">Instance of class that handles Bot activity helper methods.</param>
        /// <param name="scrumHelper">Instance of class that handles scrum helper methods.</param>
        public ScrumStatusActivityHandler(
            ILogger<ScrumStatusActivityHandler> logger,
            IStringLocalizer<Strings> localizer,
            TelemetryClient telemetryClient,
            IOptions<ScrumStatusActivityHandlerOptions> options,
            IScrumStorageProvider scrumStorageProvider,
            IScrumStatusStorageProvider scrumStatusStorageProvider,
            IScrumConfigurationStorageProvider scrumConfigurationStorageProvider,
            ITokenHelper tokenHelper,
            CardHelper cardHelper,
            ActivityHelper activityHelper,
            ScrumHelper scrumHelper)
        {
            options = options ?? throw new ArgumentNullException(nameof(options));
            this.appBaseUri = options.Value.AppBaseUri;
            this.logger = logger;
            this.localizer = localizer;
            this.telemetryClient = telemetryClient;
            this.scrumStorageProvider = scrumStorageProvider;
            this.scrumStatusStorageProvider = scrumStatusStorageProvider;
            this.scrumConfigurationStorageProvider = scrumConfigurationStorageProvider;
            this.tokenHelper = tokenHelper;
            this.cardHelper = cardHelper;
            this.scrumHelper = scrumHelper;
            this.activityHelper = activityHelper;
        }

        /// <summary>
        /// Invoked when Bot/Messaging Extension is installed in team to send welcome card.
        /// </summary>
        /// <param name="membersAdded">A list of all the members added to the conversation, as described by the conversation update activity.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents welcome card when bot is added first time by user.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.activityhandler.onmembersaddedasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task OnMembersAddedAsync(
            IList<ChannelAccount> membersAdded,
            ITurnContext<IConversationUpdateActivity> turnContext,
            CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            this.RecordEvent(nameof(this.OnMembersAddedAsync), turnContext);

            var activity = turnContext.Activity;
            this.logger.LogInformation($"conversationType: {activity.Conversation.ConversationType}, membersAdded: {activity.MembersAdded?.Count}");

            if (activity.MembersAdded.Any(member => member.Id == activity.Recipient.Id))
            {
                this.logger.LogInformation($"Bot added {activity.Conversation.Id}");
                var userWelcomeCardAttachment = WelcomeCard.GetWelcomeCardAttachmentForChannel(this.appBaseUri, this.localizer);
                await turnContext.SendActivityAsync(MessageFactory.Attachment(userWelcomeCardAttachment), cancellationToken);
            }
        }

        /// <summary>
        /// When OnTurn method receives a fetch invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents a task module response.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamstaskmodulefetchasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(
            ITurnContext<IInvokeActivity> turnContext,
            TaskModuleRequest taskModuleRequest,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnTeamsTaskModuleFetchAsync), turnContext);

                var activity = turnContext.Activity;
                AdaptiveSubmitActionData adaptiveSubmitActionData = JsonConvert.DeserializeObject<AdaptiveSubmitActionData>(taskModuleRequest?.Data?.ToString());

                if (adaptiveSubmitActionData == null)
                {
                    this.logger.LogInformation("Value obtained from task module fetch action is null");
                    return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("ErrorMessage"), this.localizer.GetString("BotFailureTitle"));
                }

                string scrumStartActivityId = adaptiveSubmitActionData.ScrumStartActivityId;
                string adaptiveActionType = adaptiveSubmitActionData.AdaptiveActionType;

                switch (adaptiveActionType.ToUpperInvariant())
                {
                    case Constants.ScrumDetailsTaskModuleCommand: // Command to show scrum details page.
                        if (!string.IsNullOrEmpty(scrumStartActivityId))
                        {
                            return await this.cardHelper.GetScrumDetailsCardResponseAsync(adaptiveSubmitActionData.ScrumMembers, adaptiveSubmitActionData.ScrumTeamConfigId, scrumStartActivityId, turnContext, cancellationToken);
                        }

                        break;

                    case Constants.UpdateStatusTaskModuleCommand: // Command to show update scrum status page.
                        string aadGroupId = await this.activityHelper.GetTeamAadGroupIdAsync(turnContext, cancellationToken);
                        var scrumInfo = await this.scrumHelper.GetActiveScrumAsync(adaptiveSubmitActionData.ScrumTeamConfigId, aadGroupId);
                        if (scrumInfo == null || scrumInfo.IsCompleted)
                        {
                            this.logger.LogInformation($"The scrum is not running at this moment: {activity.Conversation.Id}");
                            return this.cardHelper.GetTaskModuleErrorResponse(string.Format(CultureInfo.CurrentCulture, this.localizer.GetString("ErrorScrumDoesNotExist"), activity.From.Name), this.localizer.GetString("UpdateStatusTitle"));
                        }

                        var activityId = this.activityHelper.CheckUserExistsInScrumMembers(scrumInfo.MembersActivityIdMap, activity.From.Id);
                        if (string.IsNullOrEmpty(activityId))
                        {
                            this.logger.LogInformation($"Member who is updating the scrum is not the part of scrum for: {activity.Conversation.Id}");
                            return this.cardHelper.GetTaskModuleErrorResponse(string.Format(CultureInfo.CurrentCulture, this.localizer.GetString("ErrorUserIsNotPartOfRunningScrumAndTryUpdateStatus"), activity.From.Name), this.localizer.GetString("UpdateStatusTitle"));
                        }

                        return this.cardHelper.GetScrumStatusUpdateCardResponse(adaptiveSubmitActionData.ScrumMembers, adaptiveSubmitActionData.ScrumTeamConfigId, scrumStartActivityId, new ScrumStatus());

                    case Constants.SettingsTaskModuleCommand: // Command to show scrum settings page.
                        string customAPIAuthenticationToken = this.tokenHelper.GenerateAPIAuthToken(applicationBasePath: activity.ServiceUrl, fromId: activity.From.Id, JwtExpiryMinutes);
                        return this.cardHelper.GetSettingsCardResponse(customAPIAuthenticationToken, this.telemetryClient?.InstrumentationKey, activity.ServiceUrl);

                    default:
                        this.logger.LogInformation($"Invalid command for task module fetch activity. Command is: {adaptiveActionType}");
                        return this.cardHelper.GetTaskModuleErrorResponse(string.Format(CultureInfo.CurrentCulture, this.localizer.GetString("TaskModuleInvalidCommandText"), adaptiveActionType), this.localizer.GetString("BotFailureTitle"));
                }

                return null;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while fetching task module.", SeverityLevel.Error);
                return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("ErrorMessage"), this.localizer.GetString("BotFailureTitle"));
            }
        }

        /// <summary>
        /// When OnTurn method receives a submit invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents a task module response.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamstaskmodulesubmitasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(
            ITurnContext<IInvokeActivity> turnContext,
            TaskModuleRequest taskModuleRequest,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnTeamsTaskModuleSubmitAsync), turnContext);

                var activity = turnContext.Activity;

                var activityValue = JObject.Parse(activity.Value?.ToString())["data"].ToString();
                ScrumStatus scrumStatus = JsonConvert.DeserializeObject<ScrumStatus>(activityValue);
                AdaptiveSubmitActionData adaptiveSubmitActionData = JsonConvert.DeserializeObject<AdaptiveSubmitActionData>(activityValue);

                if (scrumStatus == null || adaptiveSubmitActionData == null)
                {
                    this.logger.LogInformation("Value obtained from task module submit action is null");
                    return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("ErrorMessage"), this.localizer.GetString("BotFailureTitle"));
                }

                this.logger.LogInformation($"OnTeamsTaskModuleSubmitAsync: {JObject.Parse(activity.Value.ToString())["data"]}");

                switch (adaptiveSubmitActionData.AdaptiveActionType.ToUpperInvariant())
                {
                    case Constants.UpdateStatusTaskModuleCommand: // command to handle update status page submit.
                        string scrumStartActivityId = adaptiveSubmitActionData.ScrumStartActivityId;
                        string scrumMembers = adaptiveSubmitActionData.ScrumMembers;
                        string scrumTeamConfigId = adaptiveSubmitActionData.ScrumTeamConfigId;
                        if (string.IsNullOrWhiteSpace(scrumStatus.YesterdayTaskDescription) || string.IsNullOrWhiteSpace(scrumStatus.TodayTaskDescription))
                        {
                            return this.cardHelper.GetScrumStatusValidationCardResponse(scrumMembers, scrumTeamConfigId, scrumStartActivityId, scrumStatus);
                        }

                        this.logger.LogInformation($"Getting scrum configuration details which are active. ScrumTeamConfigId: {scrumTeamConfigId}");
                        string aadGroupId = await this.activityHelper.GetTeamAadGroupIdAsync(turnContext, cancellationToken);
                        var scrumConfigurationDetail = await this.scrumConfigurationStorageProvider.GetScrumConfigurationDetailByScrumTeamConfigIdAsync(scrumTeamConfigId, aadGroupId);
                        if (scrumConfigurationDetail == null || !scrumConfigurationDetail.IsActive)
                        {
                            this.logger.LogInformation($"scrum configuration details for the scrum team configuration id: {scrumTeamConfigId} could not be found or scrum is inactive");
                            return this.cardHelper.GetTaskModuleErrorResponse(string.Format(CultureInfo.CurrentCulture, this.localizer.GetString("ErrorScrumConfigurationDetailsNullOrInactive"), activity.From.Name), this.localizer.GetString("UpdateStatusTitle"));
                        }

                        var scrumDetail = (await this.scrumStorageProvider.GetScrumsByScrumStartActivityIdAsync(scrumStartActivityId, scrumConfigurationDetail.AadGroupId)).FirstOrDefault();

                        if (scrumDetail == null)
                        {
                            await turnContext.SendActivityAsync(this.localizer.GetString("ErrorScrumDoesNotExist"), cancellationToken: cancellationToken);
                            return null;
                        }

                        await this.scrumHelper.SaveScrumStatusDetailsAsync(turnContext, scrumStatus, adaptiveSubmitActionData, scrumDetail.ScrumStartCardResponseId, aadGroupId);
                        var membersActivityIdMap = JsonConvert.DeserializeObject<Dictionary<string, string>>(scrumMembers);
                        var updatedScrumSummary = await this.scrumHelper.GetScrumSummaryAsync(scrumTeamConfigId, scrumConfigurationDetail.AadGroupId, scrumDetail.ScrumStartCardResponseId, membersActivityIdMap);
                        if (updatedScrumSummary == null)
                        {
                            this.logger.LogInformation($"No data obtained from storage to update summary card for summaryCardActivityId : {scrumDetail.ScrumStartCardResponseId}");
                            return null;
                        }

                        await this.cardHelper.UpdateSummaryCardAsync(updatedScrumSummary, scrumDetail.ScrumStartCardResponseId, scrumTeamConfigId, scrumStartActivityId, membersActivityIdMap, scrumConfigurationDetail.TimeZone, turnContext, cancellationToken);
                        return null;

                    default:
                        return null;
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while submitting task module.", SeverityLevel.Error);
                return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("ErrorMessage"), this.localizer.GetString("BotFailureTitle"));
            }
        }

        /// <summary>
        /// Invoked when a message activity is received from the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// For more information on bot messaging in Teams, see the documentation
        /// https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/conversation-basics?tabs=dotnet#receive-a-message .
        /// </remarks>
        protected override async Task OnMessageActivityAsync(
            ITurnContext<IMessageActivity> turnContext,
            CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            this.RecordEvent(nameof(this.OnMessageActivityAsync), turnContext);

            var activity = turnContext.Activity;

            switch (activity.Conversation.ConversationType)
            {
                case Constants.ChannelConversationType:
                    await this.OnMessageActivityInChannelAsync(
                        activity,
                        turnContext,
                        cancellationToken);
                    break;

                default:
                    this.logger.LogInformation($"Received unexpected conversationType {activity.Conversation.ConversationType}", SeverityLevel.Warning);
                    break;
            }
        }

        /// <summary>
        /// Handle message activity in channel.
        /// </summary>
        /// <param name="message">A message in a conversation.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task OnMessageActivityInChannelAsync(
            IMessageActivity message,
            ITurnContext<IMessageActivity> turnContext,
            CancellationToken cancellationToken)
        {
            string actionType = message.Value != null ? JObject.Parse(message.Value.ToString())["AdaptiveActionType"]?.ToString() : null;
            string scrumMembers = message.Value != null ? JObject.Parse(message.Value.ToString())["ScrumMembers"]?.ToString() : null;
            message.RemoveRecipientMention();
            string text = string.IsNullOrEmpty(message.Text) ? actionType : message.Text;

            switch (text.ToUpperInvariant().Trim())
            {
                case Constants.EndScrum: // command to handle end scrum.
                    string conversationId = message.Conversation.Id;
                    string scrumTeamConfigId = JObject.Parse(message.Value.ToString())["ScrumTeamConfigId"].ToString();
                    string aadGroupId = await this.activityHelper.GetTeamAadGroupIdAsync(turnContext, cancellationToken);
                    var scrumInfo = await this.scrumHelper.GetActiveScrumAsync(scrumTeamConfigId, aadGroupId);
                    if (scrumInfo == null || scrumInfo.IsCompleted)
                    {
                        await turnContext.SendActivityAsync(string.Format(CultureInfo.CurrentCulture, this.localizer.GetString("ErrorScrumDoesNotExist"), turnContext.Activity.From.Name), cancellationToken: cancellationToken);
                        break;
                    }

                    scrumInfo.IsCompleted = true;
                    scrumInfo.ThreadConversationId = conversationId;
                    var scrumSaveResponse = await this.scrumStorageProvider.CreateOrUpdateScrumAsync(scrumInfo);
                    if (!scrumSaveResponse)
                    {
                        this.logger.LogError("Error in saving scrum information in storage.");
                        await turnContext.SendActivityAsync(this.localizer.GetString("ErrorSavingScrumData"), cancellationToken: cancellationToken);
                        break;
                    }

                    var activitySummary = await this.activityHelper.GetEndScrumSummaryActivityAsync(scrumInfo, conversationId, scrumMembers, turnContext, cancellationToken);
                    if (activitySummary != null)
                    {
                        this.logger.LogInformation($"Scrum completed by: {turnContext.Activity.From.AadObjectId} for {conversationId} with ScrumStartCardResponseId: {scrumInfo.ScrumStartCardResponseId}");
                        await turnContext.UpdateActivityAsync(activitySummary, cancellationToken);
                        await turnContext.SendActivityAsync(this.localizer.GetString("SuccessMessageAfterEndingScrum"), cancellationToken: cancellationToken);
                    }

                    break;

                case Constants.Help: // command to show help card.
                    this.logger.LogInformation("Sending help card");
                    var helpAttachment = HelpCard.GetHelpCard(this.localizer);
                    await turnContext.SendActivityAsync(MessageFactory.Attachment(helpAttachment), cancellationToken);
                    break;

                case Constants.Settings: // Command to show adaptive card with settings CTA button.
                    this.logger.LogInformation("Sending settings button card");
                    var settingsAttachment = SettingsCard.GetSettingsCard(this.localizer);
                    await turnContext.SendActivityAsync(MessageFactory.Attachment(settingsAttachment), cancellationToken);
                    break;

                default:
                    this.logger.LogInformation("Invalid command text entered in channel. Sending help card");
                    var helpAttachmentcard = HelpCard.GetHelpCard(this.localizer);
                    await turnContext.SendActivityAsync(MessageFactory.Attachment(helpAttachmentcard), cancellationToken);
                    break;
            }
        }

        /// <summary>
        /// Records event data to Application Insights telemetry client
        /// </summary>
        /// <param name="eventName">Name of the event.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        private void RecordEvent(string eventName, ITurnContext turnContext)
        {
            this.telemetryClient.TrackEvent(eventName, new Dictionary<string, string>
            {
                { "userId", turnContext.Activity.From.AadObjectId },
                { "tenantId", turnContext.Activity.Conversation.TenantId },
                { "teamId", turnContext.Activity.Conversation.Id },
                { "channelId", turnContext.Activity.ChannelId },
            });
        }
    }
}
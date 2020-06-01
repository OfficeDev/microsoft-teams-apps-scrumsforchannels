// <copyright file="CardHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.ScrumStatus.Cards;
    using Microsoft.Teams.Apps.ScrumStatus.Common;
    using Microsoft.Teams.Apps.ScrumStatus.Models;
    using Newtonsoft.Json;

    /// <summary>
    /// Class that handles card create/update helper methods.
    /// </summary>
    public class CardHelper
    {
        /// <summary>
        /// Represents the task module height for small card in case of error/warning.
        /// </summary>
        private const string TaskModuleErrorHeight = "small";

        /// <summary>
        /// Represents the task module width for medium card in case of error/warning.
        /// </summary>
        private const string TaskModuleErrorWidth = "small";

        /// <summary>
        /// Represents the task module height.
        /// </summary>
        private const int TaskModuleHeight = 600;

        /// <summary>
        /// Represents the task module width.
        /// </summary>
        private const int TaskModuleWidth = 600;

        /// <summary>
        /// Represents the task module width for settings page.
        /// </summary>
        private const int TaskModuleSettingsWidth = 1100;

        /// <summary>
        /// Represents the Application base Uri.
        /// </summary>
        private readonly string appBaseUri;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<CardHelper> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<ScrumStatusActivityHandlerOptions> options;

        /// <summary>
        /// Instance of class that handles scrum helper methods.
        /// </summary>
        private readonly ScrumHelper scrumHelper;

        /// <summary>
        /// Instance of class that handles Bot activity helper methods.
        /// </summary>
        private readonly ActivityHelper activityHelper;

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
        /// Initializes a new instance of the <see cref="CardHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        /// <param name="scrumHelper">Instance of class that handles scrum helper methods.</param>
        /// <param name="activityHelper">Instance of class that handles Bot activity helper methods.</param>
        /// <param name="scrumStorageProvider">Scrum storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="scrumStatusStorageProvider">Scrum status storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="scrumConfigurationStorageProvider">Scrum configuration storage provider to maintain data in Microsoft Azure table storage.</param>
        public CardHelper(
            ILogger<CardHelper> logger,
            IStringLocalizer<Strings> localizer,
            IOptions<ScrumStatusActivityHandlerOptions> options,
            ScrumHelper scrumHelper,
            ActivityHelper activityHelper,
            IScrumStorageProvider scrumStorageProvider,
            IScrumStatusStorageProvider scrumStatusStorageProvider,
            IScrumConfigurationStorageProvider scrumConfigurationStorageProvider)
        {
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.appBaseUri = this.options.Value.AppBaseUri;
            this.logger = logger;
            this.localizer = localizer;
            this.scrumHelper = scrumHelper;
            this.activityHelper = activityHelper;
            this.scrumStorageProvider = scrumStorageProvider;
            this.scrumStatusStorageProvider = scrumStatusStorageProvider;
            this.scrumConfigurationStorageProvider = scrumConfigurationStorageProvider;
        }

        /// <summary>
        /// Get the error card task module on validation failure.
        /// </summary>
        /// <param name="errorMessage">Error message to be displayed in task module.</param>
        /// <param name="title">Title for task module describing type of error.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public TaskModuleResponse GetTaskModuleErrorResponse(string errorMessage, string title)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = ErrorCard.GetErrorCardAttachment(errorMessage),
                        Height = TaskModuleErrorHeight,
                        Width = TaskModuleErrorWidth,
                        Title = title,
                    },
                },
            };
        }

        /// <summary>
        /// Update the scrum summary card with updated details.
        /// </summary>
        /// <param name="scrumSummary">Scrum summary information to be shown on card.</param>
        /// <param name="summaryCardActivityId">Summary card activity id.</param>
        /// <param name="scrumTeamConfigId">Unique identifier for scrum team configuration details.</param>
        /// <param name="scrumStartActivityId">Scrum start card activity id</param>
        /// <param name="membersActivityIdMap">Members id who are part of the scrum.</param>
        /// <param name="timeZone">Used to convert scrum start time as per specified time zone.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task of type boolean where true represents summary card updated successfully while false indicates failure in updating the summary card.</returns>
        public async Task<bool> UpdateSummaryCardAsync(ScrumSummary scrumSummary, string summaryCardActivityId, string scrumTeamConfigId, string scrumStartActivityId, Dictionary<string, string> membersActivityIdMap, string timeZone, ITurnContext turnContext, CancellationToken cancellationToken)
        {
            var activitySummary = MessageFactory.Attachment(ScrumCard.GetScrumStartCard(scrumSummary, membersActivityIdMap, scrumTeamConfigId, scrumStartActivityId, this.localizer, timeZone));
            activitySummary.Id = summaryCardActivityId;
            activitySummary.Conversation = turnContext?.Activity.Conversation;
            this.logger.LogInformation($"Trail card updated for: {turnContext.Activity.Conversation.Id} summaryCardActivityId: {summaryCardActivityId}");
            await turnContext.UpdateActivityAsync(activitySummary, cancellationToken);
            return true;
        }

        /// <summary>
        /// Update the first trail card with user details.
        /// </summary>
        /// <param name="scrumSummary">Scrum summary information to be shown on card.</param>
        /// <param name="scrum">Scrum details.</param>
        /// <param name="scrumConfiguration">Scrum configuration details.</param>
        /// <param name="membersActivityIdMap">Members id who are part of the scrum.</param>
        /// <param name="timeZone">Used to convert scrum start time as per specified time zone.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task of type bool where true represents summary card updated successfully while false indicates failure in updating the summary card.</returns>
        public async Task<bool> UpdateSummaryCardWithEndScrumAsync(ScrumSummary scrumSummary, Scrum scrum, ScrumConfiguration scrumConfiguration, Dictionary<string, string> membersActivityIdMap, string timeZone, ITurnContext turnContext, CancellationToken cancellationToken)
        {
            var activitySummary = MessageFactory.Attachment(ScrumCard.GetScrumStartCard(scrumSummary, membersActivityIdMap, scrumConfiguration?.ScrumTeamConfigId, scrum?.ScrumStartActivityId, this.localizer, timeZone));
            var teamsChannelInfo = await TeamsInfo.GetTeamChannelsAsync(turnContext, scrumConfiguration.TeamId, CancellationToken.None);
            var channelInfo = teamsChannelInfo.Where(channel => channel.Id.Equals(scrumConfiguration.ChannelId, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (channelInfo == null)
            {
                return false;
            }

            activitySummary.Id = scrum.ScrumStartCardResponseId;
            activitySummary.Conversation = new ConversationAccount
            {
                Id = $"{scrumConfiguration.ChannelId};messageid={scrum.ScrumStartCardResponseId}",
            };
            this.logger.LogInformation($"Trail card updated for: {scrum.ThreadConversationId} summaryCardActivityId: {scrum.ScrumStartCardResponseId}");
            await turnContext?.UpdateActivityAsync(activitySummary, cancellationToken);

            turnContext.Activity.Conversation = new ConversationAccount
            {
                Id = $"{scrumConfiguration.ChannelId};messageid={scrum.ScrumStartCardResponseId}",
            };
            await turnContext.SendActivityAsync(this.localizer.GetString("SuccessMessageAfterEndingScrum"), cancellationToken: cancellationToken);
            return true;
        }

        /// <summary>
        /// Get scrum details adaptive card response.
        /// </summary>
        /// <param name="scrumMembers">Members who are part of the scrum.</param>
        /// <param name="scrumTeamConfigId">Unique identifier for scrum team configuration details.</param>
        /// <param name="scrumStartActivityId">Scrum start card activity Id.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Returns scrum details card to be displayed in task module.</returns>
        public async Task<TaskModuleResponse> GetScrumDetailsCardResponseAsync(string scrumMembers, string scrumTeamConfigId, string scrumStartActivityId, ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            var activity = turnContext.Activity;
            this.logger.LogInformation($"Getting information about scrum with summaryCardId: {scrumStartActivityId}");
            string aadGroupId = await this.activityHelper.GetTeamAadGroupIdAsync(turnContext, cancellationToken);
            var scrumSummary = (await this.scrumStorageProvider.GetScrumsByScrumStartActivityIdAsync(scrumStartActivityId, aadGroupId)).FirstOrDefault();
            if (scrumSummary == null)
            {
                this.logger.LogInformation($"Value obtained from storage for scrum summary is null.");
                return this.GetTaskModuleErrorResponse(string.Format(CultureInfo.CurrentCulture, this.localizer.GetString("ErrorScrumDoesNotExist"), activity.From.Name), this.localizer.GetString("ScrumDetailsTitle"));
            }

            this.logger.LogInformation($"Received information about scrum with scrumId: {scrumSummary.ScrumId}");

            var scrumStatusDetails = await this.scrumStatusStorageProvider.GetScrumStatusBySummaryCardIdAsync(scrumSummary.ScrumStartCardResponseId, aadGroupId);
            var membersActivityIdMap = JsonConvert.DeserializeObject<Dictionary<string, string>>(scrumMembers);

            var updatedScrumSummary = await this.scrumHelper.GetScrumSummaryAsync(scrumTeamConfigId, aadGroupId, scrumSummary.ScrumStartCardResponseId, membersActivityIdMap);
            if (scrumStatusDetails == null || updatedScrumSummary == null)
            {
                this.logger.LogInformation($"Value obtained from storage for scrum is null.");
                return this.GetTaskModuleErrorResponse(string.Format(CultureInfo.CurrentCulture, this.localizer.GetString("ErrorScrumDoesNotExist"), activity.From.Name), this.localizer.GetString("ScrumDetailsTitle"));
            }

            var scrumConfigurationDetails = await this.scrumConfigurationStorageProvider.GetScrumConfigurationDetailByScrumTeamConfigIdAsync(scrumSummary.ScrumTeamConfigId, scrumSummary.AadGroupId);
            activity.Conversation.Id = scrumConfigurationDetails.TeamId;
            var validScrumMembers = new List<TeamsChannelAccount>();
            string continuationToken = null;
            do
            {
                var currentPage = await TeamsInfo.GetPagedMembersAsync(turnContext, 100, continuationToken, cancellationToken);
                continuationToken = currentPage.ContinuationToken;
                validScrumMembers.AddRange(currentPage.Members.Where(member => membersActivityIdMap.ContainsKey(member.Id)));
            }
            while (continuationToken != null);

            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = ScrumCard.GetScrumDetailsCard(scrumStatusDetails, updatedScrumSummary, validScrumMembers, this.appBaseUri, this.localizer, scrumConfigurationDetails.TimeZone),
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = this.localizer.GetString("ScrumDetailsTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get scrum status update adaptive card in response to be filled by scrum member.
        /// </summary>
        /// <param name="membersActivityIdMap">Members id who are part of the scrum.</param>
        /// <param name="scrumTeamConfigId">Unique identifier for scrum team configuration details.</param>
        /// <param name="scrumStartActivityId">Scrum start card activity Id.</param>
        /// <param name="scrumStatus">Scrum status details.</param>
        /// <returns>Returns scrum status update card to be displayed in task module.</returns>
        public TaskModuleResponse GetScrumStatusUpdateCardResponse(string membersActivityIdMap, string scrumTeamConfigId, string scrumStartActivityId, ScrumStatus scrumStatus)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = ScrumCard.GetScrumStatusUpdateCard(membersActivityIdMap, scrumTeamConfigId, scrumStartActivityId, scrumStatus, this.localizer),
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = this.localizer.GetString("UpdateStatusTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get settings adaptive card  in response to configure scrums in different channels.
        /// </summary>
        /// <param name="customAPIAuthenticationToken">JWT token used by client application to authenticate HTTP calls with API.</param>
        /// <param name="instrumentationKey">Instrumentation key for all ITelemetry objects logged in this TelemetryClient.</param>
        /// <param name="activityServicePath">Activity service URL.</param>
        /// <returns>Returns settings card to be displayed in task module.</returns>
        public TaskModuleResponse GetSettingsCardResponse(string customAPIAuthenticationToken, string instrumentationKey, string activityServicePath)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Url = $"{this.appBaseUri}/settings?token={customAPIAuthenticationToken}&telemetry={instrumentationKey}&serviceurl={activityServicePath}&theme={{theme}}&locale={{locale}}",
                        Height = TaskModuleHeight,
                        Width = TaskModuleSettingsWidth,
                        Title = this.localizer.GetString("SettingsTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get scrum status validation card with errors listed.
        /// </summary>
        /// <param name="membersActivityIdMap">Members id who are part of the scrum.</param>
        /// <param name="scrumTeamConfigId">Unique identifier for scrum team configuration details.</param>
        /// <param name="scrumStartActivityId">Scrum start card activity Id.</param>
        /// <param name="scrumStatus">Scrum status details filled by user.</param>
        /// <returns>Returns scrum status validation card to be sent in response.</returns>
        public TaskModuleResponse GetScrumStatusValidationCardResponse(string membersActivityIdMap, string scrumTeamConfigId, string scrumStartActivityId, ScrumStatus scrumStatus)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = ScrumCard.GetScrumStatusUpdateCard(membersActivityIdMap, scrumTeamConfigId, scrumStartActivityId, scrumStatus, this.localizer, string.IsNullOrWhiteSpace(scrumStatus?.YesterdayTaskDescription), string.IsNullOrWhiteSpace(scrumStatus?.TodayTaskDescription)),
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = this.localizer.GetString("UpdateStatusTitle"),
                    },
                },
            };
        }
    }
}
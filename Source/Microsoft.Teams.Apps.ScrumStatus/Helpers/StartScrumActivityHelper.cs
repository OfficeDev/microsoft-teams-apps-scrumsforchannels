// <copyright file="StartScrumActivityHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Xml;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.ScrumStatus;
    using Microsoft.Teams.Apps.ScrumStatus.Cards;
    using Microsoft.Teams.Apps.ScrumStatus.Common;
    using Microsoft.Teams.Apps.ScrumStatus.Models;
    using Newtonsoft.Json;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Class to start the scrum.
    /// </summary>
    public class StartScrumActivityHelper : IStartScrumActivityHelper
    {
        /// <summary>
        /// Represents retry delay.
        /// </summary>
        private const int RetryDelay = 1000;

        /// <summary>
        /// Represents retry count.
        /// </summary>
        private const int RetryCount = 2;

        /// <summary>
        /// Retry policy with jitter, Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </summary>
        private static readonly AsyncRetryPolicy RetryPolicy = Policy.Handle<Exception>()
          .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(RetryDelay), RetryCount));

        /// <summary>
        /// An instance of card helper to send class details.
        /// </summary>
        private readonly CardHelper cardHelper;

        /// <summary>
        /// Instance of class that handles scrum helper methods.
        /// </summary>
        private readonly ScrumHelper scrumHelper;

        /// <summary>
        /// Microsoft application credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter adapter;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<ScrumStatusActivityHandlerOptions> options;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<StartScrumActivityHelper> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Storage helper for working with scrum data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumStorageProvider scrumStorageProvider;

        /// <summary>
        /// Storage helper for working with scrum master data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumMasterStorageProvider scrumMasterStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="StartScrumActivityHelper"/> class.
        /// </summary>
        /// <param name="scrumStorageProvider">Instance for scrumStorageProvider.</param>
        /// <param name="scrumMasterStorageProvider">Scrum master storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="cardHelper">Instance of card helper to send class details.</param>
        /// <param name="scrumHelper">Instance of class that handles scrum helper methods.</param>
        /// <param name="microsoftAppCredentials">Instance for Microsoft application credentials.</param>
        /// <param name="adapter">An instance of bot adapter.</param>
        /// <param name="options">A set of key/value application configuration properties for activity handler.</param>
        /// <param name="logger">An instance of logger to log exception in application insights.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        public StartScrumActivityHelper(
            IScrumStorageProvider scrumStorageProvider,
            IScrumMasterStorageProvider scrumMasterStorageProvider,
            CardHelper cardHelper,
            ScrumHelper scrumHelper,
            MicrosoftAppCredentials microsoftAppCredentials,
            IBotFrameworkHttpAdapter adapter,
            IOptions<ScrumStatusActivityHandlerOptions> options,
            ILogger<StartScrumActivityHelper> logger,
            IStringLocalizer<Strings> localizer)
        {
            this.scrumStorageProvider = scrumStorageProvider;
            this.scrumMasterStorageProvider = scrumMasterStorageProvider;
            this.cardHelper = cardHelper;
            this.scrumHelper = scrumHelper;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.adapter = adapter;
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.logger = logger;
            this.localizer = localizer;
        }

        /// <summary>
        /// Method ends the existing scrum if running and then sends the start scrum card.
        /// </summary>
        /// <param name="scrumMaster">Scrum master details obtained from storage.</param>
        /// <returns>A task that ends the existing scrum and sends the start scrum card .</returns>
        public async Task ScrumStartActivityAsync(ScrumMaster scrumMaster)
        {
            if (scrumMaster != null)
            {
                try
                {
                    string serviceUrl = scrumMaster.ServiceUrl;
                    MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);

                    var conversationReference = new ConversationReference()
                    {
                        ChannelId = Constants.ConversationChannelId,
                        Bot = new ChannelAccount() { Id = this.microsoftAppCredentials.MicrosoftAppId },
                        ServiceUrl = serviceUrl,
                        Conversation = new ConversationAccount() { ConversationType = Constants.ConversationType, IsGroup = true, Id = scrumMaster.ChannelId, TenantId = this.options.Value.TenantId },
                    };

                    this.logger.LogInformation($"Sending start scrum command to channelId- {scrumMaster.ChannelId}");

                    await RetryPolicy.ExecuteAsync(async () =>
                    {
                        try
                        {
                            await ((BotFrameworkAdapter)this.adapter).ContinueConversationAsync(
                                this.microsoftAppCredentials.MicrosoftAppId,
                                conversationReference,
                                async (conversationTurnContext, conversationCancellationToken) =>
                                {
                                    bool isValidScrum = await this.EndExistingScrumAndStartScrumAsync(conversationTurnContext, scrumMaster, conversationCancellationToken);
                                    if (!isValidScrum)
                                    {
                                        this.logger.LogInformation("Error while ending the existing scrum.");
                                        await conversationTurnContext.SendActivityAsync(this.localizer.GetString(this.localizer.GetString("ErrorMessage")));
                                    }

                                    await this.SendScrumStartCardAsync(conversationTurnContext, scrumMaster, conversationCancellationToken);
                                },
                                CancellationToken.None);
                        }
                        catch (Exception ex)
                        {
                            this.logger.LogError(ex, "Error while performing retry logic to send scrum start card.");
                            throw;
                        }
                    });
                }
                catch (Exception ex)
                {
                    this.logger.LogError(ex, "Error while sending start scrum to channel from background service.");
                }
            }
        }

        /// <summary>
        /// Method to validate the existing scrum if already running.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="scrumMaster">Values obtained from ScrumMaster table.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that ends the existing scrum.</returns>
        private async Task<bool> EndExistingScrumAndStartScrumAsync(ITurnContext turnContext, ScrumMaster scrumMaster, CancellationToken cancellationToken)
        {
            // If previous scrum is already running end that and refresh scrum start card.
            var scrumInfo = await this.scrumStorageProvider.GetScrumByScrumMasterIdAsync(scrumMaster.ScrumMasterId);
            if (scrumInfo != null)
            {
                foreach (var scrum in scrumInfo)
                {
                    if (!scrum.IsCompleted)
                    {
                        scrum.IsCompleted = true;
                        var savedData = await this.scrumStorageProvider.CreateOrUpdateScrumAsync(scrum);
                        if (!savedData)
                        {
                            this.logger.LogInformation($"Error while updating scrim table from storage for scrumMasterId : {scrumMaster.ScrumMasterId}");
                            return false;
                        }

                        this.logger.LogInformation($"Getting scrum master details which are active. ScrumMasterId: {scrum.ScrumMasterId}");
                        var scrumMasterDetails = await this.scrumMasterStorageProvider.GetScrumMasterDetailsByScrumMasterIdAsync(scrum.ScrumMasterId);
                        if (scrumMasterDetails == null || !scrumMasterDetails.IsActive)
                        {
                            return false;
                        }

                        // End the existing running scrum and refresh start card with end scrum.
                        var scrumMembers = scrum.MembersActivityIdMap;
                        var membersActivityIdMap = JsonConvert.DeserializeObject<Dictionary<string, string>>(scrumMembers);
                        var updatedScrumSummary = await this.scrumHelper.GetScrumSummaryAsync(scrum.ScrumMasterId, scrum.ScrumStartCardResponseId, membersActivityIdMap);
                        await this.cardHelper.UpdateSummaryCardWithEndScrumAsync(updatedScrumSummary, scrum, scrumMaster, membersActivityIdMap, scrumMasterDetails.TimeZone, turnContext, cancellationToken);
                        this.logger.LogInformation($"Ended existing running scrum for {scrum.ThreadConversationId}");
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Method that sends the start scrum card to the channel.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="scrumMaster">Scrum master details obtained from storage.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that sends the start scrum card.</returns>
        private async Task SendScrumStartCardAsync(ITurnContext turnContext, ScrumMaster scrumMaster, CancellationToken cancellationToken)
        {
            try
            {
                string scrumMasterId = scrumMaster.ScrumMasterId;
                if (scrumMasterId != null)
                {
                    this.logger.LogInformation($"Scrum start for ID: {scrumMasterId}");
                    var scrumSummary = await this.scrumHelper.GetScrumSummaryAsync(scrumMasterId);

                    if (scrumSummary == null)
                    {
                        this.logger.LogInformation($"Scrum master details are deleted from storage.");
                        await turnContext.SendActivityAsync(string.Format(CultureInfo.CurrentCulture, this.localizer.GetString("ErrorScrumDeleted"), scrumMaster.TeamName), cancellationToken: cancellationToken);
                        return;
                    }

                    // scrumSummary.ScrumStartTime = scrumMaster.StartTime;
                    var scrumStartActivityId = Guid.NewGuid().ToString();

                    // Fetching the members list based on the teams id:
                    turnContext.Activity.Conversation.Id = scrumMaster.TeamId;
                    var membersActivityIdMap = await this.GetActivityIdOfMembersInScrumAsync(turnContext, scrumMaster, cancellationToken);
                    string membersList = JsonConvert.SerializeObject(membersActivityIdMap);

                    // Mentioning the participants involved in the scrum
                    var mentionActivity = await this.GetMentionsActivityAsync(turnContext, scrumMaster, cancellationToken);

                    // Check if channel exists. If channel doesn't exist then scrum card will be sent in General channel.
                    scrumMaster.ChannelId = await this.GetValidChannelIdAsync(turnContext, scrumMaster);

                    // Send the start scrum card
                    turnContext.Activity.Conversation.Id = scrumMaster.ChannelId;
                    var attachment = ScrumCard.GetScrumStartCard(scrumSummary, membersActivityIdMap, scrumMasterId, scrumStartActivityId, this.localizer, scrumMaster.TimeZone);
                    var scrumStartActivity = MessageFactory.Attachment(attachment);
                    var scrumStartActivityResponse = await turnContext.SendActivityAsync(scrumStartActivity, cancellationToken);

                    // Update the conversation id to send mentioned participants as reply to scrum start card.
                    turnContext.Activity.Conversation = new ConversationAccount
                    {
                        Id = $"{scrumMaster.ChannelId};messageid={scrumStartActivityResponse.Id}",
                    };
                    await turnContext.SendActivityAsync(mentionActivity, cancellationToken);
                    await this.CreateScrumAsync(scrumStartActivityResponse.Id, scrumStartActivityId, membersList, scrumMaster, turnContext, cancellationToken);
                    this.logger.LogInformation($"Scrum start details saved to table storage for: {turnContext.Activity.Conversation.Id}");
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Start scrum failed for {turnContext.Activity.Conversation.Id}: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        ///  Create a new scrum from the input.
        /// </summary>
        /// <param name="scrumStartCardResponseId">Activity id of scrum summary card.</param>
        /// <param name="scrumCardId">Activity Id of scrum card.</param>
        /// <param name="members">JSON serialized member and activity mapping.</param>
        /// <param name="scrumMaster">An instance of scrum master details.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>void.</returns>
        private async Task CreateScrumAsync(string scrumStartCardResponseId, string scrumCardId, string members, ScrumMaster scrumMaster, ITurnContext turnContext, CancellationToken cancellationToken)
        {
            string conversationId = turnContext.Activity.Conversation.Id;
            try
            {
                Scrum scrumEntity = new Scrum
                {
                    ThreadConversationId = conversationId,
                    ScrumStartActivityId = scrumCardId,
                    IsCompleted = false,
                    MembersActivityIdMap = members,
                    ScrumStartCardResponseId = scrumStartCardResponseId,
                    ScrumMasterId = scrumMaster.ScrumMasterId,
                    ScrumId = conversationId,
                    ChannelName = scrumMaster.ChannelName,
                    TeamId = scrumMaster.TeamId,
                    CreatedOn = DateTime.UtcNow.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'", CultureInfo.InvariantCulture),
                    AADGroupID = scrumMaster.AADGroupID,
                };
                var savedData = await this.scrumStorageProvider.CreateOrUpdateScrumAsync(scrumEntity);
                if (!savedData)
                {
                    await turnContext.SendActivityAsync(this.localizer.GetString("ErrorSavingScrumData"), cancellationToken: cancellationToken);
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"For {conversationId}: saving scrum data to table storage failed. {ex.Message}", SeverityLevel.Error);
                await turnContext.SendActivityAsync(this.localizer.GetString("ErrorMessage"), cancellationToken: cancellationToken);
                throw;
            }
        }

        /// <summary>
        /// Method to show GetMentionsActivity.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="scrumMasterDetails">Scrum master details.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>member ids.</returns>
        private async Task<Activity> GetMentionsActivityAsync(ITurnContext turnContext, ScrumMaster scrumMasterDetails, CancellationToken cancellationToken)
        {
            var members = await this.scrumHelper.GetValidMembersInScrumAsync(turnContext, scrumMasterDetails, cancellationToken);
            StringBuilder membersMention = new StringBuilder();
            var entities = new List<Entity>();
            var mentions = new List<Mention>();
            foreach (var member in members)
            {
                membersMention.Append(" ");
                var mention = new Mention
                {
                    Mentioned = new ChannelAccount()
                    {
                        Id = member.Id,
                        Name = member.Name,
                    },
                    Text = $"<at>{XmlConvert.EncodeName(member.Name)}</at>",
                };
                mentions.Add(mention);
                entities.Add(mention);
                membersMention.Append(mention.Text).Append(",").Append(" ");
            }

            membersMention = membersMention.Insert(0, "CC: ");
            var replyActivity = MessageFactory.Text(membersMention.ToString().Trim().TrimEnd(','));
            replyActivity.Entities = entities;
            return replyActivity;
        }

        /// <summary>
        /// Method to show name card and get members list.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="scrumMaster">Scrum master details.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Returns member ids as collection.</returns>
        private async Task<Dictionary<string, string>> GetActivityIdOfMembersInScrumAsync(ITurnContext turnContext, ScrumMaster scrumMaster, CancellationToken cancellationToken)
        {
            var membersActivityIdMap = new Dictionary<string, string>();

            if (scrumMaster == null)
            {
                return null;
            }

            var members = await this.scrumHelper.GetValidMembersInScrumAsync(turnContext, scrumMaster, cancellationToken);
            foreach (var member in members)
            {
                membersActivityIdMap[member.Id] = Guid.NewGuid().ToString();
            }

            return membersActivityIdMap;
        }

        /// <summary>
        /// Get general channel Id if scrum channel id does not exist.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="scrumMasterDetails">Scrum master details.</param>
        /// <returns>Returns general channel Id if scrum channel id does not exist.</returns>
        private async Task<string> GetValidChannelIdAsync(ITurnContext turnContext, ScrumMaster scrumMasterDetails)
        {
            var teamsChannelInfo = await TeamsInfo.GetTeamChannelsAsync(turnContext, scrumMasterDetails.TeamId, CancellationToken.None);
            var channelInfo = teamsChannelInfo.Where(channel => channel.Id.Equals(scrumMasterDetails.ChannelId, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

            if (channelInfo != null)
            {
                return scrumMasterDetails.ChannelId;
            }

            scrumMasterDetails.ChannelId = scrumMasterDetails.TeamId;
            scrumMasterDetails.ChannelName = Constants.GeneralChannel;

            List<ScrumMaster> scrumMasters = new List<ScrumMaster>();
            scrumMasters.Add(scrumMasterDetails);

            var saveResponse = await this.scrumMasterStorageProvider.StoreOrUpdateScrumMasterEntitiesAsync(scrumMasters);
            if (!saveResponse)
            {
                this.logger.LogError("Error while saving scrum master details.");
            }

            return scrumMasterDetails.TeamId;
        }
    }
}
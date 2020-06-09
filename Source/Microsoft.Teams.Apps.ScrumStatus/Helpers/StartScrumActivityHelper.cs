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
    using Microsoft.Bot.Schema.Teams;
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
        /// Represents channel conversation id.
        /// </summary>
        public const string TeamsBotFrameworkChannelId = "msteams";

        /// <summary>
        /// Retry policy with jitter, Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </summary>
        private static readonly AsyncRetryPolicy RetryPolicy = Policy.Handle<Exception>()
          .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(1000), 2));

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
        /// Storage helper for working with scrum configuration data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumConfigurationStorageProvider scrumConfigurationStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="StartScrumActivityHelper"/> class.
        /// </summary>
        /// <param name="scrumStorageProvider">Instance for scrumStorageProvider.</param>
        /// <param name="scrumConfigurationStorageProvider">Scrum configuration storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="cardHelper">Instance of card helper to send class details.</param>
        /// <param name="scrumHelper">Instance of class that handles scrum helper methods.</param>
        /// <param name="microsoftAppCredentials">Instance for Microsoft application credentials.</param>
        /// <param name="adapter">An instance of bot adapter.</param>
        /// <param name="options">A set of key/value application configuration properties for activity handler.</param>
        /// <param name="logger">An instance of logger to log exception in application insights.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        public StartScrumActivityHelper(
            IScrumStorageProvider scrumStorageProvider,
            IScrumConfigurationStorageProvider scrumConfigurationStorageProvider,
            CardHelper cardHelper,
            ScrumHelper scrumHelper,
            MicrosoftAppCredentials microsoftAppCredentials,
            IBotFrameworkHttpAdapter adapter,
            IOptions<ScrumStatusActivityHandlerOptions> options,
            ILogger<StartScrumActivityHelper> logger,
            IStringLocalizer<Strings> localizer)
        {
            this.options = options;
            this.scrumStorageProvider = scrumStorageProvider;
            this.scrumConfigurationStorageProvider = scrumConfigurationStorageProvider;
            this.cardHelper = cardHelper;
            this.scrumHelper = scrumHelper;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.adapter = adapter;
            this.logger = logger;
            this.localizer = localizer;
        }

        /// <summary>
        /// Method ends the existing scrum if running and then sends the start scrum card.
        /// </summary>
        /// <param name="scrumConfiguration">Scrum configuration details obtained from storage.</param>
        /// <returns>A task that ends the existing scrum and sends the start scrum card .</returns>
        public async Task ScrumStartActivityAsync(ScrumConfiguration scrumConfiguration)
        {
            try
            {
                string serviceUrl = scrumConfiguration.ServiceUrl;
                MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);

                var conversationReference = new ConversationReference()
                {
                    ChannelId = TeamsBotFrameworkChannelId,
                    Bot = new ChannelAccount() { Id = $"28:{this.microsoftAppCredentials.MicrosoftAppId}" },
                    ServiceUrl = serviceUrl,
                    Conversation = new ConversationAccount() { ConversationType = Constants.ChannelConversationType, IsGroup = true, Id = scrumConfiguration.ChannelId, TenantId = this.options.Value.TenantId },
                };

                this.logger.LogInformation($"Sending start scrum card to channelId: {scrumConfiguration.ChannelId}");

                await RetryPolicy.ExecuteAsync(async () =>
                {
                    try
                    {
                        await ((BotFrameworkAdapter)this.adapter).ContinueConversationAsync(
                            this.microsoftAppCredentials.MicrosoftAppId,
                            conversationReference,
                            async (conversationTurnContext, conversationCancellationToken) =>
                            {
                                bool isValidScrum = await this.EndExistingScrumAsync(conversationTurnContext, scrumConfiguration, conversationCancellationToken);
                                if (!isValidScrum)
                                {
                                    this.logger.LogInformation("Error while ending the existing scrum.");
                                    await conversationTurnContext.SendActivityAsync(this.localizer.GetString(this.localizer.GetString("ErrorMessage")));
                                }

                                await this.SendScrumStartCardAsync(conversationTurnContext, scrumConfiguration, conversationCancellationToken);
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

        /// <summary>
        /// Method to validate the existing scrum if already running.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="scrumConfiguration">Values obtained from ScrumConfiguration table.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that ends the existing scrum.</returns>
        private async Task<bool> EndExistingScrumAsync(ITurnContext turnContext, ScrumConfiguration scrumConfiguration, CancellationToken cancellationToken)
        {
            // If previous scrum is already running end that and refresh scrum start card.
            var scrumDetails = await this.scrumStorageProvider.GetScrumsByScrumTeamConfigIdAsync(scrumConfiguration.ScrumTeamConfigId, scrumConfiguration.AadGroupId);
            if (scrumDetails == null)
            {
                return false;
            }

            foreach (var scrumDetail in scrumDetails)
            {
                if (!scrumDetail.IsCompleted)
                {
                    scrumDetail.IsCompleted = true;
                    var savedData = await this.scrumStorageProvider.CreateOrUpdateScrumAsync(scrumDetail);
                    if (!savedData)
                    {
                        this.logger.LogInformation($"Error while updating scrim table from storage for scrumTeamConfigId : {scrumConfiguration.ScrumTeamConfigId}");
                        return false;
                    }

                    this.logger.LogInformation($"Getting scrum configuration details which are active. ScrumTeamConfigId: {scrumDetail.ScrumTeamConfigId}");
                    var scrumConfigurationDetails = await this.scrumConfigurationStorageProvider.GetScrumConfigurationDetailByScrumTeamConfigIdAsync(scrumDetail.ScrumTeamConfigId, scrumDetail.AadGroupId);
                    if (scrumConfigurationDetails == null || !scrumConfigurationDetails.IsActive)
                    {
                        return false;
                    }

                    // End the existing running scrum and refresh start card with end scrum.
                    var scrumMembers = scrumDetail.MembersActivityIdMap;
                    var membersActivityIdMap = JsonConvert.DeserializeObject<Dictionary<string, string>>(scrumMembers);
                    var updatedScrumSummary = await this.scrumHelper.GetScrumSummaryAsync(scrumDetail.ScrumTeamConfigId, scrumConfiguration.AadGroupId, scrumDetail.ScrumStartCardResponseId, membersActivityIdMap);

                    if (updatedScrumSummary == null)
                    {
                        this.logger.LogInformation($"No data obtained from storage to update summary card for scrumStartCardActivityId : {scrumDetail.ScrumStartCardResponseId}");
                        continue;
                    }

                    await this.cardHelper.UpdateSummaryCardWithEndScrumAsync(updatedScrumSummary, scrumDetail, scrumConfiguration, membersActivityIdMap, scrumConfigurationDetails.TimeZone, turnContext, cancellationToken);
                    this.logger.LogInformation($"Ended existing running scrum for {scrumDetail.ThreadConversationId}");
                }
            }

            return true;
        }

        /// <summary>
        /// Method that sends the start scrum card to the channel.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="scrumConfiguration">scrum configuration details obtained from storage.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that sends the start scrum card.</returns>
        private async Task SendScrumStartCardAsync(ITurnContext turnContext, ScrumConfiguration scrumConfiguration, CancellationToken cancellationToken)
        {
            try
            {
                string scrumTeamConfigId = scrumConfiguration.ScrumTeamConfigId;
                if (scrumTeamConfigId != null)
                {
                    this.logger.LogInformation($"Scrum start for ID: {scrumTeamConfigId}");
                    var scrumSummary = await this.scrumHelper.GetScrumSummaryAsync(scrumTeamConfigId, scrumConfiguration.AadGroupId);

                    if (scrumSummary == null)
                    {
                        this.logger.LogInformation($"scrum configuration details are deleted from storage.");
                        await turnContext.SendActivityAsync(string.Format(CultureInfo.CurrentCulture, this.localizer.GetString("ErrorScrumDeleted"), scrumConfiguration.ScrumTeamName), cancellationToken: cancellationToken);
                        return;
                    }

                    var scrumStartActivityId = Guid.NewGuid().ToString();

                    // Fetching the members list based on the teams id:
                    turnContext.Activity.Conversation.Id = scrumConfiguration.TeamId;
                    var scrumMembers = await this.GetValidMembersInScrumAsync(turnContext, scrumConfiguration.UserPrincipalNames, cancellationToken);
                    if (scrumMembers == null)
                    {
                        this.logger.LogInformation($"No scrum members are available to provide the scrum status");
                        await turnContext.SendActivityAsync(this.localizer.GetString("ErrorNoScrumMembersPresent"), cancellationToken: cancellationToken);
                        return;
                    }

                    var membersActivityIdMap = this.GetActivityIdOfMembersInScrum(scrumMembers);
                    string membersList = JsonConvert.SerializeObject(membersActivityIdMap);

                    // Mentioning the participants involved in the scrum
                    var mentionActivity = this.GetMentionsActivity(scrumMembers);

                    // Check if channel exists. If channel doesn't exist then scrum card will be sent in General channel.
                    scrumConfiguration.ChannelId = await this.GetValidChannelIdAsync(turnContext, scrumConfiguration);

                    // Send the start scrum card
                    turnContext.Activity.Conversation.Id = scrumConfiguration.ChannelId;
                    var attachment = ScrumCard.GetScrumStartCard(scrumSummary, membersActivityIdMap, scrumTeamConfigId, scrumStartActivityId, this.localizer, scrumConfiguration.TimeZone);
                    var scrumStartActivity = MessageFactory.Attachment(attachment);
                    var scrumStartActivityResponse = await turnContext.SendActivityAsync(scrumStartActivity, cancellationToken);

                    // Update the conversation id to send mentioned participants as reply to scrum start card.
                    turnContext.Activity.Conversation = new ConversationAccount
                    {
                        Id = $"{scrumConfiguration.ChannelId};messageid={scrumStartActivityResponse.Id}",
                    };
                    await turnContext.SendActivityAsync(mentionActivity, cancellationToken);
                    await this.CreateScrumAsync(scrumStartActivityResponse.Id, scrumStartActivityId, membersList, scrumConfiguration, turnContext, cancellationToken);
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
        /// Get valid members in scrum and compares whether those members exist in the team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="userPrincipalNames">User principal names who are currently part of the scrum, separated by comma.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Returns list of valid members in scrum.</returns>
        private async Task<IEnumerable<TeamsChannelAccount>> GetValidMembersInScrumAsync(ITurnContext turnContext, string userPrincipalNames, CancellationToken cancellationToken)
        {
            var teamsChannelAccounts = new List<TeamsChannelAccount>();
            string continuationToken = null;
            do
            {
                var currentPage = await TeamsInfo.GetPagedMembersAsync(turnContext, 100, continuationToken, cancellationToken);
                continuationToken = currentPage.ContinuationToken;
                teamsChannelAccounts.AddRange(currentPage.Members);
            }
            while (continuationToken != null);

            if (!string.IsNullOrEmpty(userPrincipalNames))
            {
                var scrumMembers = userPrincipalNames.Split(',').Where(email => !string.IsNullOrEmpty(email));
                return teamsChannelAccounts.Where(member => scrumMembers.Any(user => user.Equals(member.UserPrincipalName, StringComparison.OrdinalIgnoreCase))).ToList();
            }

            return null;
        }

        /// <summary>
        ///  Create a new scrum from the input.
        /// </summary>
        /// <param name="scrumStartCardResponseId">Activity id of scrum summary card.</param>
        /// <param name="scrumCardId">Activity Id of scrum card.</param>
        /// <param name="members">JSON serialized member and activity mapping.</param>
        /// <param name="scrumConfiguration">An instance of scrum configuration details.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task CreateScrumAsync(string scrumStartCardResponseId, string scrumCardId, string members, ScrumConfiguration scrumConfiguration, ITurnContext turnContext, CancellationToken cancellationToken)
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
                    ScrumTeamConfigId = scrumConfiguration.ScrumTeamConfigId,
                    ScrumId = conversationId,
                    ChannelName = scrumConfiguration.ChannelName,
                    TeamId = scrumConfiguration.TeamId,
                    CreatedOn = DateTime.UtcNow.ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.InvariantCulture),
                    AadGroupId = scrumConfiguration.AadGroupId,
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
        /// <param name="scrumMembers">Scrum members present in the current scrum.</param>
        /// <returns>member ids.</returns>
        private Activity GetMentionsActivity(IEnumerable<TeamsChannelAccount> scrumMembers)
        {
            StringBuilder membersMention = new StringBuilder();
            var entities = new List<Entity>();
            var mentions = new List<Mention>();
            foreach (var member in scrumMembers)
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
        /// Method to get unique id for each user present in the scrum.
        /// </summary>
        /// <param name="scrumMembers">Scrum members present in the current scrum.</param>
        /// <returns>Returns member ids as collection.</returns>
        private Dictionary<string, string> GetActivityIdOfMembersInScrum(IEnumerable<TeamsChannelAccount> scrumMembers)
        {
            var membersActivityIdMap = new Dictionary<string, string>();

            if (scrumMembers == null)
            {
                return null;
            }

            foreach (var member in scrumMembers)
            {
                membersActivityIdMap[member.Id] = Guid.NewGuid().ToString();
            }

            return membersActivityIdMap;
        }

        /// <summary>
        /// Get general channel Id if scrum channel id does not exist.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="scrumConfigurationDetails">scrum configuration details.</param>
        /// <returns>Returns general channel Id if scrum channel id does not exist.</returns>
        private async Task<string> GetValidChannelIdAsync(ITurnContext turnContext, ScrumConfiguration scrumConfigurationDetails)
        {
            var teamsChannelInfo = await TeamsInfo.GetTeamChannelsAsync(turnContext, scrumConfigurationDetails.TeamId, CancellationToken.None);
            var channelInfo = teamsChannelInfo.Where(channel => channel.Id.Equals(scrumConfigurationDetails.ChannelId, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

            if (channelInfo == null)
            {
                scrumConfigurationDetails.ChannelId = scrumConfigurationDetails.TeamId;
                scrumConfigurationDetails.ChannelName = Strings.GeneralChannel;

                List<ScrumConfiguration> scrumConfigurations = new List<ScrumConfiguration>();
                scrumConfigurations.Add(scrumConfigurationDetails);

                var saveResponse = await this.scrumConfigurationStorageProvider.StoreOrUpdateScrumConfigurationEntitiesAsync(scrumConfigurations);
                if (!saveResponse)
                {
                    this.logger.LogError("Error while saving scrum configuration details");
                }

                return scrumConfigurationDetails.TeamId;
            }

            return scrumConfigurationDetails.ChannelId;
        }
    }
}
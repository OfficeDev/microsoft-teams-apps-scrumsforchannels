// <copyright file="ScrumConfigurationController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.ScrumStatus.Common;
    using Microsoft.Teams.Apps.ScrumStatus.Helpers;
    using Microsoft.Teams.Apps.ScrumStatus.Models;

    /// <summary>
    /// Controller to handle Scrum API operations.
    /// </summary>
    [Route("api/scrumconfiguration")]
    [ApiController]
    [Authorize]
    public class ScrumConfigurationController : BaseScrumStatusController
    {
        /// <summary>
        /// Microsoft Application ID.
        /// </summary>
        private readonly string appId;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Scrum For Channels bot adapter to get context.
        /// </summary>
        private readonly BotFrameworkAdapter botAdapter;

        /// <summary>
        /// Provider to store scrum configuration details in Azure Table Storage.
        /// </summary>
        private readonly IScrumConfigurationStorageProvider scrumConfigurationStorageProvider;

        /// <summary>
        /// Instance of class that handles scrum helper methods.
        /// </summary>
        private readonly ScrumHelper scrumHelper;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<ScrumStatusActivityHandlerOptions> options;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrumConfigurationController"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="botAdapter">Scrum For Channels bot adapter.</param>
        /// <param name="microsoftAppCredentials">Microsoft Application credentials for Bot/ME.</param>
        /// <param name="scrumConfigurationStorageProvider">Provider to store scrum configuration details in Azure Table Storage.</param>
        /// <param name="scrumHelper">Instance of class that handles scrum helper methods.</param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        public ScrumConfigurationController(
            ILogger<ScrumConfigurationController> logger,
            BotFrameworkAdapter botAdapter,
            MicrosoftAppCredentials microsoftAppCredentials,
            IScrumConfigurationStorageProvider scrumConfigurationStorageProvider,
            ScrumHelper scrumHelper,
            IOptions<ScrumStatusActivityHandlerOptions> options)
            : base()
        {
            this.logger = logger;
            this.botAdapter = botAdapter;
            this.appId = microsoftAppCredentials != null ? microsoftAppCredentials.MicrosoftAppId : throw new ArgumentNullException(nameof(microsoftAppCredentials));
            this.scrumConfigurationStorageProvider = scrumConfigurationStorageProvider;
            this.scrumHelper = scrumHelper;
            this.options = options ?? throw new ArgumentNullException(nameof(options));
        }

        /// <summary>
        /// Get team members and channel information by team id.
        /// </summary>
        /// <param name="teamId">Unique identifier of the team in which BOT is installed.</param>
        /// <returns>List of members in team.</returns>
        [HttpGet("teamdetails")]
        public async Task<IActionResult> GetTeamDetailsAsync(string teamId)
        {
            try
            {
                var userClaims = this.GetUserClaims();

                var teamsChannelAccounts = new List<TeamsChannelAccount>();
                IEnumerable<ChannelInfo> teamsChannelInfo = new List<ChannelInfo>();
                var conversationReference = new ConversationReference
                {
                    ChannelId = teamId,
                    ServiceUrl = userClaims.ServiceUrl,
                    Bot = new ChannelAccount() { Id = $"28:{this.appId}" },
                    Conversation = new ConversationAccount() { ConversationType = Constants.ChannelConversationType, IsGroup = true, Id = teamId, TenantId = this.options.Value.TenantId },
                };

                await this.botAdapter.ContinueConversationAsync(
                    this.appId,
                    conversationReference,
                    async (context, token) =>
                    {
                        string continuationToken = null;
                        do
                        {
                            var currentPage = await TeamsInfo.GetPagedMembersAsync(context, 100, continuationToken, CancellationToken.None);
                            continuationToken = currentPage.ContinuationToken;
                            teamsChannelAccounts.AddRange(currentPage.Members);
                        }
                        while (continuationToken != null);

                        teamsChannelInfo = await TeamsInfo.GetTeamChannelsAsync(context, teamId, CancellationToken.None);
                    },
                    default);

                this.logger.LogInformation("GET call for fetching team members and channels from team roster is successful");
                teamsChannelInfo.First(channel => channel.Name == null).Name = Strings.GeneralChannel;

                var teamDetails = new TeamDetail
                {
                    TeamMembers = teamsChannelAccounts
                    .Select(
                        member => new TeamMember
                        {
                            Content = member.Email,
                            Header = member.Name,
                            AzureAdObjectId = member.AadObjectId,
                        }),
                    Channels = teamsChannelInfo
                    .Select(
                        member => new TeamAccount
                        {
                            ChannelId = member.Id,
                            Header = member.Name,
                        }),
                };

                return this.Ok(teamDetails);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error occurred while getting team member list.");
                throw;
            }
        }

        /// <summary>
        /// Get scrum configuration details by Azure Active Directory group Id.
        /// </summary>
        /// <param name="groupId">Azure Active Directory group Id.</param>
        /// <returns>Returns scrum configuration details.</returns>
        [HttpGet("scrumconfigurationdetails")]
        public async Task<IActionResult> GetScrumConfigurationDetailsByAADGroupIDAsync([FromQuery]string groupId)
        {
            try
            {
                this.logger.LogInformation("Initiated call for fetching scrum configuration details from storage");
                var scrumConfigurationDetails = await this.scrumConfigurationStorageProvider.GetScrumConfigurationDetailsbyAADGroupIDAsync(groupId);
                this.logger.LogInformation("GET call for fetching scrum configuration details from storage is successful");
                return this.Ok(scrumConfigurationDetails);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while getting scrum configuration details.");
                throw;
            }
        }

        /// <summary>
        /// Get system time zones information.
        /// </summary>
        /// <returns>Returns time zone details.</returns>
        [HttpGet("timezoneinfo")]
        public async Task<IActionResult> GetTimeZoneInfoAsync()
        {
            try
            {
                this.logger.LogInformation("Initiated call for fetching time zone information.");
                var timeZoneInfo = TimeZoneInfo.GetSystemTimeZones();
                var timeZoneDetails = timeZoneInfo.Select(timeZone => new { timeZoneId = timeZone.Id, header = timeZone.DisplayName });
                this.logger.LogInformation("GET call for fetching time zone information is successful");
                return this.Ok(timeZoneDetails);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while getting time zone information.");
                throw;
            }
        }

        /// <summary>
        /// Post call to save scrum configuration details in Azure Table storage.
        /// </summary>
        /// <param name="scrumConfigurationData">Class contains details of scrum configuration details.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost("scrumconfigurationdetails")]
        public async Task<IActionResult> SaveScrumConfigurationDetailsAsync([FromBody]IEnumerable<ScrumConfiguration> scrumConfigurationData)
        {
            try
            {
                if (scrumConfigurationData == null)
                {
                    return this.BadRequest("The scrum configuration record is found as null.");
                }

                this.logger.LogInformation("Initiated call to scrum configuration storage provider.");
                scrumConfigurationData = this.scrumHelper.ConstructScrumConfigurationEntities(scrumConfigurationData)?.ToList();
                var result = await this.scrumConfigurationStorageProvider.StoreOrUpdateScrumConfigurationEntitiesAsync(scrumConfigurationData);
                this.logger.LogInformation("POST call for saving scrum configuration details is successful");
                return this.Ok(result);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while saving scrum configuration details.");
                throw;
            }
        }

        /// <summary>
        /// Put call to delete specified scrum configuration details from Azure Table storage.
        /// </summary>
        /// <param name="scrumConfigurationDataToBeDeleted">Class contains details of scrum configuration details to be deleted.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpDelete("scrumconfigurationdetails")]
        public async Task<IActionResult> DeleteScrumConfigurationDetailsAsync([FromBody]IEnumerable<ScrumConfiguration> scrumConfigurationDataToBeDeleted)
        {
            try
            {
                if (scrumConfigurationDataToBeDeleted == null)
                {
                    return this.BadRequest("No data received to be deleted from Microsoft Azure Table storage");
                }

                this.logger.LogInformation("Initiated call to scrum configuration storage provider service to delete scrum configuration details.");
                var result = await this.scrumConfigurationStorageProvider.DeleteScrumConfigurationDetailsAsync(scrumConfigurationDataToBeDeleted);
                this.logger.LogInformation("PUT call for deleting scrum configuration details in storage is successful");
                return this.Ok(result);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while deleting scrum configuration details.");
                throw;
            }
        }
    }
}
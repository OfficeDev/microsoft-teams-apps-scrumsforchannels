// <copyright file="ScrumMasterController.cs" company="Microsoft">
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
    [Route("api/scrummaster")]
    [ApiController]
    [Authorize]
    public class ScrumMasterController : BaseScrumStatusController
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
        /// Scrum Status Bot adapter to get context.
        /// </summary>
        private readonly BotFrameworkAdapter botAdapter;

        /// <summary>
        /// Provider to store scrum master details in Azure Table Storage.
        /// </summary>
        private readonly IScrumMasterStorageProvider scrumMasterStorageProvider;

        /// <summary>
        /// Instance of class that handles scrum helper methods.
        /// </summary>
        private readonly ScrumHelper scrumHelper;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<ScrumStatusActivityHandlerOptions> options;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrumMasterController"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="botAdapter">Scrum Status bot adapter.</param>
        /// <param name="microsoftAppCredentials">Microsoft Application credentials for Bot/ME.</param>
        /// <param name="scrumMasterStorageProvider">Provider to store scrum master details in Azure Table Storage.</param>
        /// <param name="scrumHelper">Instance of class that handles scrum helper methods.</param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        public ScrumMasterController(
            ILogger<ScrumMasterController> logger,
            BotFrameworkAdapter botAdapter,
            MicrosoftAppCredentials microsoftAppCredentials,
            IScrumMasterStorageProvider scrumMasterStorageProvider,
            ScrumHelper scrumHelper,
            IOptions<ScrumStatusActivityHandlerOptions> options)
            : base()
        {
            this.logger = logger;
            this.botAdapter = botAdapter;
            this.appId = microsoftAppCredentials != null ? microsoftAppCredentials.MicrosoftAppId : throw new ArgumentNullException(nameof(microsoftAppCredentials));
            this.scrumMasterStorageProvider = scrumMasterStorageProvider;
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
                    Bot = new ChannelAccount() { Id = this.appId },
                    Conversation = new ConversationAccount() { ConversationType = Constants.ConversationType, IsGroup = true, Id = teamId, TenantId = this.options.Value.TenantId },
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
                teamsChannelInfo.First(channel => channel.Name == null).Name = Constants.GeneralChannel;
                var teamDetails = new
                {
                    TeamMembers = teamsChannelAccounts.Select(member => new { content = member.Email, header = member.Name, aadobjectid = member.AadObjectId }),
                    Channels = teamsChannelInfo.Select(member => new { ChannelId = member.Id, header = member.Name }),
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
        /// Get scrum master details by Azure Active Directory group Id.
        /// </summary>
        /// <param name="groupId">Azure Active Directory group Id.</param>
        /// <returns>Returns scrum master details.</returns>
        [HttpGet("scrummasterdetails")]
        public async Task<IActionResult> GetScrumMasterDetailsByAADGroupIDAsync([FromQuery]string groupId)
        {
            try
            {
                this.logger.LogInformation("Initiated call for fetching scrum master details from storage");
                var scrumMasterDetails = await this.scrumMasterStorageProvider.GetScrumMasterDetailsbyAADGroupIDAsync(groupId);
                this.logger.LogInformation("GET call for fetching scrum master details from storage is successful");
                return this.Ok(scrumMasterDetails);
            }
            #pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                this.logger.LogError(ex, "Error while getting scrum master details.");
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
            #pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                this.logger.LogError(ex, "Error while getting time zone information.");
                throw;
            }
        }

        /// <summary>
        /// Post call to save scrum master details in Azure Table storage.
        /// </summary>
        /// <param name="scrumMastersData">Class contains details of scrum master details.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost("scrummasterdetails")]
        public async Task<IActionResult> SaveScrumMasterDetailsAsync([FromBody]IEnumerable<ScrumMaster> scrumMastersData)
        {
            try
            {
                if (scrumMastersData == null)
                {
                    return this.BadRequest();
                }

                this.logger.LogInformation("Initiated call to scrum master storage provider.");
                scrumMastersData = this.scrumHelper.GetScrumMasterEntities(scrumMastersData)?.ToList();
                var result = await this.scrumMasterStorageProvider.StoreOrUpdateScrumMasterEntitiesAsync(scrumMastersData);
                this.logger.LogInformation("POST call for saving scrum master details in storage is successful");
                return this.Ok(result);
            }
            #pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                this.logger.LogError(ex, "Error while saving scrum master details.");
                throw;
            }
        }

        /// <summary>
        /// Put call to delete specified scrum master details from Azure Table storage.
        /// </summary>
        /// <param name="scrumMastersDataToBeDeleted">Class contains details of scrum master details to be deleted.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpDelete("scrummasterdetails")]
        public async Task<IActionResult> DeleteScrumMasterDetailsAsync([FromBody]IEnumerable<ScrumMaster> scrumMastersDataToBeDeleted)
        {
            try
            {
                if (scrumMastersDataToBeDeleted == null)
                {
                    return this.BadRequest();
                }

                this.logger.LogInformation("Initiated call to scrum master storage provider service to delete scrum master details.");
                var result = await this.scrumMasterStorageProvider.DeleteScrumMasterDetailsAsync(scrumMastersDataToBeDeleted);
                this.logger.LogInformation("PUT call for deleting scrum master details in storage is successful");
                return this.Ok(result);
            }
            #pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                this.logger.LogError(ex, "Error while deleting scrum master details.");
                throw;
            }
        }
    }
}
// <copyright file="ScrumHelper.cs" company="Microsoft">
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
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.ScrumStatus.Common;
    using Microsoft.Teams.Apps.ScrumStatus.Models;

    /// <summary>
    /// Instance of class that handles scrum helper methods.
    /// </summary>
    public class ScrumHelper
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<ScrumHelper> logger;

        /// <summary>
        /// Storage helper for working with scrum data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumStorageProvider scrumStorageProvider;

        /// <summary>
        /// Storage helper for working with scrum status data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumStatusStorageProvider scrumStatusStorageProvider;

        /// <summary>
        /// Storage helper for working with scrum master data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumMasterStorageProvider scrumMasterStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrumHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="scrumStorageProvider">Scrum storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="scrumStatusStorageProvider">Scrum status storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="scrumMasterStorageProvider">Scrum master storage provider to maintain data in Microsoft Azure table storage.</param>
        public ScrumHelper(
            ILogger<ScrumHelper> logger,
            IScrumStorageProvider scrumStorageProvider,
            IScrumStatusStorageProvider scrumStatusStorageProvider,
            IScrumMasterStorageProvider scrumMasterStorageProvider)
        {
            this.logger = logger;
            this.scrumStorageProvider = scrumStorageProvider;
            this.scrumStatusStorageProvider = scrumStatusStorageProvider;
            this.scrumMasterStorageProvider = scrumMasterStorageProvider;
        }

        /// <summary>
        /// Get scrum summary from table storage.
        /// </summary>
        /// <param name="scrumMasterId">Unique identifier for scrum master details.</param>
        /// <param name="scrumCardResponseId">Scrum card response id.</param>
        /// <param name="membersActivityIdMap">Members id who are part of the scrum.</param>
        /// <returns>A task that represents the scrum summary data needs to be returned.</returns>
        public async Task<ScrumSummary> GetScrumSummaryAsync(string scrumMasterId, string scrumCardResponseId = null, Dictionary<string, string> membersActivityIdMap = null)
        {
            var scrumSummary = new ScrumSummary();

            var scrumMaster = await this.scrumMasterStorageProvider.GetScrumMasterDetailsByScrumMasterIdAsync(scrumMasterId);
            if (scrumMaster == null)
            {
                this.logger.LogInformation("Scrum master details obtained from ScrumMaster table is null");
                return null;
            }

            if (string.IsNullOrEmpty(scrumCardResponseId))
            {
                scrumSummary.ScrumName = string.IsNullOrEmpty(scrumMaster.TeamName) ? "General " : scrumMaster.TeamName;
                scrumSummary.ScrumStartTime = string.Format(CultureInfo.CurrentCulture, "{0:MM/dd/yy HH:mm}", DateTime.UtcNow);
                scrumSummary.TotalUserCount = scrumMaster.UserPrincipalNames.Split(',').Where(email => !string.IsNullOrEmpty(email)).Count();
                scrumSummary.RespondedUserCount = 0;
                scrumSummary.BlockedUsersCount = 0;
                scrumSummary.ScrumRunningStatus = ScrumRunningStatus.Active;
                return scrumSummary;
            }

            var scrum = (await this.scrumStorageProvider.GetScrumDetailsBySummaryCardActivityIdAsync(scrumCardResponseId)).FirstOrDefault();
            if (scrum == null)
            {
                this.logger.LogInformation("Scrum details obtained from Scrum table is null");
                return null;
            }

            var scrumStatus = (await this.scrumStatusStorageProvider.GetScrumStatusBySummaryCardIdAsync(scrumCardResponseId))?.ToList();
            if (scrumStatus != null && scrumStatus.Count > 0)
            {
                scrumSummary.RespondedUserCount = scrumStatus.Count;
                scrumSummary.BlockedUsersCount = scrumStatus.Where(scrumDetails => !string.IsNullOrEmpty(scrumDetails.BlockerDescription) && !string.IsNullOrWhiteSpace(scrumDetails.BlockerDescription)).ToList().Count;
            }
            else
            {
                scrumSummary.RespondedUserCount = 0;
                scrumSummary.BlockedUsersCount = 0;
            }

            scrumSummary.ScrumName = string.IsNullOrEmpty(scrumMaster.TeamName) ? Constants.GeneralChannel : scrumMaster.TeamName;
            scrumSummary.ScrumStartTime = DateTimeOffset.Parse(scrum.CreatedOn, CultureInfo.InvariantCulture).ToString("MM/dd/yy HH:mm", CultureInfo.InvariantCulture);
            scrumSummary.TotalUserCount = membersActivityIdMap == null ? 0 : membersActivityIdMap.Count;
            scrumSummary.ScrumRunningStatus = scrum.IsCompleted ? ScrumRunningStatus.Closed : ScrumRunningStatus.Active;
            return scrumSummary;
        }

        /// <summary>
        /// Get active scrum by scrum master id.
        /// </summary>
        /// <param name="scrumMasterId">Scrum master id</param>
        /// <returns>A task that represents the scrum data needs to be returned.</returns>
        public async Task<Scrum> GetActiveScrumAsync(string scrumMasterId)
        {
            Scrum activeScrum = null;
            var result = await this.scrumStorageProvider.GetScrumByScrumMasterIdAsync(scrumMasterId);
            activeScrum = result.Where(scrum => scrum.IsCompleted == false).FirstOrDefault();
            return activeScrum;
        }

        /// <summary>
        /// Gets scrum master table entity to be stored in table storage.
        /// </summary>
        /// <param name="scrumMasterData">Scrum master entities received from client application. </param>
        /// <returns>Returns updated scrum master table storage entities to be saved in storage.</returns>
        public IEnumerable<ScrumMaster> GetScrumMasterEntities(IEnumerable<ScrumMaster> scrumMasterData)
        {
            try
            {
                if (scrumMasterData == null)
                {
                    return null;
                }

                foreach (ScrumMaster scrumMasterDetails in scrumMasterData)
                {
                    scrumMasterDetails.PartitionKey = scrumMasterDetails?.TeamId;
                    scrumMasterDetails.IsActive = scrumMasterDetails?.IsActive ?? false;
                    scrumMasterDetails.CreatedOn = DateTime.UtcNow.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'", CultureInfo.InvariantCulture);
                    scrumMasterDetails.Key = string.IsNullOrEmpty(scrumMasterDetails.Key) ? Guid.NewGuid().ToString() : scrumMasterDetails.Key;
                    scrumMasterDetails.RowKey = scrumMasterDetails.Key;
                    scrumMasterDetails.ScrumMasterId = $"{scrumMasterDetails?.TeamName}_{scrumMasterDetails?.ChannelId}";
                    scrumMasterDetails.StartTime = DateTime.Parse(scrumMasterDetails.StartTime, CultureInfo.InvariantCulture)
                        .ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

                    // Convert start time with user specified time zone to UTC and store UTC hour in storage. Scrum will be scheduled as per UTC hour.
                    TimeZoneInfo userSpecifiedTimeZone = TimeZoneInfo.FindSystemTimeZoneById(scrumMasterDetails.TimeZone);
                    DateTime utcStartTime = TimeZoneInfo.ConvertTimeToUtc(DateTime.Parse(scrumMasterDetails.StartTime, CultureInfo.InvariantCulture), userSpecifiedTimeZone);
                    scrumMasterDetails.StartTimeUTCHour = Convert.ToInt32(utcStartTime.ToString("HH", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
                }

                return scrumMasterData;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in GetScrumMasterEntities", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Stores scrum status details in storage.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="scrumStatus">Scrum status entity to be stored in table storage.</param>
        /// <param name="adaptiveSubmitActionData">Data submitted in task module.</param>
        /// <param name="summaryCardId">Scrum summary card id.</param>
        /// <returns>Returns success or failure on whether data is stored in storage.</returns>
        public async Task<bool> SaveScrumStatusDetailsAsync(ITurnContext<IInvokeActivity> turnContext, ScrumStatus scrumStatus, AdaptiveSubmitActionData adaptiveSubmitActionData, string summaryCardId)
        {
            scrumStatus = scrumStatus ?? throw new ArgumentNullException(nameof(scrumStatus));
            scrumStatus.MembersActivityIdMap = adaptiveSubmitActionData?.ScrumMembers;
            scrumStatus.SummaryCardId = summaryCardId;
            scrumStatus.CreatedOn = DateTime.UtcNow.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'", CultureInfo.CurrentCulture);
            scrumStatus.Username = turnContext?.Activity.From.Name;
            scrumStatus.AadObjectId = turnContext.Activity.From.AadObjectId;
            return await this.scrumStatusStorageProvider.CreateOrUpdateScrumStatusAsync(scrumStatus);
        }

        /// <summary>
        /// Get valid members in scrum and compares whether those members exist in the channel.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="scrumMaster">Scrum master details.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Returns list of valid members in scrum.</returns>
        public async Task<IEnumerable<TeamsChannelAccount>> GetValidMembersInScrumAsync(ITurnContext turnContext, ScrumMaster scrumMaster, CancellationToken cancellationToken)
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

            var validusers = scrumMaster?.UserPrincipalNames.Split(',').Where(email => !string.IsNullOrEmpty(email));
            return teamsChannelAccounts.Where(member => validusers.Any(user => user.Equals(member.UserPrincipalName, StringComparison.OrdinalIgnoreCase))).ToList();
        }
    }
}
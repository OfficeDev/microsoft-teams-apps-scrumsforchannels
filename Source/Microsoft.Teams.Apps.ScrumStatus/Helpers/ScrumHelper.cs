// <copyright file="ScrumHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.ScrumStatus.Common;
    using Microsoft.Teams.Apps.ScrumStatus.Models;

    /// <summary>
    /// Instance of class that handles scrum helper methods.
    /// </summary>
    public class ScrumHelper
    {
        /// <summary>
        /// Date time format for scrum start time.
        /// </summary>
        private const string ScrumStartDateTimeFormat = "MM/dd/yy HH:mm";

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
        /// Storage helper for working with scrum configuration data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumConfigurationStorageProvider scrumConfigurationStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrumHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="scrumStorageProvider">Scrum storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="scrumStatusStorageProvider">Scrum status storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="scrumConfigurationStorageProvider">Scrum configuration storage provider to maintain data in Microsoft Azure table storage.</param>
        public ScrumHelper(
            ILogger<ScrumHelper> logger,
            IScrumStorageProvider scrumStorageProvider,
            IScrumStatusStorageProvider scrumStatusStorageProvider,
            IScrumConfigurationStorageProvider scrumConfigurationStorageProvider)
        {
            this.logger = logger;
            this.scrumStorageProvider = scrumStorageProvider;
            this.scrumStatusStorageProvider = scrumStatusStorageProvider;
            this.scrumConfigurationStorageProvider = scrumConfigurationStorageProvider;
        }

        /// <summary>
        /// Get scrum summary from table storage.
        /// </summary>
        /// <param name="scrumTeamConfigId">Scrum team configuration details describing scrum team name and channel id.</param>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <param name="scrumCardResponseId">Scrum card response id.</param>
        /// <param name="membersActivityIdMap">Members id who are part of the scrum.</param>
        /// <returns>A task that represents the scrum summary data needs to be returned.</returns>
        public async Task<ScrumSummary> GetScrumSummaryAsync(string scrumTeamConfigId, string aadGroupId, string scrumCardResponseId = null, Dictionary<string, string> membersActivityIdMap = null)
        {
            var scrumSummary = new ScrumSummary();

            var scrumConfiguration = await this.scrumConfigurationStorageProvider.GetScrumConfigurationDetailByScrumTeamConfigIdAsync(scrumTeamConfigId, aadGroupId);
            if (scrumConfiguration == null)
            {
                this.logger.LogInformation("scrum configuration details obtained from ScrumConfiguration table is null");
                return null;
            }

            if (string.IsNullOrEmpty(scrumCardResponseId))
            {
                scrumSummary.ScrumName = string.IsNullOrEmpty(scrumConfiguration.ScrumTeamName) ? Strings.GeneralChannel : scrumConfiguration.ScrumTeamName;
                scrumSummary.ScrumStartTime = string.Format(CultureInfo.CurrentCulture, $"{{0:{ScrumStartDateTimeFormat}}}", DateTime.UtcNow);
                scrumSummary.TotalUserCount = scrumConfiguration.UserPrincipalNames.Split(',').Where(email => !string.IsNullOrEmpty(email)).Count();
                scrumSummary.RespondedUserCount = 0;
                scrumSummary.BlockedUsersCount = 0;
                scrumSummary.ScrumRunningStatus = ScrumRunningStatus.Active;
                return scrumSummary;
            }

            var scrum = (await this.scrumStorageProvider.GetScrumsBySummaryCardActivityIdAsync(scrumCardResponseId, scrumConfiguration.AadGroupId)).FirstOrDefault();
            if (scrum == null)
            {
                this.logger.LogInformation("Scrum details obtained from Scrum table is null");
                return null;
            }

            var scrumStatus = (await this.scrumStatusStorageProvider.GetScrumStatusBySummaryCardIdAsync(scrumCardResponseId, scrumConfiguration.AadGroupId))?.ToList();
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

            scrumSummary.ScrumName = string.IsNullOrEmpty(scrumConfiguration.ScrumTeamName) ? Strings.GeneralChannel : scrumConfiguration.ScrumTeamName;
            scrumSummary.ScrumStartTime = DateTimeOffset.Parse(scrum.CreatedOn, CultureInfo.InvariantCulture).ToString(ScrumStartDateTimeFormat, CultureInfo.InvariantCulture);
            scrumSummary.TotalUserCount = membersActivityIdMap == null ? 0 : membersActivityIdMap.Count;
            scrumSummary.ScrumRunningStatus = scrum.IsCompleted ? ScrumRunningStatus.Closed : ScrumRunningStatus.Active;
            return scrumSummary;
        }

        /// <summary>
        /// Get active scrum by scrum team configuration id.
        /// </summary>
        /// <param name="scrumTeamConfigId">Scrum team configuration id</param>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>A task that represents the scrum data needs to be returned.</returns>
        public async Task<Scrum> GetActiveScrumAsync(string scrumTeamConfigId, string aadGroupId)
        {
            var result = await this.scrumStorageProvider.GetScrumsByScrumTeamConfigIdAsync(scrumTeamConfigId, aadGroupId);
            return result.Where(scrum => !scrum.IsCompleted).FirstOrDefault();
        }

        /// <summary>
        /// Gets scrum configuration table entity to be stored in table storage.
        /// </summary>
        /// <param name="scrumConfigurationData">Scrum configuration entities received from client application. </param>
        /// <returns>Returns updated scrum configuration table storage entities to be saved in storage.</returns>
        public IEnumerable<ScrumConfiguration> ConstructScrumConfigurationEntities(IEnumerable<ScrumConfiguration> scrumConfigurationData)
        {
            scrumConfigurationData = scrumConfigurationData ?? throw new ArgumentNullException(nameof(scrumConfigurationData));

            try
            {
                foreach (ScrumConfiguration scrumConfigurationDetails in scrumConfigurationData)
                {
                    scrumConfigurationDetails.IsActive = scrumConfigurationDetails.IsActive;
                    scrumConfigurationDetails.CreatedOn = DateTime.UtcNow.ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.InvariantCulture);
                    scrumConfigurationDetails.ScrumConfigurationId = string.IsNullOrEmpty(scrumConfigurationDetails.ScrumConfigurationId) ? Guid.NewGuid().ToString() : scrumConfigurationDetails.ScrumConfigurationId;
                    scrumConfigurationDetails.ScrumTeamConfigId = $"{scrumConfigurationDetails.ScrumTeamName}_{scrumConfigurationDetails.ChannelId}";
                    scrumConfigurationDetails.StartTime = DateTime.Parse(scrumConfigurationDetails.StartTime, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind).ToString(CultureInfo.CurrentCulture);

                    // Convert start time with user specified time zone to UTC and store UTC hour in storage. Scrum will be scheduled as per UTC hour.
                    TimeZoneInfo userSpecifiedTimeZone = TimeZoneInfo.FindSystemTimeZoneById(scrumConfigurationDetails.TimeZone);
                    DateTime utcStartTime = TimeZoneInfo.ConvertTimeToUtc(DateTime.Parse(scrumConfigurationDetails.StartTime, CultureInfo.InvariantCulture), userSpecifiedTimeZone);
                    scrumConfigurationDetails.StartTimeUTCHour = Convert.ToInt32(utcStartTime.ToString("HH", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
                }

                return scrumConfigurationData;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in GetScrumConfigurationEntities", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Stores scrum status details in storage.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="scrumStatus">Scrum status entity to be stored in table storage.</param>
        /// <param name="adaptiveSubmitActionData">Data submitted in task module.</param>
        /// <param name="scrumStartCardResponseId">Scrum start card response id.</param>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>Returns success or failure on whether data is stored in storage.</returns>
        public async Task<bool> SaveScrumStatusDetailsAsync(ITurnContext<IInvokeActivity> turnContext, ScrumStatus scrumStatus, AdaptiveSubmitActionData adaptiveSubmitActionData, string scrumStartCardResponseId, string aadGroupId)
        {
            scrumStatus = scrumStatus ?? throw new ArgumentNullException(nameof(scrumStatus));
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

            scrumStatus.MembersActivityIdMap = adaptiveSubmitActionData?.ScrumMembers;
            scrumStatus.ScrumStartCardResponseId = scrumStartCardResponseId;
            scrumStatus.CreatedOn = DateTime.UtcNow.ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.CurrentCulture);
            scrumStatus.Username = turnContext.Activity.From.Name;
            scrumStatus.AadGroupId = aadGroupId;
            scrumStatus.UserAadObjectId = turnContext.Activity.From.AadObjectId;
            return await this.scrumStatusStorageProvider.CreateOrUpdateScrumStatusAsync(scrumStatus);
        }
    }
}
// <copyright file="ScrumStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.ScrumStatus.Common.Models;
    using Microsoft.Teams.Apps.ScrumStatus.Models;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// Scrum storage provider.
    /// </summary>
    public class ScrumStorageProvider : BaseStorageProvider, IScrumStorageProvider
    {
        /// <summary>
        /// Represents scrum entity name.
        /// </summary>
        private const string ScrumEntity = "Scrum";

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<ScrumStorageProvider> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrumStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application configuration properties for Microsoft Azure Table storage.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public ScrumStorageProvider(IOptionsMonitor<StorageOptions> storageOptions, ILogger<ScrumStorageProvider> logger)
            : base(storageOptions, ScrumEntity)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Stores or update scrum data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumData">Holds scrum entity data.</param>
        /// <returns>A task that represents scrum entity data is saved or updated.</returns>
        public async Task<bool> CreateOrUpdateScrumAsync(Scrum scrumData)
        {
            scrumData = scrumData ?? throw new ArgumentNullException(nameof(scrumData));

            try
            {
                Scrum scrumEntity = new Scrum()
                {
                    IsCompleted = scrumData.IsCompleted,
                    ThreadConversationId = scrumData.ThreadConversationId,
                    ScrumStartActivityId = scrumData.ScrumStartActivityId,
                    ScrumStartCardResponseId = scrumData.ScrumStartCardResponseId,
                    MembersActivityIdMap = scrumData.MembersActivityIdMap,
                    ScrumTeamConfigId = scrumData.ScrumTeamConfigId,
                    ScrumId = scrumData.ThreadConversationId,
                    ChannelName = scrumData.ChannelName,
                    TeamId = scrumData.TeamId,
                    CreatedOn = scrumData.CreatedOn,
                    AadGroupId = scrumData.AadGroupId,
                };

                await this.EnsureInitializedAsync();
                TableOperation operation = TableOperation.InsertOrReplace(scrumEntity);
                var result = await this.CloudTable.ExecuteAsync(operation);
                return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in CreateOrUpdateScrumAsync: ScrumTeamConfigId : {JsonConvert.SerializeObject(scrumData)}.", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Get scrum details by summary card activity id from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="summaryCardActivityId">Summary card activity id.</param>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>Returns collection of scrum details by summary card activity id.</returns>
        public async Task<IEnumerable<Scrum>> GetScrumsBySummaryCardActivityIdAsync(string summaryCardActivityId, string aadGroupId)
        {
            summaryCardActivityId = summaryCardActivityId ?? throw new ArgumentNullException(nameof(summaryCardActivityId));

            try
            {
                await this.EnsureInitializedAsync();
                string scrumCardResponseIdfilter = TableQuery.GenerateFilterCondition(nameof(Scrum.ScrumStartCardResponseId), QueryComparisons.Equal, summaryCardActivityId);
                string aadGroupIdfilter = TableQuery.GenerateFilterCondition(nameof(ScrumConfiguration.PartitionKey), QueryComparisons.Equal, aadGroupId);
                var combinedFilter = TableQuery.CombineFilters(scrumCardResponseIdfilter, TableOperators.And, aadGroupIdfilter);
                return await this.GetScrumDetailsAsync(combinedFilter);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in GetScrumsBySummaryCardActivityIdAsync: summaryCardActivityId: {summaryCardActivityId}.", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Get scrum by scrum team configuration id from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumTeamConfigId">Unique identifier for scrum configuration details.</param>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>A <see cref="Task{TResult}"/>Representing the result of the asynchronous operation.</returns>
        public async Task<IEnumerable<Scrum>> GetScrumsByScrumTeamConfigIdAsync(string scrumTeamConfigId, string aadGroupId)
        {
            scrumTeamConfigId = scrumTeamConfigId ?? throw new ArgumentNullException(nameof(scrumTeamConfigId));

            try
            {
                await this.EnsureInitializedAsync();
                string scrumTeamConfigIdfilter = TableQuery.GenerateFilterCondition(nameof(Scrum.ScrumTeamConfigId), QueryComparisons.Equal, scrumTeamConfigId);
                string aadGroupIdfilter = TableQuery.GenerateFilterCondition(nameof(ScrumConfiguration.PartitionKey), QueryComparisons.Equal, aadGroupId);
                var combinedFilter = TableQuery.CombineFilters(scrumTeamConfigIdfilter, TableOperators.And, aadGroupIdfilter);
                return await this.GetScrumDetailsAsync(combinedFilter);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in GetScrumByScrumTeamConfigIdAsync: ScrumTeamConfigId: {scrumTeamConfigId}. {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Get scrum data from Microsoft Azure Table storage based on scrum start activity id.
        /// </summary>
        /// <param name="scrumStartActivityId">Scrum start activity id.</param>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>A task that represent object to hold user profile card activity id and user profile card id.</returns>
        public async Task<IEnumerable<Scrum>> GetScrumsByScrumStartActivityIdAsync(string scrumStartActivityId, string aadGroupId)
        {
            scrumStartActivityId = scrumStartActivityId ?? throw new ArgumentNullException(nameof(scrumStartActivityId));

            try
            {
                await this.EnsureInitializedAsync();
                string scrumStartActivityIdFilter = TableQuery.GenerateFilterCondition(nameof(Scrum.ScrumStartActivityId), QueryComparisons.Equal, scrumStartActivityId);
                string aadGroupIdfilter = TableQuery.GenerateFilterCondition(nameof(ScrumConfiguration.PartitionKey), QueryComparisons.Equal, aadGroupId);
                var combinedFilter = TableQuery.CombineFilters(scrumStartActivityIdFilter, TableOperators.And, aadGroupIdfilter);
                return await this.GetScrumDetailsAsync(combinedFilter);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in GetScrumByScrumStartActivityIdAsync: ScrumTeamConfigId: {scrumStartActivityId}. {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Get Scrum details by time stamp.
        /// </summary>
        /// <returns>task</returns>
        public async Task<IEnumerable<Scrum>> GetScrumDetailsByTimestampAsync()
        {
            try
            {
                await this.EnsureInitializedAsync();
                var date = DateTime.UtcNow.AddDays(-60);
                string filter = TableQuery.GenerateFilterConditionForDate(nameof(Scrum.Timestamp), QueryComparisons.LessThan, date);
                return await this.GetScrumDetailsAsync(filter);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "An error occurred in GetScrumDetailsByTimestampAsync", SeverityLevel.Error);
                return null;
            }
        }

        /// <summary>
        /// Delete scrum status entity from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrum">Holds scrum entity data.</param>
        /// <returns>Delete operation response.</returns>
        public async Task<TableResult> DeleteEntityAsync(Scrum scrum)
        {
            await this.EnsureInitializedAsync();
            TableOperation deleteOperation = TableOperation.Delete(scrum);
            return await this.CloudTable.ExecuteAsync(deleteOperation);
        }

        /// <summary>
        /// Get scrum collection from Microsoft Azure Table storage depending upon filter condition.
        /// </summary>
        /// <param name="filter">Filter condition to fetch data from storage.</param>
        /// <returns>Returns collection of scrum details from storage.</returns>
        private async Task<IEnumerable<Scrum>> GetScrumDetailsAsync(string filter)
        {
            TableQuery<Scrum> query = new TableQuery<Scrum>().Where(filter);
            TableContinuationToken continuationToken = null;
            var scrumCollection = new List<Scrum>();

            do
            {
                var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                if (queryResult?.Results != null)
                {
                    scrumCollection.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return scrumCollection;
        }
    }
}

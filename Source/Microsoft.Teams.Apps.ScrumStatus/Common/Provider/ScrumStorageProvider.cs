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
            try
            {
                Scrum scrumEntity = new Scrum()
                {
                    PartitionKey = scrumData?.ThreadConversationId,
                    RowKey = $"{scrumData.ThreadConversationId}_{scrumData.TeamId}",
                    IsCompleted = scrumData.IsCompleted,
                    ThreadConversationId = scrumData.ThreadConversationId,
                    ScrumStartActivityId = scrumData.ScrumStartActivityId,
                    ScrumStartCardResponseId = scrumData.ScrumStartCardResponseId,
                    MembersActivityIdMap = scrumData.MembersActivityIdMap,
                    ScrumMasterId = scrumData.ScrumMasterId,
                    ScrumId = scrumData.ThreadConversationId,
                    ChannelName = scrumData.ChannelName,
                    TeamId = scrumData.TeamId,
                    CreatedOn = scrumData.CreatedOn,
                    AADGroupID = scrumData.AADGroupID,
                };

                await this.EnsureInitializedAsync();
                TableOperation operation = TableOperation.InsertOrReplace(scrumEntity);
                var result = await this.CloudTable.ExecuteAsync(operation);
                return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in CreateOrUpdateScrumAsync: ScrumMasterId : {JsonConvert.SerializeObject(scrumData)}.", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Get scrum details by summary card activity id from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="summaryCardActivityId">Summary card activity id.</param>
        /// <returns>Returns collection of scrum details by summary card activity id.</returns>
        public async Task<IEnumerable<Scrum>> GetScrumDetailsBySummaryCardActivityIdAsync(string summaryCardActivityId)
        {
            if (string.IsNullOrEmpty(summaryCardActivityId))
            {
                return null;
            }

            try
            {
                await this.EnsureInitializedAsync();
                string filter = TableQuery.GenerateFilterCondition(nameof(Scrum.ScrumStartCardResponseId), QueryComparisons.Equal, summaryCardActivityId);
                var query = new TableQuery<Scrum>().Where(filter);
                var result = await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
                return result;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in GetScrumDetailsBySummaryCardActivityIdAsync: summaryCardActivityId: {summaryCardActivityId}.", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Get scrum by scrum master id from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumMasterId">Unique identifier for scrum master details.</param>
        /// <returns>A <see cref="Task{TResult}"/>Representing the result of the asynchronous operation.</returns>
        public async Task<IEnumerable<Scrum>> GetScrumByScrumMasterIdAsync(string scrumMasterId)
        {
            if (string.IsNullOrEmpty(scrumMasterId))
            {
                return null;
            }

            try
            {
                await this.EnsureInitializedAsync();
                string filter = TableQuery.GenerateFilterCondition(nameof(Scrum.ScrumMasterId), QueryComparisons.Equal, scrumMasterId);
                var query = new TableQuery<Scrum>().Where(filter);
                return await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in GetScrumByScrumMasterIdAsync: ScrumMasterId: {scrumMasterId}. {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Get scrum data from Microsoft Azure Table storage based on scrum start activity id.
        /// </summary>
        /// <param name="scrumStartActivityId">Scrum start activity id.</param>
        /// <returns>A task that represent object to hold user profile card activity id and user profile card id.</returns>
        public async Task<IEnumerable<Scrum>> GetScrumByScrumStartActivityIdAsync(string scrumStartActivityId)
        {
            if (string.IsNullOrEmpty(scrumStartActivityId))
            {
                return null;
            }

            try
            {
                await this.EnsureInitializedAsync();
                string filter = TableQuery.GenerateFilterCondition(nameof(Scrum.ScrumStartActivityId), QueryComparisons.Equal, scrumStartActivityId);
                var query = new TableQuery<Scrum>().Where(filter);
                return await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in GetScrumByScrumStartActivityIdAsync: ScrumMasterId: {scrumStartActivityId}. {ex.Message}", SeverityLevel.Error);
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
                var query = new TableQuery<Scrum>().Where(filter);
                var result = await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
                return result;
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
    }
}

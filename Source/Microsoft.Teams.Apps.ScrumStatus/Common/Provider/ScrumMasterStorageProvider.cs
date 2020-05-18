// <copyright file="ScrumMasterStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.ScrumStatus.Common.Models;
    using Microsoft.Teams.Apps.ScrumStatus.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Scrum master storage provider.
    /// </summary>
    public class ScrumMasterStorageProvider : BaseStorageProvider, IScrumMasterStorageProvider
    {
        /// <summary>
        /// Represents scrum status entity name.
        /// </summary>
        private const string ScrumMasterEntity = "ScrumMaster";

        /// <summary>
        /// Max number of scrum for a batch operation.
        /// </summary>
        private const int ScrumsPerBatch = 100;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<ScrumMasterStorageProvider> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrumMasterStorageProvider"/> class.
        /// Handles Microsoft Azure Table storage read write operations.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application configuration properties for Microsoft Azure Table storage.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public ScrumMasterStorageProvider(IOptionsMonitor<StorageOptions> storageOptions, ILogger<ScrumMasterStorageProvider> logger)
            : base(storageOptions, ScrumMasterEntity)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Creates or updates Microsoft Azure Table storage to store scrum master details.
        /// </summary>
        /// <param name="scrumMasterEntities">Entities to be created or updated.</param>
        /// <returns>Boolean result.</returns>
        public async Task<bool> StoreOrUpdateScrumMasterEntitiesAsync(IEnumerable<ScrumMaster> scrumMasterEntities)
        {
            if (scrumMasterEntities != null)
            {
                await this.EnsureInitializedAsync();
                TableBatchOperation tableBatchOperation = new TableBatchOperation();
                int batchCount = (int)Math.Ceiling((double)scrumMasterEntities.Count() / ScrumsPerBatch);
                for (int batchCountIndex = 0; batchCountIndex < batchCount; batchCountIndex++)
                {
                    var scrumMasterEntitiesBatch = scrumMasterEntities.Skip(batchCountIndex * ScrumsPerBatch).Take(ScrumsPerBatch);
                    foreach (var scrumMasterEntity in scrumMasterEntitiesBatch)
                    {
                        tableBatchOperation.InsertOrReplace(scrumMasterEntity);
                    }

                    if (tableBatchOperation.Count > 0)
                    {
                        await this.CloudTable.ExecuteBatchAsync(tableBatchOperation);
                    }
                }

                return true;
            }

            return false;
        }

        /// <summary>
        /// Get scrum master details by scrum master id from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumMasterId">Scrum master id.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<ScrumMaster> GetScrumMasterDetailsByScrumMasterIdAsync(string scrumMasterId)
        {
            if (string.IsNullOrEmpty(scrumMasterId))
            {
                return null;
            }

            try
            {
                await this.EnsureInitializedAsync();
                string filter = TableQuery.GenerateFilterCondition(nameof(ScrumMaster.ScrumMasterId), QueryComparisons.Equal, scrumMasterId);
                var query = new TableQuery<ScrumMaster>().Where(filter);
                var scrumMasterDetails = await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
                return scrumMasterDetails.Results.FirstOrDefault();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in GetScrumMasterDetailsByScrumMasterIdAsync: ScrumMasterId: {scrumMasterId}. {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Get scrum master details which need to be scheduled in next 1 hour by start scrum background service.
        /// </summary>
        /// <returns>Returns collection of scrum master details.</returns>
        public async Task<IEnumerable<ScrumMaster>> GetActiveScrumMasterOfNextHourAsync()
        {
            try
            {
                int nextHour = DateTime.UtcNow.AddHours(1).Hour;
                await this.EnsureInitializedAsync();
                string isActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(ScrumMaster.IsActive), QueryComparisons.Equal, true);
                string timeFilter = TableQuery.GenerateFilterConditionForInt(nameof(ScrumMaster.StartTimeUTCHour), QueryComparisons.Equal, nextHour);
                var query = new TableQuery<ScrumMaster>().Where($"{isActiveFilter} and {timeFilter}");
                return await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in getting scrum master details to be scheduled in next hour.", SeverityLevel.Error);
                return null;
            }
        }

        /// <summary>
        /// Get scrum master details by Azure Active Directory group Id.
        /// </summary>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>Returns collection of scrum master details.</returns>
        public async Task<IEnumerable<ScrumMaster>> GetScrumMasterDetailsbyAADGroupIDAsync(string aadGroupId)
        {
            if (string.IsNullOrEmpty(aadGroupId))
            {
                return null;
            }

            try
            {
                await this.EnsureInitializedAsync();
                string filter = TableQuery.GenerateFilterCondition(nameof(ScrumMaster.AADGroupID), QueryComparisons.Equal, aadGroupId);
                var query = new TableQuery<ScrumMaster>().Where(filter);
                return await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in getting scrum master details by AAD Group ID: AADGroupID: {aadGroupId}.", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Delete an entity from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumMasterEntities">Entities to be deleted.</param>
        /// <returns>Boolean result.</returns>
        public async Task<bool> DeleteScrumMasterDetailsAsync(IEnumerable<ScrumMaster> scrumMasterEntities)
        {
            try
            {
                await this.EnsureInitializedAsync();
                TableBatchOperation tableBatchOperation = new TableBatchOperation();
                int batchCount = (int)Math.Ceiling((double)scrumMasterEntities.Count() / ScrumsPerBatch);
                for (int batchCountIndex = 0; batchCountIndex < batchCount; batchCountIndex++)
                {
                    var scrumMasterEntitiesBatch = scrumMasterEntities.Skip(batchCountIndex * ScrumsPerBatch).Take(ScrumsPerBatch);
                    foreach (var scrumMasterEntity in scrumMasterEntitiesBatch)
                    {
                        tableBatchOperation.Delete(scrumMasterEntity);
                    }

                    if (tableBatchOperation.Count > 0)
                    {
                        await this.CloudTable.ExecuteBatchAsync(tableBatchOperation);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in deleting scrum master details", SeverityLevel.Error);
                throw;
            }
        }
    }
}
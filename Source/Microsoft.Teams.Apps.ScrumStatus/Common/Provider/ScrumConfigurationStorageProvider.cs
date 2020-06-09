// <copyright file="ScrumConfigurationStorageProvider.cs" company="Microsoft">
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
    public class ScrumConfigurationStorageProvider : BaseStorageProvider, IScrumConfigurationStorageProvider
    {
        /// <summary>
        /// Represents scrum status entity name.
        /// </summary>
        private const string ScrumConfigurationEntity = "ScrumConfiguration";

        /// <summary>
        /// Max number of scrum for a batch operation.
        /// </summary>
        private const int ScrumsPerBatch = 100;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<ScrumConfigurationStorageProvider> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrumConfigurationStorageProvider"/> class.
        /// Handles Microsoft Azure Table storage read write operations.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application configuration properties for Microsoft Azure Table storage.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public ScrumConfigurationStorageProvider(IOptionsMonitor<StorageOptions> storageOptions, ILogger<ScrumConfigurationStorageProvider> logger)
            : base(storageOptions, ScrumConfigurationEntity)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Creates or updates Microsoft Azure Table storage to store scrum configuration details.
        /// </summary>
        /// <param name="scrumConfigurationEntities">Entities to be created or updated.</param>
        /// <returns>Boolean result.</returns>
        public async Task<bool> StoreOrUpdateScrumConfigurationEntitiesAsync(IEnumerable<ScrumConfiguration> scrumConfigurationEntities)
        {
            scrumConfigurationEntities = scrumConfigurationEntities ?? throw new ArgumentNullException(nameof(scrumConfigurationEntities));

            await this.EnsureInitializedAsync();
            TableBatchOperation tableBatchOperation = new TableBatchOperation();
            int batchCount = (int)Math.Ceiling((double)scrumConfigurationEntities.Count() / ScrumsPerBatch);
            for (int batchCountIndex = 0; batchCountIndex < batchCount; batchCountIndex++)
            {
                var scrumConfigurationEntitiesBatch = scrumConfigurationEntities.Skip(batchCountIndex * ScrumsPerBatch).Take(ScrumsPerBatch);
                foreach (var scrumConfigurationEntity in scrumConfigurationEntitiesBatch)
                {
                    tableBatchOperation.InsertOrReplace(scrumConfigurationEntity);
                }

                if (tableBatchOperation.Count > 0)
                {
                    await this.CloudTable.ExecuteBatchAsync(tableBatchOperation);
                }
            }

            return true;
        }

        /// <summary>
        /// Get scrum configuration details by scrum team configuration id from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumTeamConfigId">scrum team configuration id.</param>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<ScrumConfiguration> GetScrumConfigurationDetailByScrumTeamConfigIdAsync(string scrumTeamConfigId, string aadGroupId)
        {
            scrumTeamConfigId = scrumTeamConfigId ?? throw new ArgumentNullException(nameof(scrumTeamConfigId));

            try
            {
                await this.EnsureInitializedAsync();
                string scrumTeamConfigIdFilter = TableQuery.GenerateFilterCondition(nameof(ScrumConfiguration.ScrumTeamConfigId), QueryComparisons.Equal, scrumTeamConfigId);
                string aadGroupIdfilter = TableQuery.GenerateFilterCondition(nameof(ScrumConfiguration.PartitionKey), QueryComparisons.Equal, aadGroupId);
                var combinedFilter = TableQuery.CombineFilters(scrumTeamConfigIdFilter, TableOperators.And, aadGroupIdfilter);
                var query = new TableQuery<ScrumConfiguration>().Where(combinedFilter);
                var scrumConfigurationDetails = await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
                return scrumConfigurationDetails.Results.FirstOrDefault();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in GetScrumConfigurationDetailsByScrumTeamConfigIdAsync: ScrumTeamConfigId: {scrumTeamConfigId}. {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Get scrum configuration details which need to be scheduled by current and previous UTC hour by start scrum background service.
        /// Hours are adjusted to honor daylight savings changes when applicable.
        /// </summary>
        /// <returns>Returns collection of scrum configuration details.</returns>
        public async Task<IEnumerable<ScrumConfiguration>> GetActiveScrumConfigurationsByUtcHourAsync()
        {
            try
            {
                int currentUtcHour = DateTime.UtcNow.Hour;
                await this.EnsureInitializedAsync();
                string isActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(ScrumConfiguration.IsActive), QueryComparisons.Equal, true);
                string currentUtcHourFilter = TableQuery.GenerateFilterConditionForInt(nameof(ScrumConfiguration.StartTimeUTCHour), QueryComparisons.Equal, currentUtcHour);

                // This filter is added to handle day light saving scenarios.
                // This will ensure scrum will not get skipped in daylight saving and scheduler will continue to send start scrum card at specified time.
                string previousUtcHourFilter = TableQuery.GenerateFilterConditionForInt(nameof(ScrumConfiguration.StartTimeUTCHour), QueryComparisons.Equal, currentUtcHour - 1);
                var utcHourFilter = TableQuery.CombineFilters(currentUtcHourFilter, TableOperators.Or, previousUtcHourFilter);
                var combinedFilter = TableQuery.CombineFilters(isActiveFilter, TableOperators.And, utcHourFilter);

                TableQuery<ScrumConfiguration> query = new TableQuery<ScrumConfiguration>().Where(combinedFilter);
                TableContinuationToken continuationToken = null;
                var scrumConfigurations = new List<ScrumConfiguration>();

                do
                {
                    var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                    if (queryResult?.Results != null)
                    {
                        scrumConfigurations.AddRange(queryResult.Results);
                        continuationToken = queryResult.ContinuationToken;
                    }
                }
                while (continuationToken != null);

                return scrumConfigurations;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in getting scrum configuration details to be scheduled in next hour.", SeverityLevel.Error);
                return null;
            }
        }

        /// <summary>
        /// Get scrum configuration details by Azure Active Directory group Id.
        /// </summary>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>Returns collection of scrum configuration details.</returns>
        public async Task<IEnumerable<ScrumConfiguration>> GetScrumConfigurationDetailsbyAADGroupIDAsync(string aadGroupId)
        {
            aadGroupId = aadGroupId ?? throw new ArgumentNullException(nameof(aadGroupId));

            try
            {
                await this.EnsureInitializedAsync();
                string filter = TableQuery.GenerateFilterCondition(nameof(ScrumConfiguration.AadGroupId), QueryComparisons.Equal, aadGroupId);

                TableQuery<ScrumConfiguration> query = new TableQuery<ScrumConfiguration>().Where(filter);
                TableContinuationToken continuationToken = null;
                var scrumConfigurations = new List<ScrumConfiguration>();

                do
                {
                    var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                    if (queryResult?.Results != null)
                    {
                        scrumConfigurations.AddRange(queryResult.Results);
                        continuationToken = queryResult.ContinuationToken;
                    }
                }
                while (continuationToken != null);

                return scrumConfigurations;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in getting scrum configuration details by AAD Group ID: AadGroupId: {aadGroupId}.", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Delete an entity from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumConfigurationEntities">Entities to be deleted.</param>
        /// <returns>Boolean result.</returns>
        public async Task<bool> DeleteScrumConfigurationDetailsAsync(IEnumerable<ScrumConfiguration> scrumConfigurationEntities)
        {
            try
            {
                scrumConfigurationEntities = scrumConfigurationEntities ?? throw new ArgumentNullException(nameof(scrumConfigurationEntities));

                await this.EnsureInitializedAsync();

                TableBatchOperation tableBatchOperation = new TableBatchOperation();
                int batchCount = (int)Math.Ceiling((double)scrumConfigurationEntities.Count() / ScrumsPerBatch);
                for (int batchCountIndex = 0; batchCountIndex < batchCount; batchCountIndex++)
                {
                    var scrumConfigurationEntitiesBatch = scrumConfigurationEntities.Skip(batchCountIndex * ScrumsPerBatch).Take(ScrumsPerBatch);
                    foreach (var scrumConfigurationEntity in scrumConfigurationEntitiesBatch)
                    {
                        tableBatchOperation.Delete(scrumConfigurationEntity);
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
                this.logger.LogError(ex, $"An error occurred in deleting scrum configuration details", SeverityLevel.Error);
                throw;
            }
        }
    }
}
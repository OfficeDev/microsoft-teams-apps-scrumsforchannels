// <copyright file="ScrumStatusStorageProvider.cs" company="Microsoft">
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
    /// Implements storage provider which helps in storing, updating, deleting scrum status data in Microsoft Azure Table storage.
    /// </summary>
    public class ScrumStatusStorageProvider : BaseStorageProvider, IScrumStatusStorageProvider
    {
        /// <summary>
        /// Represents scrum status entity name.
        /// </summary>
        private const string ScrumStatusEntity = "ScrumStatus";

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<ScrumStatusStorageProvider> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrumStatusStorageProvider"/> class.
        /// Handles Microsoft Azure Table storage read write operations.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application configuration properties for Microsoft Azure Table storage.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public ScrumStatusStorageProvider(IOptionsMonitor<StorageOptions> storageOptions, ILogger<ScrumStatusStorageProvider> logger)
            : base(storageOptions, ScrumStatusEntity)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Stores or update scrum status data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumStatusData">Holds scrum status entity data.</param>
        /// <returns>A task that represents scrum status entity data is saved or updated.</returns>
        public async Task<bool> CreateOrUpdateScrumStatusAsync(ScrumStatus scrumStatusData)
        {
            try
            {
                ScrumStatus scrumStatusDataEntity = new ScrumStatus()
                {
                    PartitionKey = scrumStatusData?.SummaryCardId.Trim() + "_" + scrumStatusData.Username,
                    TodayTaskDescription = scrumStatusData.TodayTaskDescription,
                    YesterdayTaskDescription = scrumStatusData.YesterdayTaskDescription,
                    BlockerDescription = scrumStatusData.BlockerDescription,
                    SummaryCardId = scrumStatusData.SummaryCardId.Trim(),
                    Username = scrumStatusData.Username,
                    AadObjectId = scrumStatusData.AadObjectId,
                    MembersActivityIdMap = scrumStatusData.MembersActivityIdMap,
                    RowKey = scrumStatusData.SummaryCardId.Trim(),
                    CreatedOn = scrumStatusData.CreatedOn,
                };

                await this.EnsureInitializedAsync();
                TableOperation operation = TableOperation.InsertOrReplace(scrumStatusDataEntity);
                var result = await this.CloudTable.ExecuteAsync(operation);
                return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in CreateOrUpdateScrumStatusAsync: {JsonConvert.SerializeObject(scrumStatusData)}. {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Get scrum status by summary card id from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="summaryCardId">Scrum summary response card Id.</param>
        /// <returns>Returns collection of scrum status details.</returns>
        public async Task<IEnumerable<ScrumStatus>> GetScrumStatusBySummaryCardIdAsync(string summaryCardId)
        {
            if (string.IsNullOrEmpty(summaryCardId))
            {
                return null;
            }

            try
            {
                await this.EnsureInitializedAsync();
                string filter = TableQuery.GenerateFilterCondition(nameof(ScrumStatus.SummaryCardId), QueryComparisons.Equal, summaryCardId.Trim());
                var query = new TableQuery<ScrumStatus>().Where(filter);
                return await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in getting scrum status by summary card id: ScrumSummaryId: {summaryCardId}.", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Delete scrum status entity from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumStatus">Holds scrum status entity data.</param>
        /// <returns>Delete operation response.</returns>
        public async Task<TableResult> DeleteEntityAsync(ScrumStatus scrumStatus)
        {
            await this.EnsureInitializedAsync();
            TableOperation deleteOperation = TableOperation.Delete(scrumStatus);
            return await this.CloudTable.ExecuteAsync(deleteOperation);
        }
    }
}

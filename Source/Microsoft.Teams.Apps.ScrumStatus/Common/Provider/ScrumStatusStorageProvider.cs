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
                scrumStatusData = scrumStatusData ?? throw new ArgumentNullException(nameof(scrumStatusData));

                ScrumStatus scrumStatusDataEntity = new ScrumStatus()
                {
                    RowKey = $"{scrumStatusData?.ScrumStartCardResponseId.Trim()}_{scrumStatusData.Username}",
                    TodayTaskDescription = scrumStatusData.TodayTaskDescription,
                    YesterdayTaskDescription = scrumStatusData.YesterdayTaskDescription,
                    BlockerDescription = scrumStatusData.BlockerDescription,
                    ScrumStartCardResponseId = scrumStatusData.ScrumStartCardResponseId,
                    Username = scrumStatusData.Username,
                    AadGroupId = scrumStatusData.AadGroupId,
                    UserAadObjectId = scrumStatusData.UserAadObjectId,
                    MembersActivityIdMap = scrumStatusData.MembersActivityIdMap,
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
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>Returns collection of scrum status details.</returns>
        public async Task<IEnumerable<ScrumStatus>> GetScrumStatusBySummaryCardIdAsync(string summaryCardId, string aadGroupId)
        {
            summaryCardId = summaryCardId ?? throw new ArgumentNullException(nameof(summaryCardId));

            try
            {
                await this.EnsureInitializedAsync();
                string scrumStartCardResponseIdFilter = TableQuery.GenerateFilterCondition(nameof(ScrumStatus.ScrumStartCardResponseId), QueryComparisons.Equal, summaryCardId);
                string aadGroupIdfilter = TableQuery.GenerateFilterCondition(nameof(ScrumConfiguration.PartitionKey), QueryComparisons.Equal, aadGroupId);
                var combinedFilter = TableQuery.CombineFilters(scrumStartCardResponseIdFilter, TableOperators.And, aadGroupIdfilter);

                TableQuery<ScrumStatus> query = new TableQuery<ScrumStatus>().Where(combinedFilter);
                TableContinuationToken continuationToken = null;
                var scrumStatusCollection = new List<ScrumStatus>();

                do
                {
                    var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                    if (queryResult?.Results != null)
                    {
                        scrumStatusCollection.AddRange(queryResult.Results);
                        continuationToken = queryResult.ContinuationToken;
                    }
                }
                while (continuationToken != null);

                return scrumStatusCollection;
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

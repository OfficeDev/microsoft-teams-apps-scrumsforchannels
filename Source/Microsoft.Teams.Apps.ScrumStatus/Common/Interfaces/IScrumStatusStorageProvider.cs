// <copyright file="IScrumStatusStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.ScrumStatus.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Interface for provider class which helps in storing, updating, deleting scrum status details in Microsoft Azure Table storage.
    /// </summary>
    public interface IScrumStatusStorageProvider
    {
        /// <summary>
        /// Stores or update scrum status data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumStatusData">Holds scrum status entity data.</param>
        /// <returns>A task that represents scrum status entity data is saved or updated.</returns>
        Task<bool> CreateOrUpdateScrumStatusAsync(ScrumStatus scrumStatusData);

        /// <summary>
        /// Get scrum status by summary card id from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="summaryCardId">Scrum summary response card Id.</param>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>Returns collection of scrum status details.</returns>
        Task<IEnumerable<ScrumStatus>> GetScrumStatusBySummaryCardIdAsync(string summaryCardId, string aadGroupId);

        /// <summary>
        /// Delete scrum status entity from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumStatus">Holds scrum status entity data.</param>
        /// <returns>Delete operation response.</returns>
        Task<TableResult> DeleteEntityAsync(ScrumStatus scrumStatus);
    }
}

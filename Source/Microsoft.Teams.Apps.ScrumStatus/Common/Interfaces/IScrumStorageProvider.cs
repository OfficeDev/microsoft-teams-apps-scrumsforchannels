// <copyright file="IScrumStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.ScrumStatus.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Interface for provider class which helps in storing, updating, deleting scrum details in Microsoft Azure Table storage.
    /// </summary>
    public interface IScrumStorageProvider
    {
        /// <summary>
        /// Stores or update scrum data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumData">Holds scrum entity data.</param>
        /// <returns>A task that represents scrum entity data is saved or updated.</returns>
        Task<bool> CreateOrUpdateScrumAsync(Scrum scrumData);

        /// <summary>
        /// Get scrum details by summary card activity id from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="summaryCardActivityId">Summary card activity id.</param>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>Returns collection of scrum details by summary card activity id.</returns>
        Task<IEnumerable<Scrum>> GetScrumsBySummaryCardActivityIdAsync(string summaryCardActivityId, string aadGroupId);

        /// <summary>
        /// Get scrum by scrum team configuration id from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumTeamConfigId">Unique identifier for scrum team configuration details.</param>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>A <see cref="Task{TResult}"/>Representing the result of the asynchronous operation.</returns>
        Task<IEnumerable<Scrum>> GetScrumsByScrumTeamConfigIdAsync(string scrumTeamConfigId, string aadGroupId);

        /// <summary>
        /// Get scrum data from Microsoft Azure Table storage based on scrum start activity id.
        /// </summary>
        /// <param name="scrumStartActivityId">Scrum start activity id.</param>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>Returns collection of scrum details by summary card activity id.</returns>
        Task<IEnumerable<Scrum>> GetScrumsByScrumStartActivityIdAsync(string scrumStartActivityId, string aadGroupId);

        /// <summary>
        /// Get Scrum details by time stamp.
        /// </summary>
        /// <returns>task</returns>
        Task<IEnumerable<Scrum>> GetScrumDetailsByTimestampAsync();

        /// <summary>
        /// Delete scrum status entity from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrum">Holds scrum entity data.</param>
        /// <returns>Delete operation response.</returns>
        Task<TableResult> DeleteEntityAsync(Scrum scrum);
    }
}

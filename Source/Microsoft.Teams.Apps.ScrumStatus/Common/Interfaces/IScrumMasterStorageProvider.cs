// <copyright file="IScrumMasterStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.ScrumStatus.Models;

    /// <summary>
    ///  Scrum master provider helps in fetching and storing scrum master information in storage table.
    /// </summary>
    public interface IScrumMasterStorageProvider
    {
        /// <summary>
        /// Creates or updates Microsoft Azure Table storage to store scrum master details.
        /// </summary>
        /// <param name="scrumMasterEntities">Entities to be created or updated.</param>
        /// <returns>Boolean result.</returns>
        Task<bool> StoreOrUpdateScrumMasterEntitiesAsync(IEnumerable<ScrumMaster> scrumMasterEntities);

        /// <summary>
        /// Get scrum master details by scrum master id from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumMasterId">Unique identifier for scrum master data.</param>
        /// <returns>A <see cref="Task{TResult}"/> Represents the result of the asynchronous operation.</returns>
        Task<ScrumMaster> GetScrumMasterDetailsByScrumMasterIdAsync(string scrumMasterId);

        /// <summary>
        /// Get scrum master details by Azure Active Directory group Id.
        /// </summary>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>Returns collection of scrum master details.</returns>
        Task<IEnumerable<ScrumMaster>> GetScrumMasterDetailsbyAADGroupIDAsync(string aadGroupId);

        /// <summary>
        /// Get scrum master details which need to be scheduled in next 1 hour by start scrum background service.
        /// </summary>
        /// <returns>Returns collection of scrum master details.</returns>
        Task<IEnumerable<ScrumMaster>> GetActiveScrumMasterOfNextHourAsync();

        /// <summary>
        /// Delete an entity from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumMasterEntities">Entities to be deleted.</param>
        /// <returns>Boolean result.</returns>
        Task<bool> DeleteScrumMasterDetailsAsync(IEnumerable<ScrumMaster> scrumMasterEntities);
    }
}

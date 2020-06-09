// <copyright file="IScrumConfigurationStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.ScrumStatus.Models;

    /// <summary>
    ///  Scrum configuration provider helps in fetching and storing scrum configuration information in storage table.
    /// </summary>
    public interface IScrumConfigurationStorageProvider
    {
        /// <summary>
        /// Creates or updates Microsoft Azure Table storage to store scrum configuration details.
        /// </summary>
        /// <param name="scrumConfigurationEntities">Entities to be created or updated.</param>
        /// <returns>Boolean result.</returns>
        Task<bool> StoreOrUpdateScrumConfigurationEntitiesAsync(IEnumerable<ScrumConfiguration> scrumConfigurationEntities);

        /// <summary>
        /// Get scrum configuration details by scrum team configuration id from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumTeamConfigId">Unique identifier for scrum configuration data.</param>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>A <see cref="Task{TResult}"/> Represents the result of the asynchronous operation.</returns>
        Task<ScrumConfiguration> GetScrumConfigurationDetailByScrumTeamConfigIdAsync(string scrumTeamConfigId, string aadGroupId);

        /// <summary>
        /// Get scrum configuration details by Azure Active Directory group Id.
        /// </summary>
        /// <param name="aadGroupId">Azure Active Directory group Id.</param>
        /// <returns>Returns collection of scrum configuration details.</returns>
        Task<IEnumerable<ScrumConfiguration>> GetScrumConfigurationDetailsbyAADGroupIDAsync(string aadGroupId);

        /// <summary>
        /// Get scrum configuration details which need to be scheduled by current and previous UTC hour by start scrum background service.
        /// </summary>
        /// <returns>Returns collection of scrum configuration details.</returns>
        Task<IEnumerable<ScrumConfiguration>> GetActiveScrumConfigurationsByUtcHourAsync();

        /// <summary>
        /// Delete an entity from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumConfigurationEntities">Entities to be deleted.</param>
        /// <returns>Boolean result.</returns>
        Task<bool> DeleteScrumConfigurationDetailsAsync(IEnumerable<ScrumConfiguration> scrumConfigurationEntities);
    }
}

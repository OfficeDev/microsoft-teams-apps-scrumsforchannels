// <copyright file="IGraphUtilityHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common
{
    using System.Data;
    using System.Threading.Tasks;
    using Microsoft.Graph;

    /// <summary>
    /// This interface will contain the necessary methods to interface with Microsoft Graph.
    /// </summary>
    public interface IGraphUtilityHelper
    {
        /// <summary>
        /// Get drive details.
        /// </summary>
        /// <param name="groupId">Group id of the team in which channel is to be created.</param>
        /// <returns>A task that returns list of all channels in a team.</returns>
        Task<Drive> GetDriveDetailsAsync(string groupId);

        /// <summary>
        /// Upload excel file to provided drive location path.
        /// </summary>
        /// <param name="dataTable">Data table to be uploaded as excel sheet.</param>
        /// <param name="filePath">File path to upload.</param>
        /// <param name="driveId">Id of the team's drive.</param>
        /// <returns>A task that represents a HTTP response message including the status code and data.</returns>
        Task<string> UploadFileAsync(DataTable dataTable, string filePath, string driveId);
    }
}

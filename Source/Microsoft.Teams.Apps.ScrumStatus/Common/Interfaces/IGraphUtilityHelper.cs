// <copyright file="IGraphUtilityHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common
{
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.ScrumStatus.Models.Graph;
    using Microsoft.Teams.Shifts.Integration.Common.Models.Graph;

    /// <summary>
    /// This interface will contain the necessary methods to interface with Microsoft Graph.
    /// </summary>
    public interface IGraphUtilityHelper
    {
        /// <summary>
        /// Gets the Graph API token for the Tenant.
        /// </summary>
        /// <param name="tenantId">The tenantId.</param>
        /// <param name="clientId">The Azure AD App ID for the Configuration Web App.</param>
        /// <param name="clientSecret">The Azure AD App Secret for the Configuration Web App.</param>
        /// <returns>The graph token.</returns>
        Task<GraphTokenResponse> ObtainGraphTokenAsync(
            string tenantId,
            string clientId,
            string clientSecret);

        /// <summary>
        /// Get drive details.
        /// </summary>
        /// <param name="token">Azure Active Directory (AAD) token to access graph API.</param>
        /// <param name="groupId">groupId of the team in which channel is to be created.</param>
        /// <returns>A task that returns list of all channels in a team.</returns>
        Task<DriveDetails> GetDriveDetailsAsync(string token, string groupId);

        /// <summary>
        /// Method to post data to API.
        /// </summary>
        /// <param name="token">Microsoft Graph API user access token.</param>
        /// <param name="stream">In memory stream of data.</param>
        /// <param name="filePath">File path to upload.</param>
        /// <param name="driveId">Id of the Team's one drive</param>
        /// <returns>A task that represents a HTTP response message including the status code and data.</returns>
        Task<string> PutAsync(string token, MemoryStream stream, string filePath, string driveId);
    }
}

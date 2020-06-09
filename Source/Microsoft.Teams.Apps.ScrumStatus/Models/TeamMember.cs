// <copyright file="TeamMember.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Team members
    /// </summary>
    public class TeamMember
    {
        /// <summary>
        /// Gets or sets content
        /// </summary>
        [JsonProperty("content")]
        public string Content { get; set; }

        /// <summary>
        /// Gets or sets header value
        /// </summary>
        [JsonProperty("header")]
        public string Header { get; set; }

        /// <summary>
        /// Gets or sets team member Azure AD user object identifier
        /// </summary>
        [JsonProperty("aadobjectid")]
        public string AzureAdObjectId { get; set; }
    }
}

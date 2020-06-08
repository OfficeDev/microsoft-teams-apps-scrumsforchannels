// <copyright file="TeamDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Models
{
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// Team detail to share team member and account information
    /// </summary>
    public class TeamDetail
    {
        /// <summary>
        /// Gets or sets team members
        /// </summary>
        [JsonProperty("teamMembers")]
        public IEnumerable<TeamMember> TeamMembers { get; set; }

        /// <summary>
        /// Gets or sets team account details
        /// </summary>
        [JsonProperty("channels")]
        public IEnumerable<TeamAccount> Channels { get; set; }
    }
}

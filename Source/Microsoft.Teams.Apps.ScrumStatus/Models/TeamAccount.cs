// <copyright file="TeamAccount.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Team account information
    /// </summary>
    public class TeamAccount
    {
        /// <summary>
        /// Gets or sets team's channel id
        /// </summary>
        [JsonProperty("channelId")]
        public string ChannelId { get; set; }

        /// <summary>
        /// Gets or sets header value
        /// </summary>
        [JsonProperty("header")]
        public string Header { get; set; }
    }
}

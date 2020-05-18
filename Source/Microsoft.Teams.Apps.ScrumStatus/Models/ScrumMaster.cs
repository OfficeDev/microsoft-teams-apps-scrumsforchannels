// <copyright file="ScrumMaster.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Models
{
    // using System;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// Class containing scrum master details such as member, team and channel details.
    /// </summary>
    public class ScrumMaster : TableEntity
    {
        /// <summary>
        /// Gets or sets Unique Id of the row.
        /// </summary>
        [JsonProperty("Key")]
        public string Key { get; set; }

        /// <summary>
        /// Gets or sets scrum master Id to uniquely identify each scrum configuration.
        /// </summary>
        [JsonProperty("ScrumMasterId")]
        public string ScrumMasterId { get; set; }

        /// <summary>
        /// Gets or sets channel id in which scrum is configured.
        /// </summary>
        [JsonProperty("ChannelId")]
        public string ChannelId { get; set; }

        /// <summary>
        /// Gets or sets name of the channel in which scrum is configured.
        /// </summary>
        [JsonProperty("ChannelName")]
        public string ChannelName { get; set; }

        /// <summary>
        /// Gets or sets team Id in which bot is installed.
        /// </summary>
        [JsonProperty("TeamId")]
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets start time of the scrum at which scrum will get started daily.
        /// </summary>
        [JsonProperty("StartTime")]
        public string StartTime { get; set; }

        /// <summary>
        /// Gets or sets UTC hour of user specified start time.
        /// </summary>
        [JsonProperty("StartTimeUTCHour")]
        public int StartTimeUTCHour { get; set; }

        /// <summary>
        /// Gets or sets user specified time zone in which scrum is configured.
        /// </summary>
        [JsonProperty("TimeZone")]
        public string TimeZone { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether scrum is active.
        /// </summary>
        [JsonProperty("IsActive")]
        public bool IsActive { get; set; }

        /// <summary>
        /// Gets or sets user principle names of members of scrum.
        /// </summary>
        [JsonProperty("UserPrincipalNames")]
        public string UserPrincipalNames { get; set; }

        /// <summary>
        /// Gets or sets date on which scrum is created.
        /// </summary>
        [JsonProperty("CreatedOn")]
        public string CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets name of the person who created scrum.
        /// </summary>
        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets scrum team name.
        /// </summary>
        [JsonProperty("TeamName")]
        public string TeamName { get; set; }

        /// <summary>
        /// Gets or sets Azure Active Directory group Id in which bot is installed.
        /// </summary>
        [JsonProperty("AadGroupId")]
        public string AADGroupID { get; set; }

        /// <summary>
        /// Gets or sets activity service URL.
        /// </summary>
        [JsonProperty("ServiceUrl")]
        public string ServiceUrl { get; set; }
    }
}

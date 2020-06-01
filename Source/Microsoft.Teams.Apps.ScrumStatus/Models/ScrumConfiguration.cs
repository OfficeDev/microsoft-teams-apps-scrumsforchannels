// <copyright file="ScrumConfiguration.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Models
{
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// Class containing scrum configuration details such as member, team and channel details.
    /// </summary>
    public class ScrumConfiguration : TableEntity
    {
        /// <summary>
        /// Gets or sets Unique Id of the row.
        /// </summary>
        [JsonProperty("ScrumConfigurationId")]
        public string ScrumConfigurationId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets Azure Active Directory group id (Teams) where bot is installed.
        /// </summary>
        [JsonProperty("AadGroupId")]
        public string AadGroupId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets scrum team configuration id to identify each scrum team configuration details like team name, channel id in which scrum is configured.
        /// </summary>
        [JsonProperty("ScrumTeamConfigId")]
        public string ScrumTeamConfigId { get; set; }

        /// <summary>
        /// Gets or sets channel id in which scrum is configured, in case channel is deleted, setting it's value to "General channel id".
        /// </summary>
        [JsonProperty("ChannelId")]
        public string ChannelId { get; set; }

        /// <summary>
        /// Gets or sets name of the channel in which scrum is configured, in case channel is deleted, setting it's value to "General".
        /// </summary>
        [JsonProperty("ChannelName")]
        public string ChannelName { get; set; }

        /// <summary>
        /// Gets or sets team Id in which bot is installed.
        /// </summary>
        [JsonProperty("TeamId")]
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets start time of the scrum at which scrum will get started daily in round-trip format.
        /// </summary>
        [JsonProperty("StartTime")]
        public string StartTime { get; set; }

        /// <summary>
        /// Gets or sets UTC hour of user specified start time in "HH" format.
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
        /// Gets or sets user principal names of members of scrum, separated by comma(,).
        /// </summary>
        [JsonProperty("UserPrincipalNames")]
        public string UserPrincipalNames { get; set; }

        /// <summary>
        /// Gets or sets date on which scrum is created in Rfc3339DateTimeFormat.
        /// </summary>
        [JsonProperty("CreatedOn")]
        public string CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets name of the person who created scrum.
        /// </summary>
        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets scrum team name with which scrum will be running daily.
        /// </summary>
        [JsonProperty("ScrumTeamName")]
        public string ScrumTeamName { get; set; }

        /// <summary>
        /// Gets or sets activity service URL.
        /// </summary>
        [JsonProperty("ServiceUrl")]
        public string ServiceUrl { get; set; }
    }
}

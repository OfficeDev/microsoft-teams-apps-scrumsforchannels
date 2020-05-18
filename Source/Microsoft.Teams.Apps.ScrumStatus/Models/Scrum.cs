// <copyright file="Scrum.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Models
{
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// Class containing daily scrum details such as scrum card activity id, conversation id etc.
    /// </summary>
    public class Scrum : TableEntity
    {
        /// <summary>
        /// Gets or sets the conversation ID of the team chat that started the scrum.
        /// </summary>
        [JsonProperty("ThreadConversationId")]
        public string ThreadConversationId { get; set; }

        /// <summary>
        /// Gets or sets the unique id of the start scrum card.
        /// </summary>
        [JsonProperty("ScrumStartActivityId")]
        public string ScrumStartActivityId { get; set; }

        /// <summary>
        /// Gets or sets the id of the root scrum card.
        /// </summary>
        [JsonProperty("SummaryCardActivityId")]
        public string ScrumStartCardResponseId { get; set; }

        /// <summary>
        /// Gets or sets the scrum members activity id  which are mapped to start scrum card.
        /// </summary>
        [JsonProperty("MembersActivityIdMap")]
        public string MembersActivityIdMap { get; set; }

        /// <summary>
        /// Gets or sets the unique identifier of scrum master details.
        /// </summary>
        [JsonProperty("ScrumMasterId")]
        public string ScrumMasterId { get; set; }

        /// <summary>
        /// Gets or sets unique id of the scrum.
        /// </summary>
        [JsonProperty("ScrumId")]
        public string ScrumId { get; set; }

        /// <summary>
        /// Gets or sets TeamId in which bot is installed.
        /// </summary>
        [JsonProperty("TeamId")]
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets name of the channel in which scrum is configured.
        /// </summary>
        [JsonProperty("ChannelName")]
        public string ChannelName { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether scrum is completed or active.
        /// </summary>
        [JsonProperty("IsCompleted")]
        public bool IsCompleted { get; set; }

        /// <summary>
        /// Gets or sets date on which scrum is created.
        /// </summary>
        [JsonProperty("CreatedOn")]
        public string CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets Azure Active Directory group Id in which bot is installed.
        /// </summary>
        [JsonProperty("AadGroupId")]
        public string AADGroupID { get; set; }
    }
}

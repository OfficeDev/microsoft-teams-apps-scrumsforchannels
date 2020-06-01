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
        /// Gets or sets the conversation id of the Teams channel that started the scrum.
        /// </summary>
        [JsonProperty("ThreadConversationId")]
        public string ThreadConversationId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets the unique id(new guid) of the start scrum card.
        /// </summary>
        [JsonProperty("ScrumStartActivityId")]
        public string ScrumStartActivityId { get; set; }

        /// <summary>
        /// Gets or sets adaptive card response id to update the card after submitting scrum status.
        /// </summary>
        [JsonProperty("SummaryCardActivityId")]
        public string ScrumStartCardResponseId { get; set; }

        /// <summary>
        /// Gets or sets member activity id(guid) against each user principal name which are mapped to start scrum card.
        /// Dictionary is serialized into string and will be stored in this field.
        /// </summary>
        [JsonProperty("MembersActivityIdMap")]
        public string MembersActivityIdMap { get; set; }

        /// <summary>
        /// Gets or sets the unique identifier of scrum configuration details.
        /// </summary>
        [JsonProperty("ScrumTeamConfigId")]
        public string ScrumTeamConfigId { get; set; }

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
        /// Gets or sets name of the channel in which scrum is configured, in case channel is deleted, it's default value will be updated to "General".
        /// </summary>
        [JsonProperty("ChannelName")]
        public string ChannelName { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether scrum is completed or active.
        /// </summary>
        [JsonProperty("IsCompleted")]
        public bool IsCompleted { get; set; }

        /// <summary>
        /// Gets or sets date on which scrum is created in Rfc3339DateTimeFormat.
        /// </summary>
        [JsonProperty("CreatedOn")]
        public string CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets Azure Active Directory group Id in which bot is installed.
        /// </summary>
        [JsonProperty("AadGroupId")]
        public string AadGroupId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }
    }
}

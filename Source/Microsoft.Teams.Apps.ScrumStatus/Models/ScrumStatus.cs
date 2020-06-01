// <copyright file="ScrumStatus.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Models
{
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// class containing scrum status details which are updated by scrum members.
    /// </summary>
    public class ScrumStatus : TableEntity
    {
        /// <summary>
        /// Gets or sets yesterday task description updated by user.
        /// </summary>
        [JsonProperty("YesterdayTaskDescription")]
        public string YesterdayTaskDescription { get; set; }

        /// <summary>
        /// Gets or sets today task description updated by user.
        /// </summary>
        [JsonProperty("TodayTaskDescription")]
        public string TodayTaskDescription { get; set; }

        /// <summary>
        /// Gets or sets blockers updated by user.
        /// </summary>
        [JsonProperty("BlockerDescription")]
        public string BlockerDescription { get; set; }

        /// <summary>
        /// Gets or sets date on which scrum status is added by user in Rfc3339DateTimeFormat.
        /// </summary>
        [JsonProperty("CreatedOn")]
        public string CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets adaptive card response id to update the card after submitting scrum status.
        /// </summary>
        [JsonProperty("ScrumStartCardResponseId")]
        public string ScrumStartCardResponseId { get; set; }

        /// <summary>
        /// Gets or sets member activity id(new guid) against each user principal name to identify the valid scrum members.
        /// Dictionary is serialized into string and will be stored in this field.
        /// </summary>
        [JsonProperty("MembersActivityIdMap")]
        public string MembersActivityIdMap { get; set; }

        /// <summary>
        /// Gets or sets Azure Active Directory object id of user.
        /// </summary>
        [JsonProperty("UserAadObjectId")]
        public string UserAadObjectId { get; set; }

        /// <summary>
        /// Gets or sets user name.
        /// </summary>
        [JsonProperty("Username")]
        public string Username { get; set; }

        /// <summary>
        /// Gets or sets Azure Active Directory group id.
        /// </summary>
        [JsonProperty("AadGroupId")]
        public string AadGroupId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }
    }
}

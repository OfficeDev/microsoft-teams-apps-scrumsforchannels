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
        /// Gets or sets date at which card is updated.
        /// </summary>
        [JsonProperty("CreatedOn")]
        public string CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets summary card id.
        /// </summary>
        [JsonProperty("SummaryCardId")]
        public string SummaryCardId { get; set; }

        /// <summary>
        /// Gets or sets member activity id to map to actions.
        /// </summary>
        [JsonProperty("MembersActivityIdMap")]
        public string MembersActivityIdMap { get; set; }

        /// <summary>
        /// Gets or sets user name.
        /// </summary>
        [JsonProperty("Username")]
        public string Username { get; set; }

        /// <summary>
        /// Gets or sets Azure Active Directory object id of user.
        /// </summary>
        [JsonProperty("AadObjectId")]
        public string AadObjectId { get; set; }
    }
}

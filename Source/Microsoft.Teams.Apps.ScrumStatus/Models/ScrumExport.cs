// <copyright file="ScrumExport.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Model for exporting scrum data.
    /// </summary>
    public class ScrumExport
    {
        /// <summary>
        /// Gets or sets created on.
        /// </summary>
        [JsonProperty("CreatedOn")]
        public string DateOfScrum { get; set; }

        /// <summary>
        /// Gets or sets member name.
        /// </summary>
        [JsonProperty("Username")]
        public string MemberName { get; set; }

        /// <summary>
        /// Gets or sets yesterday field description.
        /// </summary>
        [JsonProperty("YesterdayTaskDescription")]
        public string WorkedUponYesterday { get; set; }

        /// <summary>
        /// Gets or sets blocker field description.
        /// </summary>
        [JsonProperty("BlockerDescription")]
        public string Blockers { get; set; }

        /// <summary>
        /// Gets or sets today field description.
        /// </summary>
        [JsonProperty("TodayTaskDescription")]
        public string PlanForToday { get; set; }
    }
}

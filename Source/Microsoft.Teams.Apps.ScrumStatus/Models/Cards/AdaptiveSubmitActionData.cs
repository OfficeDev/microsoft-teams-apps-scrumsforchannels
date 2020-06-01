// <copyright file="AdaptiveSubmitActionData.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Models
{
    using Microsoft.Bot.Schema;
    using Newtonsoft.Json;

    /// <summary>
    /// Adaptive submit card action data to post scrum and adaptive card related data.
    /// </summary>
    public class AdaptiveSubmitActionData
    {
        /// <summary>
        /// Gets or sets the Teams-specific action.
        /// </summary>
        [JsonProperty("msteams")]
        public CardAction MsTeams { get; set; }

        /// <summary>
        /// Gets or sets scrum members activity id.
        /// </summary>
        [JsonProperty("ScrumMembers")]
        public string ScrumMembers { get; set; }

        /// <summary>
        /// Gets or sets scrum summary id.
        /// </summary>
        [JsonProperty("summaryCardId")]
        public string SummaryCardId { get; set; }

        /// <summary>
        /// Gets or sets adaptive action type.
        /// </summary>
        [JsonProperty("AdaptiveActionType")]
        public string AdaptiveActionType { get; set; }

        /// <summary>
        /// Gets or sets scrum team configuration id obtained from storage while initiating the scrum.
        /// </summary>
        [JsonProperty("ScrumTeamConfigId")]
        public string ScrumTeamConfigId { get; set; }

        /// <summary>
        /// Gets or sets scrum start card activity id.
        /// </summary>
        [JsonProperty("ScrumStartActivityId")]
        public string ScrumStartActivityId { get; set; }
    }
}
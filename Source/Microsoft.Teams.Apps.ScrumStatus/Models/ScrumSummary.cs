// <copyright file="ScrumSummary.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Models
{
    /// <summary>
    /// Enumeration of scrum status.
    /// </summary>
    public enum ScrumRunningStatus
    {
        /// <summary>
        /// active flag.
        /// </summary>
        Active = 0,

        /// <summary>
        /// closed flag.
        /// </summary>
        Closed = 1,
    }

    /// <summary>
    /// Class containing scrum summary details.
    /// </summary>
    public class ScrumSummary
    {
        /// <summary>
        /// Gets or sets scrum name.
        /// </summary>
        public string ScrumName { get; set; }

        /// <summary>
        /// Gets or sets scrum start time.
        /// </summary>
        public string ScrumStartTime { get; set; }

        /// <summary>
        /// Gets or sets responded user count in the scrum.
        /// </summary>
        public int RespondedUserCount { get; set; }

        /// <summary>
        /// Gets or sets total user count present in the scrum.
        /// </summary>
        public int TotalUserCount { get; set; }

        /// <summary>
        /// Gets or sets scrum running status.
        /// </summary>
        public ScrumRunningStatus ScrumRunningStatus { get; set; }

        /// <summary>
        /// Gets or sets count of members who are blocked.
        /// </summary>
        public int BlockedUsersCount { get; set; }
    }
}

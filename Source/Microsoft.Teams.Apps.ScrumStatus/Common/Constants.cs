// <copyright file="Constants.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common
{
    /// <summary>
    /// Constant values that are used in multiple files.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Command to show help card.
        /// </summary>
        public const string Help = "HELP";

        /// <summary>
        /// Command to show settings card.
        /// </summary>
        public const string Settings = "OPEN SETTINGS";

        /// <summary>
        /// Command to end scrum.
        /// </summary>
        public const string EndScrum = "END SCRUM";

        /// <summary>
        /// Scrum details task module command when invoked.
        /// </summary>
        public const string ScrumDetailsTaskModuleCommand = "SCRUM DETAILS";

        /// <summary>
        /// Update scrum task module command when invoked.
        /// </summary>
        public const string UpdateStatusTaskModuleCommand = "UPDATE STATUS";

        /// <summary>
        /// Settings task module command when invoked.
        /// </summary>
        public const string SettingsTaskModuleCommand = "OPEN SETTINGS";

        /// <summary>
        /// Represents task module task/fetch string.
        /// </summary>
        public const string TaskModuleFetchType = "task/fetch";

        /// <summary>
        /// Represents task module task/submit string.
        /// </summary>
        public const string TaskModuleSubmitType = "task/submit";

        /// <summary>
        /// Described adaptive card version to be used. Version can be upgraded or changed using this value.
        /// </summary>
        public const string AdaptiveCardVersion = "1.2";

        /// <summary>
        /// Date time format to support adaptive card text feature.
        /// </summary>
        /// <remarks>
        /// refer adaptive card text feature https://docs.microsoft.com/en-us/adaptive-cards/authoring-cards/text-features#datetime-formatting-and-localization.
        /// </remarks>
        public const string Rfc3339DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'";

        /// <summary>
        /// Represents channel conversation type.
        /// </summary>
        public const string ChannelConversationType = "channel";
    }
}

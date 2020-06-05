// <copyright file="ExportOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common.Models
{
    /// <summary>
    /// Provides flag that whether scrum details need to be exported and deleted.
    /// </summary>
    public class ExportOptions
    {
        /// <summary>
        /// Gets or sets a value indicating whether gets or sets whether you want to export the scrum details.
        /// Set to true if export is enable.
        /// </summary>
        public bool IsExportEnabled { get; set; }
    }
}

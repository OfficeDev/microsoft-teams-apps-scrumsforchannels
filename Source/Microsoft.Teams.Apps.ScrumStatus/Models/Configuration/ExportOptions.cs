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
        /// Gets or sets whether you want to export the scrum details. Value is either true  or false.
        /// </summary>
        public string IsExportEnabled { get; set; }
    }
}

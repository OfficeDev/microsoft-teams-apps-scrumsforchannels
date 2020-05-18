// <copyright file="DriveDetails.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Models.Graph
{
    using System;
    using Newtonsoft.Json;

    /// <summary>
    /// Model for response obtained from graph api for drive details.
    /// </summary>
    public class DriveDetails
    {
        /// <summary>
        /// Gets or sets data context.
        /// </summary>
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        /// <summary>
        /// Gets or sets unique drive Id.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }
    }
}

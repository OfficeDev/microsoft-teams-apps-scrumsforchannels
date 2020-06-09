// <copyright file="ErrorResponse.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Class is wrapper to API error details.
    /// </summary>
    public class ErrorResponse
    {
        /// <summary>
        /// Gets or sets error status code.
        /// </summary>
        [JsonProperty("code")]
        public string StatusCode { get; set; }

        /// <summary>
        /// Gets or sets error message.
        /// </summary>
        [JsonProperty("message")]
        public string ErrorMessage { get; set; }
    }
}

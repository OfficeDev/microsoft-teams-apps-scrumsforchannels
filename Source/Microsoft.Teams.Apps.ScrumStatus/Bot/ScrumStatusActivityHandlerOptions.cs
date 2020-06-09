// <copyright file="ScrumStatusActivityHandlerOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus
{
    /// <summary>
    /// The ScrumStatusActivityHandlerOptions are the options for the <see cref="ScrumStatusActivityHandler" /> bot.
    /// </summary>
    public sealed class ScrumStatusActivityHandlerOptions
    {
        /// <summary>
        /// Gets or sets unique id of Tenant.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets application base Uri.
        /// </summary>
        public string AppBaseUri { get; set; }
    }
}
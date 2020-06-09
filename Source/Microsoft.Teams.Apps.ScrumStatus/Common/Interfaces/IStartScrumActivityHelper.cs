// <copyright file="IStartScrumActivityHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.ScrumStatus.Models;

    /// <summary>
    /// Handles the start scrum flow.
    /// </summary>
    public interface IStartScrumActivityHelper
    {
        /// <summary>
        /// Method to send start scrum card in the channel.
        /// </summary>
        /// <param name="scrumConfiguration">Scrum configuration details obtained from storage.</param>
        /// <returns>A task that sends start scrum card.</returns>
        Task ScrumStartActivityAsync(ScrumConfiguration scrumConfiguration);
    }
}
// <copyright file="ErrorCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Cards
{
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.ScrumStatus.Common;

    /// <summary>
    /// Class having method to return generic error card attachment.
    /// </summary>
    public static class ErrorCard
    {
        /// <summary>
        /// Get the generic error card on bot failure.
        /// </summary>
        /// <param name="errorMessage">Error message to be displayed in task module</param>
        /// <returns>Generic error card attachment.</returns>
        public static Attachment GetErrorCardAttachment(string errorMessage)
        {
            AdaptiveCard genericErrorCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = errorMessage,
                        Wrap = true,
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = genericErrorCard,
            };
        }
    }
}
// <copyright file="HelpCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Cards
{
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.ScrumStatus.Common;
    using Microsoft.Teams.Apps.ScrumStatus.Models;

    /// <summary>
    /// Class having method to return help card attachment.
    /// </summary>
    public static class HelpCard
    {
        /// <summary>
        /// Get the help card.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Help card attachment.</returns>
        public static Attachment GetHelpCard(IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard helpCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("HelpCardTitleText"),
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("SettingsButtonText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new TaskModuleAction(
                                localizer.GetString("SettingsButtonText"),
                                new AdaptiveSubmitActionData
                                {
                                    AdaptiveActionType = Constants.Settings,
                                }),
                        },
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = helpCard,
            };
        }
    }
}
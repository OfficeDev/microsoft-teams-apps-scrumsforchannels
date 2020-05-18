// <copyright file="WelcomeCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.ScrumStatus.Common;
    using Microsoft.Teams.Apps.ScrumStatus.Models;

    /// <summary>
    /// Class having method to return welcome card attachment.
    /// </summary>
    public static class WelcomeCard
    {
        /// <summary>
        /// Application Logo column width.
        /// </summary>
        public const string AppLogoColumnWidth = "2";

        /// <summary>
        /// Welcome card header column width.
        /// </summary>
        public const string HeaderColumnWidth = "10";

        /// <summary>
        /// Get welcome card attachment to show on Microsoft Teams team scope when bot is installed in team.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL to get the logo of the application.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Welcome card attachment.</returns>
        public static Attachment GetWelcomeCardAttachmentForChannel(string applicationBasePath, IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard welcomeCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AppLogoColumnWidth,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = new Uri($"{applicationBasePath}/Artifacts/appLogo.png"),
                                        Size = AdaptiveImageSize.Large,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = HeaderColumnWidth,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Large,
                                        Wrap = true,
                                        Text = localizer.GetString("WelcomeCardTitle"),
                                        Weight = AdaptiveTextWeight.Bolder,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Default,
                                        Wrap = true,
                                        Text = localizer.GetString("WelcomeCardSubtitleText"),
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = $"{localizer.GetString("WelcomeCardStartScrumTextHeading")}: {localizer.GetString("WelcomeCardStartScrumTextDesc")}",
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = $"{localizer.GetString("WelcomeCardEndScrumTextHeading")}: {localizer.GetString("WelcomeCardEndScrumTextDesc")}",
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("WelcomeCardWantToStartScrumText"),
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
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = welcomeCard,
            };
            return adaptiveCardAttachment;
        }
    }
}
// <copyright file="ScrumCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.ScrumStatus.Common;
    using Microsoft.Teams.Apps.ScrumStatus.Models;
    using Newtonsoft.Json;

    /// <summary>
    /// Class having methods to return attachments related to scrum.
    /// </summary>
    public static class ScrumCard
    {
        /// <summary>
        /// Maximum length for yesterday, today and any blocker input field.
        /// </summary>
        public const int InputFieldMaximumLength = 300;

        /// <summary>
        /// Get scrum start card when user clicks on update status button, end scrum button or scrum is started by background service.
        /// </summary>
        /// <param name="scrumSummary">Instance containing scrum related details.</param>
        /// <param name="membersActivityIdMap">Members id who are part of the scrum.</param>
        /// <param name="scrumTeamConfigId">Unique identifier for scrum configuration details.</param>
        /// <param name="scrumStartActivityId">Scrum start card activity id. This will be used to refresh card if needed.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="timeZone">Used to convert scrum start time as per specified time zone.</param>
        /// <returns>Scrum status update card attachment.</returns>
        public static Attachment GetScrumStartCard(ScrumSummary scrumSummary, Dictionary<string, string> membersActivityIdMap, string scrumTeamConfigId, string scrumStartActivityId, IStringLocalizer<Strings> localizer, string timeZone)
        {
            string userSpecifiedDateTime = FormatDateStringToAdaptiveCardDateFormat(scrumSummary?.ScrumStartTime, timeZone);
            string scrumMembers = JsonConvert.SerializeObject(membersActivityIdMap);
            string isAnyUserResponded = scrumSummary.RespondedUserCount == 0 ? localizer.GetString("StartScrumCardWelcomeText") : string.Format(CultureInfo.CurrentCulture, localizer.GetString("StartScrumCardBlockedText"), scrumSummary.BlockedUsersCount);
            AdaptiveCard getScrumStartCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                        Text = scrumSummary.ScrumName,
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Wrap = true,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                        Wrap = true,
                                        Text = scrumSummary.ScrumRunningStatus == ScrumRunningStatus.Active ? localizer.GetString("ScrumRunningStatusActive") : localizer.GetString("ScrumRunningStatusClosed"),
                                        Color = scrumSummary.ScrumRunningStatus == ScrumRunningStatus.Active ? AdaptiveTextColor.Good : AdaptiveTextColor.Default,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                        Wrap = true,
                                        Text = string.Format(CultureInfo.CurrentCulture, localizer.GetString("ScrumDetailsRespondedVsTotalUserStatus"), scrumSummary.RespondedUserCount, scrumSummary.TotalUserCount),
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                        Text = userSpecifiedDateTime,
                                        Wrap = true,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                { // Represents empty column required for alignment.
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                { // Represents empty column required for alignment.
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = isAnyUserResponded,
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                        Color = scrumSummary.RespondedUserCount == 0 ? AdaptiveTextColor.Default : AdaptiveTextColor.Attention,
                                        Height = AdaptiveHeight.Stretch,
                                        Separator = true,
                                    },
                                },
                                Spacing = AdaptiveSpacing.Medium,
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                { // Represents empty column required for alignment.
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                { // Represents empty column required for alignment.
                                },
                            },
                        },
                    },
                },
            };

            var scrumDetailsAdaptiveSubmitAction = new AdaptiveSubmitAction()
            {
                Title = localizer.GetString("ScrumDetailsButtonText"),
                Data = new AdaptiveSubmitActionData
                {
                    MsTeams = new CardAction
                    {
                        Type = Constants.TaskModuleFetchType,
                    },
                    AdaptiveActionType = Constants.ScrumDetailsTaskModuleCommand,
                    ScrumTeamConfigId = scrumTeamConfigId,
                    ScrumStartActivityId = scrumStartActivityId,
                    ScrumMembers = scrumMembers,
                },
            };

            if (scrumSummary.ScrumRunningStatus == ScrumRunningStatus.Active)
            {
                getScrumStartCard.Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction()
                    {
                        Title = localizer.GetString("UpdateStatusButtonText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new CardAction
                            {
                                Type = Constants.TaskModuleFetchType,
                            },
                            AdaptiveActionType = Constants.UpdateStatusTaskModuleCommand,
                            ScrumTeamConfigId = scrumTeamConfigId,
                            ScrumStartActivityId = scrumStartActivityId,
                            ScrumMembers = scrumMembers,
                        },
                    },
                    scrumDetailsAdaptiveSubmitAction,
                    new AdaptiveSubmitAction()
                    {
                        Title = localizer.GetString("EndScrumButtonText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                            },
                            AdaptiveActionType = Constants.EndScrum,
                            ScrumTeamConfigId = scrumTeamConfigId,
                            ScrumStartActivityId = scrumStartActivityId,
                            ScrumMembers = scrumMembers,
                        },
                    },
                };
            }
            else
            {
                getScrumStartCard.Actions = new List<AdaptiveAction>
                {
                    scrumDetailsAdaptiveSubmitAction,
                };
            }

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = getScrumStartCard,
            };
            return adaptiveCardAttachment;
        }

        /// <summary>
        /// Get scrum details card for viewing all user status details.
        /// </summary>
        /// <param name="scrumStatuses">Scrum statuses details filled by members.</param>
        /// <param name="scrumSummary">Scrum related summary.</param>
        /// <param name="scrumMembers">Scrum members present in the scrum and part of the Team.</param>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="timeZone">Used to convert scrum start time as per specified time zone.</param>
        /// <returns>Returns scrum details card attachment.</returns>
        public static Attachment GetScrumDetailsCard(IEnumerable<ScrumStatus> scrumStatuses, ScrumSummary scrumSummary, IEnumerable<TeamsChannelAccount> scrumMembers, string applicationBasePath, IStringLocalizer<Strings> localizer, string timeZone)
        {
            string userSpecifiedDateTime = FormatDateStringToAdaptiveCardDateFormat(scrumSummary?.ScrumStartTime, timeZone);
            AdaptiveCard scrumDetailsCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                        Text = scrumSummary.ScrumName,
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Wrap = true,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Color = scrumSummary.ScrumRunningStatus == ScrumRunningStatus.Active ? AdaptiveTextColor.Good : AdaptiveTextColor.Default,
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                        Text = scrumSummary.ScrumRunningStatus == ScrumRunningStatus.Active ? localizer.GetString("ScrumRunningStatusActive") : localizer.GetString("ScrumRunningStatusClosed"),
                                        Wrap = true,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                        Text = string.Format(CultureInfo.CurrentCulture, localizer.GetString("ScrumDetailsRespondedVsTotalUserStatus"), scrumSummary.RespondedUserCount, scrumSummary.TotalUserCount),
                                        Wrap = true,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                        Text = userSpecifiedDateTime,
                                        Wrap = true,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                { // Represents empty column required for alignment.
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                { // Represents empty column required for alignment.
                                },
                            },
                        },
                    },
                },
            };

            List<AdaptiveElement> userScrumStatusList = GetAllUserScrumStatusCard(scrumStatuses, scrumMembers, applicationBasePath, localizer);
            scrumDetailsCard.Body.AddRange(userScrumStatusList);

            var scrumDetailsCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = scrumDetailsCard,
            };

            return scrumDetailsCardAttachment;
        }

        /// <summary>
        /// Get scrum status details card for all members for viewing status details.
        /// </summary>
        /// <param name="scrumStatuses">Scrum statuses filled by all members.</param>
        /// <param name="scrumMembers">Scrum members present in the scrum and part of the Team.</param>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns scrum details card attachment.</returns>
        public static List<AdaptiveElement> GetAllUserScrumStatusCard(IEnumerable<ScrumStatus> scrumStatuses, IEnumerable<TeamsChannelAccount> scrumMembers, string applicationBasePath, IStringLocalizer<Strings> localizer)
        {
            int id = 1;
            List<AdaptiveElement> allUserScrumStatusDetails = new List<AdaptiveElement>();

            if (scrumMembers == null)
            {
                return null;
            }

            foreach (TeamsChannelAccount member in scrumMembers)
            {
                var userInfo = scrumStatuses.Where(scrumStatus => scrumStatus.UserAadObjectId.Equals(member.AadObjectId, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
                if (userInfo == null)
                {
                    allUserScrumStatusDetails.AddRange(GetUserScrumStatusCard(new ScrumStatus(), member, applicationBasePath, localizer, id, localizer.GetString("ScrumStatusNotStarted")));
                }
                else
                {
                    bool isBlocked = string.IsNullOrEmpty(userInfo.BlockerDescription) ? false : true;
                    allUserScrumStatusDetails.AddRange(GetUserScrumStatusCard(userInfo, member, applicationBasePath, localizer, id, localizer.GetString("ScrumStatusCompletedForUser"), isBlocked));
                }

                id += 1;
            }

            return allUserScrumStatusDetails;
        }

        /// <summary>
        /// Get user scrum status details card for viewing all user status details.
        /// </summary>
        /// <param name="scrumStatus">Scrum status details filled by a member.</param>
        /// <param name="member">Member information whose scrum status is to be showed.</param>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="id">Unique id for each row.</param>
        /// <param name="userScrumStatus">Individual user scrum status i.e. Not started or completed.</param>
        /// <param name="isBlocked">Determine whether user is blocked.</param>
        /// <returns>Returns scrum details card attachment.</returns>
        public static List<AdaptiveElement> GetUserScrumStatusCard(ScrumStatus scrumStatus, TeamsChannelAccount member, string applicationBasePath, IStringLocalizer<Strings> localizer, int id, string userScrumStatus, bool isBlocked = false)
        {
            string cardContent = $"CardContent{id}";
            string chevronUp = $"ChevronUp{id}";
            string chevronDown = $"ChevronDown{id}";
            List<AdaptiveElement> userScrumStatusDetails = new List<AdaptiveElement>
            {
                new AdaptiveColumnSet
                {
                    Separator = true,
                    Spacing = AdaptiveSpacing.Medium,
                    Columns = new List<AdaptiveColumn>
                    {
                        new AdaptiveColumn
                        {
                            VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                            Width = AdaptiveColumnWidth.Stretch,
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveTextBlock
                                {
                                    Text = member?.Name,
                                    Wrap = true,
                                },
                            },
                        },
                        new AdaptiveColumn
                        {
                            VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                            Width = AdaptiveColumnWidth.Auto,
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveTextBlock
                                {
                                    HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                    Spacing = string.Equals(userScrumStatus, localizer.GetString("ScrumStatusNotStarted"), StringComparison.OrdinalIgnoreCase) ? AdaptiveSpacing.Small : AdaptiveSpacing.None,
                                    Text = userScrumStatus,
                                    Wrap = true,
                                },
                            },
                        },
                        new AdaptiveColumn
                        {
                            VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                            Width = AdaptiveColumnWidth.Stretch,
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveTextBlock
                                {
                                    HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                    Spacing = AdaptiveSpacing.None,
                                    Text = localizer.GetString("ScrumDetailsUserIsBlocked"),
                                    Wrap = true,
                                    Color = AdaptiveTextColor.Attention,
                                    IsVisible = isBlocked,
                                },
                            },
                        },
                        new AdaptiveColumn
                        {
                            Id = chevronDown,
                            VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                            Width = AdaptiveColumnWidth.Auto,
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveImage
                                {
                                    AltText = "collapse",
                                    PixelWidth = 16,
                                    PixelHeight = 8,
                                    SelectAction = new AdaptiveToggleVisibilityAction
                                    {
                                        Title = "collapse",
                                        Type = "Action.ToggleVisibility",
                                        TargetElements = new List<AdaptiveTargetElement>
                                        {
                                            cardContent,
                                            chevronUp,
                                            chevronDown,
                                        },
                                    },
                                    Style = AdaptiveImageStyle.Default,
                                    Url = new Uri(applicationBasePath + "/Artifacts/chevronDown.png"),
                                },
                            },
                        },
                        new AdaptiveColumn
                        {
                            Id = chevronUp,
                            IsVisible = false,
                            VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                            Width = AdaptiveColumnWidth.Auto,
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveImage
                                {
                                    AltText = "expand",
                                    PixelWidth = 16,
                                    PixelHeight = 8,
                                    SelectAction = new AdaptiveToggleVisibilityAction
                                    {
                                        Title = "expand",
                                        Type = "Action.ToggleVisibility",
                                        TargetElements = new List<AdaptiveTargetElement>
                                        {
                                            cardContent,
                                            chevronUp,
                                            chevronDown,
                                        },
                                    },
                                    Style = AdaptiveImageStyle.Default,
                                    Url = new Uri(applicationBasePath + "/Artifacts/chevronUp.png"),
                                },
                            },
                        },
                    },
                },
                new AdaptiveContainer
                {
                    Id = cardContent,
                    IsVisible = false,
                    Items = new List<AdaptiveElement>
                    {
                        new AdaptiveContainer
                        {
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveTextBlock
                                {
                                    IsSubtle = true,
                                    HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                    Text = localizer.GetString("UpdateStatusYesterdayDescriptionTitle"),
                                    Wrap = true,
                                },
                                new AdaptiveContainer
                                {
                                    Spacing = AdaptiveSpacing.Small,
                                    Style = AdaptiveContainerStyle.Emphasis,
                                    Items = new List<AdaptiveElement>
                                    {
                                        new AdaptiveTextBlock
                                        {
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                            Spacing = AdaptiveSpacing.Small,
                                            Text = string.IsNullOrEmpty(scrumStatus?.YesterdayTaskDescription) ? string.Empty : scrumStatus?.YesterdayTaskDescription,
                                            Wrap = true,
                                        },
                                    },
                                },
                                new AdaptiveTextBlock
                                {
                                    IsSubtle = true,
                                    HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                    Text = localizer.GetString("UpdateStatusTodayDescriptionTitle"),
                                    Wrap = true,
                                },
                                new AdaptiveContainer
                                {
                                    Spacing = AdaptiveSpacing.Small,
                                    Style = AdaptiveContainerStyle.Emphasis,
                                    Items = new List<AdaptiveElement>
                                    {
                                        new AdaptiveTextBlock
                                        {
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                            Spacing = AdaptiveSpacing.Small,
                                            Text = string.IsNullOrEmpty(scrumStatus?.TodayTaskDescription) ? string.Empty : scrumStatus?.TodayTaskDescription,
                                            Wrap = true,
                                        },
                                    },
                                },
                                new AdaptiveTextBlock
                                {
                                    Color = AdaptiveTextColor.Attention,
                                    IsSubtle = true,
                                    HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                    Text = localizer.GetString("UpdateStatusAnyBlockersTitle"),
                                    Weight = AdaptiveTextWeight.Bolder,
                                    Wrap = true,
                                },
                                new AdaptiveContainer
                                {
                                    Spacing = AdaptiveSpacing.Small,
                                    Style = AdaptiveContainerStyle.Emphasis,
                                    Items = new List<AdaptiveElement>
                                    {
                                        new AdaptiveTextBlock
                                        {
                                            Color = AdaptiveTextColor.Attention,
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                            Spacing = AdaptiveSpacing.Small,
                                            Text = string.IsNullOrEmpty(scrumStatus?.BlockerDescription) ? string.Empty : scrumStatus?.BlockerDescription,
                                            Wrap = true,
                                        },
                                    },
                                },
                            },
                        },
                    },
                },
            };

            return userScrumStatusDetails;
        }

        /// <summary>
        /// Get scrum status update card when user clicks on update status button.
        /// </summary>
        /// <param name="membersId">Members id to verify the user.</param>
        /// <param name="scrumTeamConfigId">Scrum team configuration id to show updated status.</param>
        /// <param name="scrumStartActivityId">Activity id of scrum start card.</param>
        /// <param name="scrumStatus">Scrum status information.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="isYesterdayValidationFailure">Determines whether to show yesterday validation error.</param>
        /// <param name="isTodayValidationFailure">Determines whether to show today validation error.</param>
        /// <returns>Scrum status update card attachment.</returns>
        public static Attachment GetScrumStatusUpdateCard(string membersId, string scrumTeamConfigId, string scrumStartActivityId, ScrumStatus scrumStatus, IStringLocalizer<Strings> localizer, bool isYesterdayValidationFailure = false, bool isTodayValidationFailure = false)
        {
            AdaptiveCard scrumStatusUpdateCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = localizer.GetString("UpdateStatusYesterdayDescriptionTitle"),
                                        Spacing = AdaptiveSpacing.None,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = localizer.GetString("UpdateStatusYesterdayDescriptionErrorMessage"),
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                        Color = AdaptiveTextColor.Attention,
                                        IsVisible = isYesterdayValidationFailure,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextInput
                    {
                        Spacing = AdaptiveSpacing.Small,
                        Placeholder = localizer.GetString("UpdateStatusYesterdayDescriptionPlaceholder"),
                        IsMultiline = true,
                        Style = AdaptiveTextInputStyle.Text,
                        Id = "yesterdaytaskdescription",
                        MaxLength = InputFieldMaximumLength,
                        Value = scrumStatus?.YesterdayTaskDescription,
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = localizer.GetString("UpdateStatusTodayDescriptionTitle"),
                                        Spacing = AdaptiveSpacing.None,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = localizer.GetString("UpdateStatusTodayDescriptionErrorMessage"),
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                        Color = AdaptiveTextColor.Attention,
                                        IsVisible = isTodayValidationFailure,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextInput
                    {
                        Spacing = AdaptiveSpacing.Small,
                        Placeholder = localizer.GetString("UpdateStatusTodayDescriptionPlaceholder"),
                        IsMultiline = true,
                        Style = AdaptiveTextInputStyle.Text,
                        Id = "todaytaskdescription",
                        MaxLength = InputFieldMaximumLength,
                        Value = scrumStatus.TodayTaskDescription,
                    },
                    new AdaptiveTextBlock
                    {
                        Size = AdaptiveTextSize.Medium,
                        Wrap = true,
                        Text = localizer.GetString("UpdateStatusAnyBlockersTitle"),
                        Color = AdaptiveTextColor.Attention,
                    },
                    new AdaptiveTextInput
                    {
                        Spacing = AdaptiveSpacing.Small,
                        Placeholder = localizer.GetString("UpdateStatusAnyBlockersPlaceholder"),
                        IsMultiline = true,
                        Style = AdaptiveTextInputStyle.Text,
                        Id = "blockerdescription",
                        MaxLength = InputFieldMaximumLength,
                        Value = scrumStatus.BlockerDescription,
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                { // Represents empty column required for alignment.
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                { // Represents empty column required for alignment.
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                { // Represents empty column required for alignment.
                                },
                            },
                        },
                    },
                },
            };

            scrumStatusUpdateCard.Actions.Add(
                    new AdaptiveSubmitAction()
                    {
                        Title = localizer.GetString("UpdateStatusSubmitButtonText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new CardAction
                            {
                                Type = Constants.TaskModuleSubmitType,
                            },
                            ScrumMembers = membersId,
                            ScrumTeamConfigId = scrumTeamConfigId,
                            ScrumStartActivityId = scrumStartActivityId,
                            AdaptiveActionType = Constants.UpdateStatusTaskModuleCommand,
                        },
                    });

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = scrumStatusUpdateCard,
            };
            return adaptiveCardAttachment;
        }

        /// <summary>
        /// Convert Date time format to adaptive card text feature as per user specified time zone.
        /// </summary>
        /// <param name="inputDateTimeText">Input date time string.</param>
        /// <param name="timeZone">Used to convert scrum start time as per specified time zone.</param>
        /// <returns>Adaptive card supported date time format.</returns>
        private static string FormatDateStringToAdaptiveCardDateFormat(string inputDateTimeText, string timeZone)
        {
            try
            {
                // Convert scrum start time with user specified time zone
                TimeZoneInfo timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(timeZone);
                DateTime scrumStartTimeAsSpecifiedDateTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.Parse(inputDateTimeText, CultureInfo.InvariantCulture), timeZoneInfo);
                return scrumStartTimeAsSpecifiedDateTime.ToString("ddd, MMM dd, yyy", CultureInfo.CurrentCulture);
            }
            #pragma warning disable CA1031 // Do not catch general exception types
            catch
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                return inputDateTimeText;
            }
        }
    }
}
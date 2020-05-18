// <copyright file="ResourceController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Controllers
{
    using System;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// Controller to handle resource strings related request.
    /// </summary>
    [Route("api/resource")]
    [Authorize]
    [ApiController]
    public class ResourceController : ControllerBase
    {
        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Initializes a new instance of the <see cref="ResourceController"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        public ResourceController(ILogger<ResourceController> logger, IStringLocalizer<Strings> localizer)
        {
            this.logger = logger;
            this.localizer = localizer;
        }

        /// <summary>
        /// Get resource strings for displaying in client application.
        /// </summary>
        /// <returns>Resource strings according to user locale.</returns>
        [Route("resourcestrings")]
        public IActionResult GetResourceStrings()
        {
            try
            {
                var strings = new
                {
                    TeamNameTitle = this.localizer.GetString("TeamNameTitle").Value,
                    MembersListTitle = this.localizer.GetString("MembersListTitle").Value,
                    StartEveryDayTitle = this.localizer.GetString("StartEveryDayTitle").Value,
                    TimeZoneTitle = this.localizer.GetString("TimeZoneTitle").Value,
                    AddToChannelTitle = this.localizer.GetString("AddToChannelTitle").Value,
                    SelectTimePlaceholder = this.localizer.GetString("SelectTimePlaceholder").Value,
                    SelectTimeZonePlaceholder = this.localizer.GetString("SelectTimeZonePlaceholder").Value,
                    SelectChannelPlaceholder = this.localizer.GetString("SelectChannelPlaceholder").Value,
                    AddNewScrumButtonText = this.localizer.GetString("AddNewScrumButtonText").Value,
                    SaveButtonText = this.localizer.GetString("SaveButtonText").Value,
                    TeamNameValidationText = this.localizer.GetString("TeamNameValidationText").Value,
                    TeamMembersValidationText = this.localizer.GetString("TeamMembersValidationText").Value,
                    StartTimeValidationText = this.localizer.GetString("StartTimeValidationText").Value,
                    TimeZoneValidationText = this.localizer.GetString("TimeZoneValidationText").Value,
                    ChannelNameValidationText = this.localizer.GetString("ChannelNameValidationText").Value,
                    DuplicateScrumValidationText = this.localizer.GetString("DuplicateScrumValidationText").Value,
                    ErrorMessage = this.localizer.GetString("ErrorMessage").Value,
                    UnauthorizedAccess = this.localizer.GetString("UnauthorizedAccess").Value,
                    ScrumEnableButtonTitle = this.localizer.GetString("ScrumEnableButtonTitle").Value,
                    SelectUserPlaceholder = this.localizer.GetString("SelectUserPlaceholder").Value,
                    DeleteButtonText = this.localizer.GetString("DeleteButtonText").Value,
                    CancelButtonText = this.localizer.GetString("CancelButtonText").Value,
                    NoMatchesFoundText = this.localizer.GetString("NoMatchesFoundText").Value,
                };
                return this.Ok(strings);
            }
            #pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                this.logger.LogError(ex, "Error while fetching resource strings.");
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }
    }
}
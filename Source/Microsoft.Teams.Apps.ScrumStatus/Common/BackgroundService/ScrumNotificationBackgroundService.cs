// <copyright file="ScrumNotificationBackgroundService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common.BackgroundService
{
    using System;
    using System.Globalization;
    using System.Threading;
    using System.Threading.Tasks;
    using Cronos;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// Background service to send proactive notifications on registered channels.
    /// </summary>
    public sealed class ScrumNotificationBackgroundService : BackgroundService
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<ScrumNotificationBackgroundService> logger;

        /// <summary>
        /// Storage helper for working with scrum configuration data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumConfigurationStorageProvider scrumConfigurationStorageProvider;

        /// <summary>
        /// Start scrum activity helper to send card in channel.
        /// </summary>
        private readonly IStartScrumActivityHelper startScrumActivityHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrumNotificationBackgroundService"/> class.
        /// </summary>
        /// <param name="scrumConfigurationStorageProvider">scrumConfigurationStorageProvider</param>
        /// <param name="startScrumActivityHelper">startScrumActivityHelper</param>
        /// <param name="logger">logger</param>
        public ScrumNotificationBackgroundService(
            IScrumConfigurationStorageProvider scrumConfigurationStorageProvider,
            IStartScrumActivityHelper startScrumActivityHelper,
            ILogger<ScrumNotificationBackgroundService> logger)
        {
            this.scrumConfigurationStorageProvider = scrumConfigurationStorageProvider;
            this.startScrumActivityHelper = startScrumActivityHelper;
            this.logger = logger;
        }

        /// <summary>
        ///  This method is called when the Microsoft.Extensions.Hosting.IHostedService starts.
        ///  The implementation should return a task that represents the lifetime of the long
        ///  running operation(s) being performed.
        /// </summary>
        /// <param name="stoppingToken">Triggered when Microsoft.Extensions.Hosting.IHostedService.StopAsync(System.Threading.CancellationToken) is called.</param>
        /// <returns>A System.Threading.Tasks.Task that represents the long running operations.</returns>
        protected async override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    this.logger.LogInformation("Scrum notification service has started...");
                    await this.ProcessScheduledScrumNotificationAsync();

                    // Schedule next run in 30 minutes
                    CronExpression storageCronExpression = CronExpression.Parse("*/30 * * * *");
                    var next = storageCronExpression.GetNextOccurrence(DateTimeOffset.Now, TimeZoneInfo.Local);
                    var delay = next.HasValue ? next.Value - DateTimeOffset.Now : TimeSpan.FromMinutes(30);
                    await Task.Delay(delay, stoppingToken);
                }
#pragma warning disable CA1031 // Catching general exceptions that might arise during execution to avoid blocking next run.
                catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions that might arise during execution to avoid blocking next run.
                {
                    this.logger.LogError($"Exception occurred while executing scrum notification.", ex);
                }
            }
        }

        private async Task ProcessScheduledScrumNotificationAsync()
        {
            // get the current UTC hour schedule
            // this is reading the data for current UTC hour from storage.
            var scrumConfigurationDetails = await this.scrumConfigurationStorageProvider
                .GetActiveScrumConfigurationsByUtcHourAsync();

            if (scrumConfigurationDetails != null)
            {
                // remove seconds from UTC date time to compare with scrum start time.
                var dateTimeWithoutSeconds = DateTime.UtcNow.ToString("MM/dd/yyyy HH:mm", CultureInfo.InvariantCulture);
                var schedulerTime = DateTime.Parse(dateTimeWithoutSeconds, CultureInfo.InvariantCulture);

                foreach (var scrumConfiguration in scrumConfigurationDetails)
                {
                    // get scrum configured start time and time zone
                    TimeZoneInfo userSpecifiedTimeZone = TimeZoneInfo.FindSystemTimeZoneById(scrumConfiguration.TimeZone);
                    DateTime scheduledStartTime = TimeZoneInfo.ConvertTimeToUtc(DateTime.Parse(scrumConfiguration.StartTime, CultureInfo.InvariantCulture), userSpecifiedTimeZone);

                    // if the start time is scheduled for next 30 minutes, then start processing else skip.
                    // if the start window is missed, scrum will be skipped for the day and will be scheduled on next day.
                    if (schedulerTime.TimeOfDay <= scheduledStartTime.TimeOfDay
                        && schedulerTime.AddMinutes(30).TimeOfDay > scheduledStartTime.TimeOfDay)
                    {
                        this.logger.LogInformation($"scrum for team {scrumConfiguration.ScrumTeamName} is ready to start {scheduledStartTime.TimeOfDay}");
                        await this.startScrumActivityHelper.ScrumStartActivityAsync(scrumConfiguration);
                    }
                }
            }
        }
    }
}
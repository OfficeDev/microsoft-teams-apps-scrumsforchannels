// <copyright file="StartScrumBackgroundService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Common
{
    using System;
    using System.Globalization;
    using System.Threading;
    using System.Threading.Tasks;
    using Cronos;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.ScrumStatus.Helpers;
    using Microsoft.Teams.Apps.ScrumStatus.Models;

    /// <summary>
    /// BackgroundService class that inherits IHostedService and implements the methods related to background tasks for sending notification two times a day.
    /// </summary>
    public class StartScrumBackgroundService : IHostedService, IDisposable
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<StartScrumBackgroundService> logger;

        /// <summary>
        /// Storage helper for working with scrum master data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumMasterStorageProvider scrumMasterStorageProvider;

        /// <summary>
        /// Start scrum activity helper to send card in channel.
        /// </summary>
        private readonly IStartScrumActivityHelper startScrumActivityHelper;

        /// <summary>
        /// Timer to schedule scrum.
        /// </summary>
        private System.Timers.Timer timer;

        /// <summary>
        /// Timer to schedule scrum master storage call.
        /// </summary>
        private System.Timers.Timer scrumTimer;

        /// <summary>
        /// Execution count.
        /// </summary>
        private int executionCount = 0;

        /// <summary>
        /// Flag: Has Dispose already been called?.
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="StartScrumBackgroundService"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to sending notification tasks.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="scrumMasterStorageProvider">Provider to provide scrum master storage details.</param>
        /// <param name="startScrumActivityHelper">An instance of scrum activity helper.</param>
        public StartScrumBackgroundService(
            ILogger<StartScrumBackgroundService> logger,
            IScrumMasterStorageProvider scrumMasterStorageProvider,
            IStartScrumActivityHelper startScrumActivityHelper)
        {
            this.logger = logger;
            this.scrumMasterStorageProvider = scrumMasterStorageProvider;
            this.startScrumActivityHelper = startScrumActivityHelper;
        }

        /// <summary>
        /// Method to start the background task when application starts.
        /// </summary>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task instance.</returns>
        public Task StartAsync(CancellationToken cancellationToken)
        {
            this.logger.LogInformation("Start scrum Hosted Service is running");
            this.ScheduleStorage(); // Schedule the storage call to fetch scrum details.
            return Task.CompletedTask;
        }

        /// <summary>
        /// Triggered when the host is performing a graceful shutdown.
        /// </summary>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task instance.</returns>
        public Task StopAsync(CancellationToken cancellationToken)
        {
            this.logger.LogInformation("Start scrum Hosted Service is stopping");
            return Task.CompletedTask;
        }

        /// <summary>
        /// This code added to correctly implement the disposable pattern.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        /// <param name="disposing">True if already disposed else false.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (this.disposed)
            {
                return;
            }

            if (disposing)
            {
                this.timer.Dispose();
                this.scrumTimer.Dispose();
            }

            this.disposed = true;
        }

        /// <summary>
        /// Set the timer and enqueue start scrum task if timer matched as per CRON expression.
        /// </summary>
        /// <returns>A task that schedules scrum of next hour.</returns>
        /// <remark>Scrum schedules for next hour. For example at 9 am it schedules all scrum of 10 and 10:30 am.</remark>
        private Task ScheduleScrumAsync(CronExpression expression, ScrumMaster scrumMaster)
        {
            var count = Interlocked.Increment(ref this.executionCount);
            this.logger.LogInformation($"Scheduling scrum task is working. Count: {count}");

            // Get the timezone entered by the user to schedule scrum on user specified time.
            TimeZoneInfo timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(scrumMaster.TimeZone);
            var next = expression.GetNextOccurrence(DateTimeOffset.Now, timeZoneInfo);
            if (next.HasValue)
            {
                var delay = next.Value - DateTimeOffset.Now;
                this.scrumTimer = new System.Timers.Timer(delay.TotalMilliseconds);
                this.scrumTimer.Elapsed += (sender, args) =>
                {
                    this.logger.LogInformation($"Timer matched to send notification at timer value : {this.scrumTimer}");
                    this.scrumTimer.Stop();  // reset timer
                    this.StartScrumAsync(scrumMaster);
                };

                this.scrumTimer.AutoReset = false;
                this.scrumTimer.Start();
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// Method invokes send notification task which gets channel name and send the notification.
        /// </summary>
        /// <param name="scrumMaster">values obtained from scrum master table.</param>
        /// <returns>A task that sends notification in channel for group activity.</returns>
        private async Task StartScrumAsync(ScrumMaster scrumMaster)
        {
            this.logger.LogInformation($"Send the scrum start card for {scrumMaster.ScrumMasterId}");
            await this.startScrumActivityHelper.ScrumStartActivityAsync(scrumMaster);
        }

        /// <summary>
        /// Method schedules the storage call to fetch details of scrum when timer is elapsed.
        /// </summary>
        private void ScheduleStorage()
        {
            var count = Interlocked.Increment(ref this.executionCount);
            this.logger.LogInformation($"Scheduling storage to get scrum master details. Count: {count}");

            // Schedule storage call hourly.
            CronExpression storageCronExpression = CronExpression.Parse("0 */1 * * *");
            var next = storageCronExpression.GetNextOccurrence(DateTimeOffset.Now, TimeZoneInfo.Local);
            if (next.HasValue)
            {
                var delay = next.Value - DateTimeOffset.Now;
                this.timer = new System.Timers.Timer(delay.TotalMilliseconds);
                this.timer.Elapsed += (sender, args) =>
                {
                    this.logger.LogInformation($"Timer matched to fetch scrum master details at timer value : {this.timer}");
                    this.timer.Stop();  // reset timer
                    this.GetAllScrumDetailAndScheduleAsync();
                    this.ScheduleStorage();    // reschedule next
                };

                this.timer.AutoReset = false;
                this.timer.Start();
            }
        }

        /// <summary>
        /// Get all active scrum of next hour from storage and schedule based on given start time.
        /// For ex: If current UTC time is 10.00 AM, this will fetch scrums scheduled for 11.00 AM and 11.30 AM from storage by UTC hour calculated as 11.
        /// </summary>
        /// <returns>A task that schedules scrum.</returns>
        private async Task GetAllScrumDetailAndScheduleAsync()
        {
            var scrumMasterDetails = await this.scrumMasterStorageProvider.GetActiveScrumMasterOfNextHourAsync();

            if (scrumMasterDetails != null)
            {
                foreach (var scrumMaster in scrumMasterDetails)
                {
                    CronExpression expression = CronExpressionHelper.CreateCronExpression(DateTimeOffset.Parse(scrumMaster.StartTime, CultureInfo.InvariantCulture));
                    await this.ScheduleScrumAsync(expression, scrumMaster);
                }
            }
        }
    }
}

// <copyright file="ArchivalBackgroundService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
namespace Microsoft.Teams.Apps.ScrumStatus.Common
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Cronos;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.ScrumStatus.Common.Models;
    using Microsoft.Teams.Apps.ScrumStatus.Helpers;
    using Microsoft.Teams.Apps.ScrumStatus.Models;

    /// <summary>
    /// BackgroundService class that inherits IHostedService and implements the methods related to background tasks for archival of scrum data.
    /// </summary>
    public class ArchivalBackgroundService : IHostedService, IDisposable
    {
        /// <summary>
        /// Storage helper for working with scrum data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumStorageProvider scrumStorageProvider;

        /// <summary>
        /// Storage helper for working with scrum status data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IScrumStatusStorageProvider scrumStatusStorageProvider;

        /// <summary>
        /// Instance to get the graph methods.
        /// </summary>
        private readonly IGraphUtilityHelper graphUtility;

        /// <summary>
        /// Instance for creating data table and workbook.
        /// </summary>
        private readonly ExportHelper exportHelper;

        /// <summary>
        /// Instance to log details in application insights.
        /// </summary>
        private readonly ILogger<ArchivalBackgroundService> logger;

        /// <summary>
        /// Microsoft app credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<ScrumStatusActivityHandlerOptions> options;

        /// <summary>
        /// Gets configuration setting whether to export scrum details.
        /// </summary>
        private readonly IOptionsMonitor<ExportOptions> exportOption;

        /// <summary>
        /// Flag to check whether export is enabled by the user.
        /// </summary>
        private readonly string isExportEnabled = "True";

        /// <summary>
        /// Name of excel sheet which is exported.
        /// </summary>
        private readonly string exportedSheetName = "Scrum_Report";

        /// <summary>
        /// Name of data table for which scrum status has to be exported.
        /// </summary>
        private DataTable scrumStatusExportDataTable;

        /// <summary>
        /// Execution count.
        /// </summary>
        private int executionCount = 0;

        /// <summary>
        /// Timer.
        /// </summary>
        private System.Timers.Timer archivalJobTimer;

        // Flag: Has Dispose already been called?
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="ArchivalBackgroundService"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to sending notification tasks.
        /// </summary>
        /// <param name="scrumStorageProvider">Scrum status storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="scrumStatusStorageProvider">An instance of scrum status storage provider.</param>
        /// /// <param name="graphUtility">Instance of graph utility helper.</param>
        /// <param name="exportHelper">Instance for creating data table and workbook.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="microsoftAppCredentials">MicrosoftAppCredentials of bot.</param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        /// <param name="exportOption">Configuration value to check whether to export scrum details.</param>
        public ArchivalBackgroundService(
          IScrumStorageProvider scrumStorageProvider,
          IScrumStatusStorageProvider scrumStatusStorageProvider,
          IGraphUtilityHelper graphUtility,
          ExportHelper exportHelper,
          ILogger<ArchivalBackgroundService> logger,
          MicrosoftAppCredentials microsoftAppCredentials,
          IOptions<ScrumStatusActivityHandlerOptions> options,
          IOptionsMonitor<ExportOptions> exportOption)
        {
            this.scrumStorageProvider = scrumStorageProvider;
            this.scrumStatusStorageProvider = scrumStatusStorageProvider;
            this.graphUtility = graphUtility;
            this.exportHelper = exportHelper;
            this.logger = logger;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.exportOption = exportOption;
        }

        /// <summary>
        /// Method to start the background task when application starts.
        /// </summary>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task instance.</returns>
        public Task StartAsync(CancellationToken cancellationToken)
        {
            if (this.exportOption.CurrentValue.IsExportEnabled.Equals(this.isExportEnabled, StringComparison.OrdinalIgnoreCase))
            {
                this.ScheduleArchivalJob();
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// Triggered when the host is performing a graceful shutdown.
        /// </summary>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task instance.</returns>
        public Task StopAsync(CancellationToken cancellationToken)
        {
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
                this.archivalJobTimer.Dispose();
                this.scrumStatusExportDataTable.Dispose();
            }

            this.disposed = true;
        }

        /// <summary>
        /// Get archival data from Microsoft Azure Table storage.
        /// </summary>
        /// <returns>A task that Enqueue sends notification task.</returns>
        private async Task GetArchivalDataAsync()
        {
            try
            {
                List<ScrumExport> scrumToExport = new List<ScrumExport>();
                string filePath = string.Empty;
                string fileName = string.Empty;
                var scrum = this.scrumStorageProvider.GetScrumDetailsByTimestampAsync().Result;
                if (scrum == null)
                {
                    this.logger.LogInformation("Scrum obtained is null in archival job");
                    return;
                }

                foreach (var scrumData in scrum)
                {
                    if (string.IsNullOrEmpty(scrumData.AadGroupId))
                    {
                        this.logger.LogInformation("AAD group id is null in scrum data in archival job");
                        continue;
                    }

                    var scrumStatus = await this.scrumStatusStorageProvider.GetScrumStatusBySummaryCardIdAsync(scrumData.ScrumStartCardResponseId, scrumData.AadGroupId);
                    var driveDetails = await this.graphUtility.GetDriveDetailsAsync(scrumData.AadGroupId);
                    scrumToExport = scrumStatus.Select(
                                                        scrumExport => new ScrumExport
                                                        {
                                                            DateOfScrum = scrumExport.CreatedOn,
                                                            MemberName = scrumExport.Username,
                                                            WorkedUponYesterday = scrumExport.YesterdayTaskDescription,
                                                            Blockers = scrumExport.BlockerDescription,
                                                            PlanForToday = scrumExport.TodayTaskDescription,
                                                        }).ToList();
                    fileName = this.GetCurrentTimestamp() + ".xlsx";
                    filePath = $"{scrumData.ChannelName}/ScrumReports/{scrumData.ScrumTeamConfigId.Split("_")[0]}/{fileName}";
                    using (this.scrumStatusExportDataTable = this.exportHelper.ConvertToDataTable(scrumToExport, this.exportedSheetName))
                    {
                        var uploadContext = await this.graphUtility.PutAsync(this.scrumStatusExportDataTable, filePath, driveDetails.Id);
                        if (uploadContext != null)
                        {
                            this.logger.LogInformation($"File uploaded- {fileName}");
                            if (scrumStatus.Any())
                            {
                                await this.DeleteScrumStatusAsync(scrumStatus);
                            }

                            await this.DeleteScrumAsync(scrumData);
                        }
                        else
                        {
                            this.logger.LogInformation($"Error while uploading the file- {fileName}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred while exporting data in archival job in GetArchivalDataAsync", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Method to get current time stamp used as filename
        /// </summary>
        /// <returns>A string of current time stamp with _ as delimiter.</returns>
        private string GetCurrentTimestamp()
        {
            return DateTimeOffset.Now.ToString("yyyy-MM-dd_hh-mm-ss", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Set the timer and enqueue start scrum task if timer matched as per CRON expression.
        /// </summary>
        private void ScheduleArchivalJob()
        {
            var count = Interlocked.Increment(ref this.executionCount);
            this.logger.LogInformation($"Start scrum Hosted Service is working. Count: {count}");

            // Schedule storage call monthly.
            var lastDayOfMonth = this.GetLastDayOfMonth();
            CronExpression expression = CronExpression.Parse($"0 0 15,{lastDayOfMonth} * *");
            var next = expression.GetNextOccurrence(DateTimeOffset.Now, TimeZoneInfo.Local);
            if (next.HasValue)
            {
                var delay = next.Value - DateTimeOffset.Now;
                this.archivalJobTimer = new System.Timers.Timer(delay.TotalMilliseconds);
                this.archivalJobTimer.Elapsed += (sender, args) =>
                {
                    this.archivalJobTimer.Stop();  // reset timer

                    // Export excel and delete rows if it is last day of the month.
                    if (lastDayOfMonth.Equals(next.Value.Day))
                    {
                        this.logger.LogInformation($"Last day of the month {lastDayOfMonth} and exporting the data.");
                        Task.Run(() => this.GetArchivalDataAsync()).Wait();
                    }

                    this.ScheduleArchivalJob();    // reschedule next
                };

                this.archivalJobTimer.AutoReset = false;
                this.archivalJobTimer.Start();
            }
        }

        /// <summary>
        /// Method to delete scrum status from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrumStatuses">Collection of scrum status.</param>
        private async Task DeleteScrumStatusAsync(IEnumerable<ScrumStatus> scrumStatuses)
        {
            try
            {
                foreach (var scrumStatus in scrumStatuses)
                {
                    var deleteResponse = await this.scrumStatusStorageProvider.DeleteEntityAsync(scrumStatus);
                    if (deleteResponse != null)
                    {
                        this.logger.LogInformation($"Scrum status deleted: {scrumStatus.UserAadObjectId}");
                    }
                    else
                    {
                        this.logger.LogInformation($"Scrum status deletion failed: {scrumStatus.UserAadObjectId}");
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while deleting data from ScrumStatus at DeleteScrumStatusAsync: {ex}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Method to delete scrum status from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="scrum">Scrum data.</param>
        private async Task DeleteScrumAsync(Scrum scrum)
        {
            try
            {
                var deleteResponse = await this.scrumStorageProvider.DeleteEntityAsync(scrum);
                if (deleteResponse != null)
                {
                    this.logger.LogInformation($"Scrum deleted for scrum team configuration id: {scrum.ScrumTeamConfigId}");
                }
                else
                {
                    this.logger.LogInformation($"Scrum deletion failed for scrum team configuration id: {scrum.ScrumTeamConfigId}");
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while deleting data from Scrum at DeleteScrumAsync: {ex}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Gets the last day of the month.
        /// </summary>
        /// <returns>Last day of the month.</returns>
        private int GetLastDayOfMonth()
        {
            var date = DateTime.Now;
            var daysInMonth = DateTime.DaysInMonth(date.Year, date.Month);
            var lastDay = new DateTime(date.Year, date.Month, daysInMonth);
            this.logger.LogInformation($"Last day of the month obtained is : {lastDay}");
            return lastDay.Day;
        }
    }
}

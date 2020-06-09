// <copyright file="ArchivalBackgroundService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
namespace Microsoft.Teams.Apps.ScrumStatus.Common.BackgroundService
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.ScrumStatus.Common.Models;
    using Microsoft.Teams.Apps.ScrumStatus.Helpers;
    using Microsoft.Teams.Apps.ScrumStatus.Models;

    /// <summary>
    /// BackgroundService class that inherits IHostedService and implements the methods related to background tasks for archival of scrum data.
    /// </summary>
    public sealed class ArchivalBackgroundService : BackgroundService
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
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<ScrumStatusActivityHandlerOptions> options;

        /// <summary>
        /// Gets configuration setting whether to export scrum details.
        /// </summary>
        private readonly IOptionsMonitor<ExportOptions> exportOption;

        /// <summary>
        /// Initializes a new instance of the <see cref="ArchivalBackgroundService"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to sending notification tasks.
        /// </summary>
        /// <param name="scrumStorageProvider">Scrum status storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="scrumStatusStorageProvider">An instance of scrum status storage provider.</param>
        /// /// <param name="graphUtility">Instance of graph utility helper.</param>
        /// <param name="exportHelper">Instance for creating data table and workbook.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        /// <param name="exportOption">Configuration value to check whether to export scrum details.</param>
        public ArchivalBackgroundService(
          IScrumStorageProvider scrumStorageProvider,
          IScrumStatusStorageProvider scrumStatusStorageProvider,
          IGraphUtilityHelper graphUtility,
          ExportHelper exportHelper,
          ILogger<ArchivalBackgroundService> logger,
          IOptions<ScrumStatusActivityHandlerOptions> options,
          IOptionsMonitor<ExportOptions> exportOption)
        {
            this.scrumStorageProvider = scrumStorageProvider;
            this.scrumStatusStorageProvider = scrumStatusStorageProvider;
            this.graphUtility = graphUtility;
            this.exportHelper = exportHelper;
            this.logger = logger;
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.exportOption = exportOption;
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
            this.logger.LogInformation($"Export to SharePoint archival settings found as {this.exportOption.CurrentValue.IsExportEnabled}");

            while (this.exportOption.CurrentValue.IsExportEnabled
                && !stoppingToken.IsCancellationRequested)
            {
                this.logger.LogInformation("Archival background job execution has started...");

                // get the last day of current month.
                var lastDayOfMonth = this.GetLastDayOfMonth();

                // check if the current day is the last day of the month.
                if (DateTimeOffset.UtcNow.Day == lastDayOfMonth)
                {
                    await this.GetArchivalDataAsync();
                }

                await Task.Delay(TimeSpan.FromDays(1), stoppingToken);
            }

            this.logger.LogInformation("Archival job execution has either stopped or did not executed .");
        }

        /// <summary>
        /// Get archival data from Microsoft Azure Table storage.
        /// </summary>
        /// <returns>A task that Enqueue sends notification task.</returns>
        private async Task GetArchivalDataAsync()
        {
            // Name of excel sheet which is exported.
            string exportedSheetName = "Scrum_Report";

            DataTable scrumStatusExportDataTable;
            try
            {
                List<ScrumExport> scrumToExport = new List<ScrumExport>();
                string filePath = string.Empty;
                string fileName = string.Empty;
                var scrums = await this.scrumStorageProvider.GetScrumDetailsByTimestampAsync();
                if (scrums == null || !scrums.Any())
                {
                    this.logger.LogInformation("Scrum obtained is null in archival job");
                    return;
                }

                foreach (var scrum in scrums)
                {
                    if (string.IsNullOrEmpty(scrum.AadGroupId))
                    {
                        this.logger.LogInformation("AAD group id is null in scrum data in archival job");
                        continue;
                    }

                    var scrumStatus = await this.scrumStatusStorageProvider
                        .GetScrumStatusBySummaryCardIdAsync(scrum.ScrumStartCardResponseId, scrum.AadGroupId);
                    var driveDetails = await this.graphUtility.GetDriveDetailsAsync(scrum.AadGroupId);

                    scrumToExport = scrumStatus.Select(
                        scrumExport => new ScrumExport
                        {
                            DateOfScrum = scrumExport.CreatedOn,
                            MemberName = scrumExport.Username,
                            WorkedUponYesterday = scrumExport.YesterdayTaskDescription,
                            Blockers = scrumExport.BlockerDescription,
                            PlanForToday = scrumExport.TodayTaskDescription,
                        }).ToList();

                    var scrumTeamName = scrum.ScrumTeamConfigId.Split("_")[0];
                    fileName = this.GetCurrentTimestamp() + ".xlsx";
                    filePath = $"{scrum.ChannelName}/ScrumReports/{scrumTeamName}/{fileName}";
                    using (scrumStatusExportDataTable = this.exportHelper.ConvertToDataTable(scrumToExport, exportedSheetName))
                    {
                        var uploadContext = await this.graphUtility.UploadFileAsync(scrumStatusExportDataTable, filePath, driveDetails.Id);
                        if (uploadContext != null)
                        {
                            this.logger.LogInformation($"File uploaded- {fileName}");
                            if (scrumStatus.Any())
                            {
                                await this.DeleteScrumStatusAsync(scrumStatus);
                            }

                            await this.DeleteScrumAsync(scrum);
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
        private int GetLastDayOfMonth() => DateTime.DaysInMonth(DateTime.UtcNow.Year, DateTime.UtcNow.Month);
    }
}

// <copyright file="GraphUtilityHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Helpers
{
    using System;
    using System.Data;
    using System.IO;
    using System.Threading.Tasks;
    using ClosedXML.Excel;
    using Microsoft.Extensions.Options;
    using Microsoft.Graph;
    using Microsoft.Graph.Auth;
    using Microsoft.Identity.Client;
    using Microsoft.Teams.Apps.ScrumStatus.Common;
    using Microsoft.Teams.Apps.ScrumStatus.Models.Configuration;

    /// <summary>
    /// Implements the methods that are defined in <see cref="IGraphUtilityHelper"/>.
    /// </summary>
    public class GraphUtilityHelper : IGraphUtilityHelper
    {
        /// <summary>
        /// Graph service client instance.
        /// </summary>
        private readonly GraphServiceClient graphServiceClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphUtilityHelper"/> class.
        /// </summary>
        /// <param name="appOptions">Options.</param>
        public GraphUtilityHelper(IOptions<MicrosoftAppOptions> appOptions)
        {
            appOptions = appOptions ?? throw new ArgumentNullException(nameof(appOptions));
            this.graphServiceClient = this.GetGraphServiceClient(appOptions);
        }

        /// <summary>
        /// Get drive details.
        /// </summary>
        /// <param name="groupId">Group id of the team in which channel is to be created.</param>
        /// <returns>A task that returns list of all channels in a team.</returns>
        public async Task<Drive> GetDriveDetailsAsync(string groupId)
        {
            return await this.graphServiceClient
                                    .Groups[groupId]
                                    .Drive
                                    .Request()
                                    .GetAsync();
        }

        /// <summary>
        /// Upload file to given file path in Drive.
        /// </summary>
        /// <param name="dataTable">Data table to be uploaded as excel sheet.</param>
        /// <param name="filePath">File path to upload.</param>
        /// <param name="driveId">Id of the team's drive</param>
        /// <returns>A task that represents a HTTP response message including the status code and data.</returns>
        public async Task<string> UploadFileAsync(DataTable dataTable, string filePath, string driveId)
        {
            using (XLWorkbook workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add(dataTable);
                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);

                    var request = this.graphServiceClient
                                   .Drives[driveId]
                                   .Root
                                   .ItemWithPath(filePath)
                                   .Content
                                   .Request();

                    request.Headers.Add(new HeaderOption("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

                    stream.Position = 0;
                    return (await request
                                   .PutAsync<DriveItem>(stream)).ToString();
                }
            }
        }

        /// <summary>
        /// Create graph service client instance
        /// </summary>
        /// <returns>Microsoft Graph service client</returns>
        private GraphServiceClient GetGraphServiceClient(IOptions<MicrosoftAppOptions> options)
        {
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                            .Create(options.Value.ClientId)
                            .WithTenantId(options.Value.TenantId)
                            .WithClientSecret(options.Value.ClientSecret)
                            .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);
            return new GraphServiceClient(authProvider);
        }
    }
}

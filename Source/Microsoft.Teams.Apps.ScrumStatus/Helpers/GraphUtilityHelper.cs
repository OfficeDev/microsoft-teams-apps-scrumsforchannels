// <copyright file="GraphUtilityHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Helpers
{
    using System;
    using System.IO;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.ScrumStatus.Common;
    using Microsoft.Teams.Apps.ScrumStatus.Models.Graph;
    using Microsoft.Teams.Shifts.Integration.Common.Models.Graph;
    using Newtonsoft.Json;

    /// <summary>
    /// Implements the methods that are defined in <see cref="IGraphUtilityHelper"/>.
    /// </summary>
    public class GraphUtilityHelper : IGraphUtilityHelper
    {
        /// <summary>
        /// Graph api base URL.
        /// </summary>
        private const string GraphApiBaseUrl = "https://graph.microsoft.com";

        /// <summary>
        /// Login request base URL.
        /// </summary>
        private const string LoginRequestBaseUrl = "https://login.microsoftonline.com";

        /// <summary>
        /// Provides a base class for sending HTTP requests and receiving HTTP responses from a resource identified by a URI.
        /// </summary>
        private readonly HttpClient httpClient;

        /// <summary>
        /// Instance to log details in application insights.
        /// </summary>
        private readonly ILogger<GraphUtilityHelper> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphUtilityHelper"/> class.
        /// </summary>
        /// <param name="httpClient">Instance of HttpClient</param>
        /// <param name="logger">Instance of ILogger</param>
        public GraphUtilityHelper(HttpClient httpClient, ILogger<GraphUtilityHelper> logger)
        {
            this.httpClient = httpClient;
            this.logger = logger;
        }

        /// <summary>
        /// Returns the Graph Token response object.
        /// </summary>
        /// <param name="tenantId">The TenantId.</param>
        /// <param name="clientId">The Azure AD Application ID for the Configuration Web App.</param>
        /// <param name="clientSecret">The Azure AD Application Secret for the Configuration Web App.</param>
        /// <returns>A unit of execution that has the <see cref="GraphTokenResponse"/> boxed in which contains the Graph API token.</returns>
        public async Task<GraphTokenResponse> ObtainGraphTokenAsync(string tenantId, string clientId, string clientSecret)
        {
            var requestUrl = $"{LoginRequestBaseUrl}/{tenantId}/oauth2/v2.0/token";
            var stringQuery = $"client_id={clientId}&scope={Uri.EscapeDataString($"{GraphApiBaseUrl}/.default")}&client_secret={Uri.EscapeDataString(clientSecret)}&grant_type=client_credentials";
            using (var httpContent = new StringContent(stringQuery, Encoding.UTF8, "application/x-www-form-urlencoded"))
            {
                var response = await this.httpClient.PostAsync(new Uri(requestUrl), httpContent);

                if (response.IsSuccessStatusCode)
                {
                    var responseContent = await response.Content.ReadAsStringAsync();
                    var graphTokenResponse = JsonConvert.DeserializeObject<GraphTokenResponse>(responseContent);
                    this.logger.LogInformation($"Token received: {graphTokenResponse.AccessToken}");
                    return graphTokenResponse;
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Get drive details.
        /// </summary>
        /// <param name="token">Azure Active Directory (AAD) token to access graph API.</param>
        /// <param name="groupId">groupId of the team in which channel is to be created.</param>
        /// <returns>A task that returns list of all channels in a team.</returns>
        public async Task<DriveDetails> GetDriveDetailsAsync(string token, string groupId)
        {
            var response = await this.GetAsync(token, $"{GraphApiBaseUrl}/v1.0/groups/{groupId}/drive");
            if (response.IsSuccessStatusCode)
            {
                var responseContent = await response.Content.ReadAsStringAsync();
                var driveDetailsResponse = JsonConvert.DeserializeObject<DriveDetails>(responseContent);
                return driveDetailsResponse;
            }

            var errorMessage = await response.Content.ReadAsStringAsync();
            this.logger.LogInformation($"Graph API get drive detail error: {errorMessage} statusCode: {response.StatusCode}");
            return null;
        }

        /// <summary>
        /// Method to post data to API.
        /// </summary>
        /// <param name="token">Microsoft Graph API user access token.</param>
        /// <param name="stream">In memory stream of data.</param>
        /// <param name="filePath">File path to upload.</param>
        /// <param name="driveId">Id of the team's drive</param>
        /// <returns>A task that represents a HTTP response message including the status code and data.</returns>
        public async Task<string> PutAsync(string token, MemoryStream stream, string filePath, string driveId)
        {
            var url = $"{GraphApiBaseUrl}/v1.0/drives/{driveId}//root:/{filePath}:/content";
            this.httpClient.DefaultRequestHeaders.Remove("Authorization");
            this.httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

            byte[] fileContents = stream?.ToArray();
            using (var content = new ByteArrayContent(fileContents))
            {
                content.Headers.Add("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                return await this.httpClient.PutAsync(new Uri(url), content).Result.Content.ReadAsStringAsync();
            }
        }

        /// <summary>
        /// Method to get data from API.
        /// </summary>
        /// <param name="token">Microsoft Graph API user access token.</param>
        /// <param name="requestUrl">Microsoft Graph API request URL.</param>
        /// <returns>A task that represents a HTTP response message including the status code and data.</returns>
        private async Task<HttpResponseMessage> GetAsync(string token, string requestUrl)
        {
            HttpMethod httpMethod = new HttpMethod("GET");
            using (var request = new HttpRequestMessage(httpMethod, requestUrl))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                return await this.httpClient.SendAsync(request);
            }
        }
    }
}

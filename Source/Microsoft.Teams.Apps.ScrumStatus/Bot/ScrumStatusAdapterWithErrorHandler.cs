// <copyright file="ScrumStatusAdapterWithErrorHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Bot
{
    using System;
    using System.Threading;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.ScrumStatus;

    /// <summary>
    /// Implements Error Handler.
    /// </summary>
    public class ScrumStatusAdapterWithErrorHandler : BotFrameworkHttpAdapter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ScrumStatusAdapterWithErrorHandler"/> class.
        /// </summary>
        /// <param name="configuration">Application configurations.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="scrumStatusActivityMiddleware">Represents middle ware that can operate on incoming activities.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="conversationState">Conversation state.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        public ScrumStatusAdapterWithErrorHandler(
            IConfiguration configuration,
            ILogger<IBotFrameworkHttpAdapter> logger,
            ScrumStatusActivityMiddleware scrumStatusActivityMiddleware,
            IStringLocalizer<Strings> localizer,
            ConversationState conversationState = null,
            CancellationToken cancellationToken = default)
            : base(configuration)
        {
            if (scrumStatusActivityMiddleware == null)
            {
                throw new NullReferenceException(nameof(scrumStatusActivityMiddleware));
            }

            // Add activity middle ware to the adapter's middle ware pipeline
            this.Use(scrumStatusActivityMiddleware);

            this.OnTurnError = async (turnContext, exception) =>
            {
                // Log any leaked exception from the application.
                logger.LogError(exception, $"Exception caught : {exception.Message}");

                // Send a catch-all apology to the user.
                await turnContext.SendActivityAsync(localizer.GetString("ErrorMessage"), cancellationToken: cancellationToken);

                if (conversationState != null)
                {
                    try
                    {
                        // Delete the conversationState for the current conversation to prevent the
                        // bot from getting stuck in a error-loop caused by being in a bad state.
                        // ConversationState should be thought of as similar to "cookie-state" in a Web pages.
                        await conversationState.DeleteAsync(turnContext);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, $"Exception caught on attempting to Delete ConversationState : {ex.Message}");
                    }
                }
            };
        }
    }
}
// <copyright file="CronExpressionHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Helpers
{
    using System;
    using Cronos;

    /// <summary>
    /// Class for creating CRON expression.
    /// </summary>
    public static class CronExpressionHelper
    {
        /// <summary>
        /// Creates CRON expression based on given date time.
        /// </summary>
        /// <param name="scrumTime">Time to start the scrum.</param>
        /// <returns>CRON expression</returns>
        public static CronExpression CreateCronExpression(DateTimeOffset scrumTime)
        {
            int hourofTheDay = scrumTime.Hour;
            int mintuesOftheDay = scrumTime.Minute;

            // CRON Expression to send start scrum based on start time and on every weekdays except weekends.
            CronExpression expression = CronExpression.Parse($"{mintuesOftheDay} {hourofTheDay} * * 1-5");
            return expression;
        }
    }
}

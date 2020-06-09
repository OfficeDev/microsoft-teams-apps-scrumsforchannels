// <copyright file="ServicesExtension.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using Microsoft.AspNetCore.Authentication.JwtBearer;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Localization;
    using Microsoft.Bot.Builder.BotFramework;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Teams.Apps.ScrumStatus.Common;
    using Microsoft.Teams.Apps.ScrumStatus.Common.BackgroundService;
    using Microsoft.Teams.Apps.ScrumStatus.Common.Models;
    using Microsoft.Teams.Apps.ScrumStatus.Helpers;
    using Microsoft.Teams.Apps.ScrumStatus.Models.Configuration;

    /// <summary>
    /// Class to extend ServiceCollection.
    /// </summary>
    public static class ServicesExtension
    {
        /// <summary>
        /// Adds application configuration settings to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddConfigurationSettings(this IServiceCollection services, IConfiguration configuration)
        {
            services.Configure<ScrumStatusActivityHandlerOptions>(options =>
            {
                options.TenantId = configuration.GetValue<string>("Bot:TenantId");
                options.AppBaseUri = configuration.GetValue<string>("Bot:AppBaseUri");
            });
            services.Configure<TokenOptions>(options =>
            {
                options.SecurityKey = configuration.GetValue<string>("Token:SecurityKey");
            });
            services.Configure<StorageOptions>(options =>
            {
                options.ConnectionString = configuration.GetValue<string>("Storage:ConnectionString");
            });
            services.Configure<ExportOptions>(options =>
            {
                options.IsExportEnabled = configuration.GetValue<bool>("Export:IsExportEnabled");
            });
            services.Configure<MicrosoftAppOptions>(options =>
            {
                options.ClientId = configuration.GetValue<string>("MicrosoftAppId");
                options.ClientSecret = configuration.GetValue<string>("MicrosoftAppPassword");
                options.TenantId = configuration.GetValue<string>("Bot:TenantId");
            });
        }

        /// <summary>
        /// Adds providers to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        public static void AddProviders(this IServiceCollection services)
        {
            services
                .AddTransient<IScrumConfigurationStorageProvider, ScrumConfigurationStorageProvider>();
            services
                .AddTransient<IScrumStatusStorageProvider, ScrumStatusStorageProvider>();
            services
                .AddTransient<IScrumStorageProvider, ScrumStorageProvider>();
            services
               .AddHostedService<ArchivalBackgroundService>();
            services
                .AddHostedService<ScrumNotificationBackgroundService>();
        }

        /// <summary>
        /// Adds helpers to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddHelpers(this IServiceCollection services, IConfiguration configuration)
        {
            services
                .AddApplicationInsightsTelemetry(configuration.GetValue<string>("ApplicationInsights:InstrumentationKey"));
            services
                .AddSingleton<ITokenHelper, TokenHelper>();
            services
                .AddSingleton<IStartScrumActivityHelper, StartScrumActivityHelper>();
            services
                .AddSingleton<IGraphUtilityHelper, GraphUtilityHelper>();
            services
                .AddSingleton<CardHelper>();
            services
                .AddSingleton<ActivityHelper>();
            services
                .AddSingleton<ScrumHelper>();
            services
                .AddSingleton<ExportHelper>();
        }

        /// <summary>
        /// Adds custom JWT authentication to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddCustomJWTAuthentication(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
                .AddJwtBearer(options =>
                {
                    options.TokenValidationParameters = new TokenValidationParameters
                    {
                        ValidateAudience = true,
                        ValidAudiences = new List<string> { configuration.GetValue<string>("Bot:AppBaseUri") },
                        ValidIssuers = new List<string> { configuration.GetValue<string>("Bot:AppBaseUri") },
                        ValidateIssuer = true,
                        ValidateIssuerSigningKey = true,
                        IssuerSigningKey = new SymmetricSecurityKey(Encoding.ASCII.GetBytes(configuration.GetValue<string>("Token:SecurityKey"))),
                        RequireExpirationTime = true,
                        ValidateLifetime = true,
                        ClockSkew = TimeSpan.FromSeconds(30),
                    };
                });
        }

        /// <summary>
        /// Adds credential providers for authentication.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddCredentialProviders(this IServiceCollection services, IConfiguration configuration)
        {
            services
                .AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();
            services
                .AddSingleton(new MicrosoftAppCredentials(configuration.GetValue<string>("MicrosoftAppId"), configuration.GetValue<string>("MicrosoftAppPassword")));
        }

        /// <summary>
        /// Adds localization settings to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddLocalizationSettings(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddLocalization(options => options.ResourcesPath = "Resources");
            services.Configure<RequestLocalizationOptions>(options =>
            {
                var defaultCulture = CultureInfo.GetCultureInfo(configuration.GetValue<string>("i18n:DefaultCulture"));
                var supportedCultures = configuration.GetValue<string>("i18n:SupportedCultures").Split(',')
                    .Select(culture => CultureInfo.GetCultureInfo(culture))
                    .ToList();

                options.DefaultRequestCulture = new RequestCulture(defaultCulture);
                options.SupportedCultures = supportedCultures;
                options.SupportedUICultures = supportedCultures;

                options.RequestCultureProviders = new List<IRequestCultureProvider>
                {
                    new BotLocalizationCultureProvider(),
                };
            });
        }
    }
}

// <copyright file="Startup.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using Microsoft.AspNetCore.Authentication.JwtBearer;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Localization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Azure.Cosmos;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration;
    using Microsoft.Teams.App.VirtualConsult.Common.Repositories;

    /// <summary>
    /// The Startup class is reponsible for configuring the DI container and acts as the composition root.
    /// </summary>
    public sealed class Startup
    {
        private readonly IConfiguration configuration;

        /// <summary>
        /// Initializes a new instance of the <see cref="Startup"/> class.
        /// </summary>
        /// <param name="configuration">The environment provided configuration.</param>
        public Startup(IConfiguration configuration)
        {
            this.configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        }

        /// <summary>
        /// Configure the composition root for the application.
        /// </summary>
        /// <param name="services">The stub composition root.</param>
        /// <remarks>
        /// For more information see: https://go.microsoft.com/fwlink/?LinkID=398940.
        /// </remarks>
#pragma warning disable CA1506 // Composition root expected to have coupling with many components.
        public void ConfigureServices(IServiceCollection services)
        {
            #if DEBUG
            Microsoft.IdentityModel.Logging.IdentityModelEventSource.ShowPII = true;
            #endif

            // Get app configuration sections
            var botSection = this.configuration.GetSection("Bot");
            var i18nSection = this.configuration.GetSection("i18n");
            var cosmosDBSection = this.configuration.GetSection("CosmosDb");
            var azureADSection = this.configuration.GetSection("AzureAD");
            var teamsSection = this.configuration.GetSection("Teams");

            // Configure the service collection
            services.AddOptions();
            services.Configure<BotSettings>(botSection);
            services.Configure<I18nSettings>(i18nSection);
            services.Configure<CosmosDBSettings>(cosmosDBSection);
            services.Configure<AzureADSettings>(azureADSection);
            services.Configure<TeamsSettings>(teamsSection);

            // Add credential provider for bot based on bot settings configured above
            var botSettings = botSection.Get<BotSettings>();
            var appId = botSettings.Id;
            var appPassword = botSettings.Password;
            ICredentialProvider credentialProvider = new SimpleCredentialProvider(
                appId: appId,
                password: appPassword);

            services
                .AddSingleton(credentialProvider);

            services.AddApplicationInsightsTelemetry();

            services
                .AddTransient<IBotFrameworkHttpAdapter, BotFrameworkHttpAdapter>();

            services.AddMvc();

            /*
            // Create storage and state for bot
            var storageConnectionString = Configuration.GetConnectionString("AzureStorage");
            var storage = new AzureBlobStorage(storageConnectionString, "state");
            var globalState = new GlobalState(storage, "global");
            services.AddSingleton(globalState);
            */

            services
                .AddTransient<IBot, Bot.BotActivityHandler>();

            // Enable CORS
            services.AddCors(options =>
            {
                options.AddPolicy("Default", builder =>
                {
                    builder
                        .AllowAnyOrigin()
                        .AllowAnyMethod()
                        .AllowAnyHeader();
                });
            });
            services
                .AddMvc()
                .SetCompatibilityVersion(CompatibilityVersion.Version_2_1);

            // Add i18n.
            services.AddLocalization(options => options.ResourcesPath = "Resources");
            services.Configure<RequestLocalizationOptions>(options =>
            {
                var i18nSettings = i18nSection.Get<I18nSettings>();
                var defaultCulture = CultureInfo.GetCultureInfo(i18nSettings.DefaultCulture);
                var supportedCultures = i18nSettings.SupportedCultures
                    .Split(',')
                    .Select(culture => CultureInfo.GetCultureInfo(culture.Trim()))
                    .ToList();

                options.DefaultRequestCulture = new RequestCulture(defaultCulture);
                options.SupportedCultures = supportedCultures;
                options.SupportedUICultures = supportedCultures;

                // Remove all default culture providers, to ensure that default culture is always set
                options.RequestCultureProviders = Array.Empty<IRequestCultureProvider>();
            });

            services.AddAuthentication(o =>
            {
                 o.DefaultScheme = JwtBearerDefaults.AuthenticationScheme;
            })
            .AddJwtBearer(o =>
            {
                var azureADSettings = azureADSection.Get<AzureADSettings>();
                o.Authority = $"https://sts.windows.net/{azureADSettings.TenantId}/";
                o.TokenValidationParameters = new Microsoft.IdentityModel.Tokens.TokenValidationParameters
                {
                    // Both App ID URI and client id are valid audiences in the access token
                    ValidAudiences = new List<string>
                    {
                        azureADSettings.AppId,
                        $"api://{azureADSettings.AppId}",
                        $"api://{azureADSettings.HostDomain}/{azureADSettings.AppId}",
                    },
                };
            });

            // Initialize Cosmos client
            var cosmosDBSettings = cosmosDBSection.Get<CosmosDBSettings>();
            var cosmosClientOptions = new CosmosClientOptions
            {
                SerializerOptions = new CosmosSerializationOptions
                {
                    PropertyNamingPolicy = CosmosPropertyNamingPolicy.CamelCase,
                },
            };
            var cosmosClient = new CosmosClient(cosmosDBSettings.ConnectionString, cosmosClientOptions);
            services.AddSingleton<CosmosClient>(cosmosClient);

            services.AddTransient<IRequestRepository<CosmosItemKey>, RequestRepository>();
            services.AddTransient<IAgentRepository<CosmosItemKey>, AgentRepository>();
            services.AddTransient<IChannelRepository<CosmosItemKey>, ChannelRepository>();
            services.AddTransient<IChannelMappingRepository<CosmosItemKey>, ChannelMappingRepository>();
        }
#pragma warning restore CA1506

        /// <summary>
        /// Configure the application request pipeline.
        /// </summary>
        /// <param name="app">The application.</param>
#pragma warning disable CA1822 // This method is provided by the framework
        public void Configure(IApplicationBuilder app)
#pragma warning restore CA1822
        {
            app.UseRequestLocalization();
            app.UseStaticFiles();
            app.UseCors("Default");
            app.UseAuthentication();
            app.UseMvc(routes =>
            {
                routes.MapRoute(
                    name: "default",
                    template: "{controller=Home}/{action=Index}/{id?}");

                routes.MapSpaFallbackRoute(
                    name: "spa-fallback",
                    defaults: new { controller = "Home", action = "Index" });
            });
        }
    }
}
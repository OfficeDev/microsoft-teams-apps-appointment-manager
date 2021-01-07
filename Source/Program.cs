// <copyright file="Program.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]

namespace Microsoft.Teams.App.VirtualConsult
{
    using Microsoft.AspNetCore;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.Extensions.Configuration;

    /// <summary>
    /// The Program class is responsible for holding the entrypoint of the program.
    /// </summary>
    public static class Program
    {
        /// <summary>
        /// The entrypoint for the program.
        /// </summary>
        /// <param name="args">The command line arguments.</param>
        public static void Main(string[] args)
        {
            CreateWebHostBuilder(args).Build().Run();
        }

        /// <summary>
        /// Build the webhost for servicing HTTP requests.
        /// </summary>
        /// <param name="args">The command line arguments.</param>
        /// <returns> The WebHostBuilder configured from the arguments with the composition root defined in <see cref="Startup" />.</returns>
        public static IWebHostBuilder CreateWebHostBuilder(string[] args) =>
            WebHost
                .CreateDefaultBuilder(args)
                .ConfigureAppConfiguration((hostingContext, config) =>
                {
                    config
                        .AddEnvironmentVariables();

                    if (hostingContext.HostingEnvironment.IsDevelopment())
                    {
                        // Using dotnet secrets to store the settings during development
                        // https://docs.microsoft.com/en-us/aspnet/core/security/app-secrets?view=aspnetcore-3.0&tabs=windows
                        config.AddUserSecrets<Startup>();
                    }
                })
                .UseStartup<Startup>();
    }
}

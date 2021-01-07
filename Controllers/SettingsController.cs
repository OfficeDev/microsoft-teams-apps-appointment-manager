// <copyright file="SettingsController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Controllers
{
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Responses;

    /// <summary>
    /// Web API for getting general app settings
    /// </summary>
    [ApiController]
    public class SettingsController : ControllerBase
    {
        private readonly IOptions<I18nSettings> i18nOptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="SettingsController"/> class.
        /// </summary>
        /// <param name="i18nOptions">i18n configuration options.</param>
        public SettingsController(IOptions<I18nSettings> i18nOptions)
        {
            this.i18nOptions = i18nOptions;
        }

        /// <summary>
        /// Get general settings for the app.
        /// </summary>
        /// <response code="200">The app's general settings.</response>
        /// <returns>IActionResult.</returns>
        [HttpGet]
        [Route("/api/settings")]
        [ResponseCache(Duration = 3600)]
        public IActionResult Get()
        {
            return this.Ok(new GetSettingsResponse
            {
                DefaultLocale = this.i18nOptions.Value.DefaultCulture,
            });
        }
    }
}

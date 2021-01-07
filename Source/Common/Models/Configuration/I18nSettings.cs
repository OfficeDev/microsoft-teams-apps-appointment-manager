// <copyright file="I18nSettings.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration
{
    using System.Collections.Generic;

    /// <summary>
    /// Class I18nSettings.
    /// </summary>
    public class I18nSettings
    {
        /// <summary>
        /// Gets or sets the app's default culture.
        /// </summary>
        /// <value>The default culture.</value>
        public string DefaultCulture { get; set; }

        /// <summary>
        /// Gets or sets the app's supported cultures.
        /// </summary>
        /// <value>The supported cultures (comma-delimited).</value>
        public string SupportedCultures { get; set; }
    }
}

// <copyright file="GetSettingsResponse.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models.Responses
{
    /// <summary>
    /// The model representing the response body of the GetSettings endpoint.
    /// </summary>
    public class GetSettingsResponse
    {
        /// <summary>
        /// Gets or sets the default locale for the app.
        /// </summary>
        public string DefaultLocale { get; set; }
    }
}

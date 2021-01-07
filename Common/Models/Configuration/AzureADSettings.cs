// <copyright file="AzureADSettings.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration
{
    /// <summary>
    /// Class AzureAD Settings.
    /// </summary>
    public class AzureADSettings
    {
        /// <summary>
        /// Gets or sets the Tenant ID for the Azure AD app registration.
        /// </summary>
        /// <value>The tenant id.</value>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets the App ID for the Azure AD app registration.
        /// </summary>
        /// <value>The app id.</value>
        public string AppId { get; set; }

        /// <summary>
        /// Gets or sets the App Password/Secret for the Azure AD app registration.
        /// </summary>
        /// <value>The app password.</value>
        public string AppPassword { get; set; }

        /// <summary>
        /// Gets or sets the Host Domain for the Azure AD app registration that is also the permission scope prefix.
        /// </summary>
        /// <value>The name of the host domain.</value>
        public string HostDomain { get; set; }
    }
}

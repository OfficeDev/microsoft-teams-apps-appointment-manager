// <copyright file="DeepLinkUtil.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Utils
{
    using System;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;

    /// <summary>
    /// Utility class for generating a deep link for Microsoft Teams.
    /// </summary>
    public static class DeepLinkUtil
    {
        /// <summary>
        /// Generates a Teams deep link to a static tab.
        /// </summary>
        /// <param name="appId">The application id of the teams application.</param>
        /// <param name="tabEntityId">The entity id defined for the static tab in manifest.</param>
        /// <param name="label">The label passed in the deep link.</param>
        /// <param name="route">The absolute route to the static tab view.</param>
        /// <param name="subEntityId">Optional subEntityId to pass to the static tab.</param>
        /// <param name="channelId">Optional channelId to pass to the static tab.</param>
        /// <returns>Deep link to static tab.</returns>
        public static string GetStaticTabDeepLink(string appId, TabEntityId tabEntityId, string label, string route, string subEntityId = null, string channelId = null)
        {
            // Ensure appId is not empty
            if (string.IsNullOrEmpty(appId))
            {
                throw new ArgumentNullException();
            }

            // Ensure appId is valid
            Guid appIdGuid;
            if (!Guid.TryParse(appId, out appIdGuid))
            {
                throw new ArgumentException();
            }

            // Return the deeplink
            return $"https://teams.microsoft.com/l/entity/{appId}/{tabEntityId.ToString().ToLower()}?webUrl={route}&label={label}&context=%7B%22subEntityId%22%3A+%22{subEntityId}%22%2C+%22channelId%22%{channelId}%22%22%7D";
        }
    }
}
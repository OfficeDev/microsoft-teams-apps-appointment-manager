// <copyright file="ActivityType.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    /// <summary>
    /// Enumeration for the type of activity that occurred on a consult request.
    /// </summary>
    [JsonConverter(typeof(StringEnumConverter))]
    public enum ActivityType
    {
        /// <summary>
        /// The consult request was assigned.
        /// </summary>
        Assigned,

        /// <summary>
        /// The consult request was requested to be reassigned.
        /// </summary>
        ReassignRequested,

        /// <summary>
        /// The consult request was completed.
        /// </summary>
        Completed,
    }
}
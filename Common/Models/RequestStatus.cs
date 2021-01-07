// <copyright file="RequestStatus.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    /// <summary>
    /// Enumeration for the status of a consult request.
    /// </summary>
    [JsonConverter(typeof(StringEnumConverter))]
    public enum RequestStatus
    {
        /// <summary>
        /// The consult request has not been assigned yet.
        /// </summary>
        Unassigned,

        /// <summary>
        /// The consult request has been assigned.
        /// </summary>
        Assigned,

        /// <summary>
        /// The consult request is awaiting reassignment.
        /// </summary>
        ReassignRequested,

        /// <summary>
        /// The consult request has been completed.
        /// </summary>
        Completed,
    }
}

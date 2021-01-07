// <copyright file="Agent.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System.Text;

    /// <summary>
    /// Model that describes the agent.
    /// </summary>
    public class Agent : BaseModel
    {
        /// <summary>
        /// Gets or Sets UserPrincipalName.
        /// </summary>
        public string UserPrincipalName { get; set; }

        /// <summary>
        /// Gets or Sets AADObjectId.
        /// </summary>
        public string AADObjectId { get; set; }

        /// <summary>
        /// Gets or Sets TeamId.
        /// </summary>
        public string TeamsId { get; set; }

        /// <summary>
        /// Gets or Sets Name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or Sets ServiceUrl.
        /// </summary>
        public string ServiceUrl { get; set; }

        /// <summary>
        /// Gets or Sets BookingsStaffMemberId.
        /// </summary>
        public string BookingsStaffMemberId { get; set; }

        /// <summary>
        /// Gets or Sets Locale.
        /// </summary>
        public string Locale { get; set; }
    }
}

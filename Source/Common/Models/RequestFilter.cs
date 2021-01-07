// <copyright file="RequestFilter.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;

    /// <summary>
    /// Model that describes the consult request submitted by the customer.
    /// </summary>
    public class RequestFilter
    {
        /// <summary>
        /// Gets or Sets Categories.
        /// </summary>
        [Required]
        public List<string> Categories { get; set; }

        /// <summary>
        /// Gets or Sets Statuses.
        /// </summary>
        [Required]
        public List<RequestStatus> Statuses { get; set; }
    }
}

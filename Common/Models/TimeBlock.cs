// <copyright file="TimeBlock.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.Text;

    /// <summary>
    /// Model that descibes the time slot the customer is available for a given day.
    /// </summary>
    public class TimeBlock
    {
        /// <summary>
        /// Gets or Sets StartDateTime.
        /// </summary>
        [Required]
        public DateTimeOffset StartDateTime { get; set; }

        /// <summary>
        /// Gets or Sets EndDateTime.
        /// </summary>
        [Required]
        public DateTimeOffset EndDateTime { get; set; }
    }
}

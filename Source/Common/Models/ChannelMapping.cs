// <copyright file="ChannelMapping.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.Text;

    /// <summary>
    /// Model that describes the Team channel mappings of category to channel id.
    /// </summary>
    public class ChannelMapping : BaseModel
    {
        /// <summary>
        /// Gets or Sets ChannelId.
        /// </summary>
        [Required]
        public string ChannelId { get; set; }

        /// <summary>
        /// Gets or Sets Category.
        /// </summary>
        [Required]
        public string Category { get; set; }

        /// <summary>
        /// Gets or Sets BookingBusiness.
        /// </summary>
        public IdName BookingsBusiness { get; set; }

        /// <summary>
        /// Gets or Sets BookingsService.
        /// </summary>
        public IdName BookingsService { get; set; }

        /// <summary>
        /// Gets or Sets Supervisors.
        /// </summary>
        public List<IdName> Supervisors { get; set; }
    }
}

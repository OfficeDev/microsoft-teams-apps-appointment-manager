// <copyright file="MeetingDetails.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System.ComponentModel.DataAnnotations;
    using System.Text;

    /// <summary>
    /// Model that describes the meeting details.
    /// </summary>
    public class MeetingDetails
    {
        /// <summary>
        /// Gets or Sets Meeting Subject.
        /// </summary>
        [Required]
        public string Subject { get; set; }

        /// <summary>
        /// Gets or Sets MeetingTime.
        /// </summary>
        [Required]
        public TimeBlock MeetingTime { get; set; }
    }
}

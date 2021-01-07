// <copyright file="Activity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System.ComponentModel.DataAnnotations;
    using System.Text;

    /// <summary>
    /// Model that describes an activity that was performed on the consult request.
    /// </summary>
    public class Activity : CreatedByUserBaseModel
    {
        /// <summary>
        /// Gets or sets the type of activity that occurred.
        /// </summary>
        [Required]
        public ActivityType Type { get; set; }

        /// <summary>
        /// Gets or sets the ID of the user on whom the activity was performed.
        /// </summary>
        /// <remarks>
        /// For example, if user A assigns a consult to user B, then this field represents user B.
        /// For user A, use <see cref="CreatedByUserBaseModel.CreatedById"/>.
        /// </remarks>
        [Required]
        public string ActivityForUserId { get; set; }

        /// <summary>
        /// Gets or sets the name of the user on whom the activity was performed.
        /// </summary>
        /// <remarks>
        /// Corresponds to the same user as <see cref="ActivityForUserId"/>.
        /// </remarks>
        [Required]
        public string ActivityForUserName { get; set; }

        /// <summary>
        /// Gets or sets a comment entered by the user who performed the activity.
        /// </summary>
        public string Comment { get; set; }
    }
}

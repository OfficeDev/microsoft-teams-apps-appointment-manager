// <copyright file="Request.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.Text;

    /// <summary>
    /// Model that describes the consult request submitted by the customer.
    /// </summary>
    public class Request : BaseModel
    {
        /// <summary>
        /// Gets or Sets CustomerName.
        /// </summary>
        [Required]
        public string CustomerName { get; set; }

        /// <summary>
        /// Gets or Sets CustomerPhone.
        /// </summary>
        [Required]
        public string CustomerPhone { get; set; }

        /// <summary>
        /// Gets or Sets CustomerEmail.
        /// </summary>
        [Required]
        public string CustomerEmail { get; set; }

        /// <summary>
        /// Gets or Sets Query.
        /// </summary>
        [Required]
        public string Query { get; set; }

        /// <summary>
        /// Gets or Sets PreferredTimes.
        /// </summary>
        [Required]
        public List<TimeBlock> PreferredTimes { get; set; }

        /// <summary>
        /// Gets or Sets FriendlyId.
        /// </summary>
        public string FriendlyId { get; set; }

        /// <summary>
        /// Gets or Sets Category.
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// Gets or Sets Status.
        /// </summary>
        public RequestStatus Status { get; set; }

        /// <summary>
        /// Gets or Sets AssignedToId.
        /// </summary>
        public string AssignedToId { get; set; }

        /// <summary>
        /// Gets or Sets AssignedToName.
        /// </summary>
        public string AssignedToName { get; set; }

        /// <summary>
        /// Gets or Sets AssignedTimeBlock.
        /// </summary>
        public TimeBlock AssignedTimeBlock { get; set; }

        /// <summary>
        /// Gets or Sets BookingsBusinessId.
        /// </summary>
        public string BookingsBusinessId { get; set; }

        /// <summary>
        /// Gets or Sets BookingsServiceId.
        /// </summary>
        public string BookingsServiceId { get; set; }

        /// <summary>
        /// Gets or Sets BookingsAppointmentId.
        /// </summary>
        public string BookingsAppointmentId { get; set; }

        /// <summary>
        /// Gets or Sets JoinUri.
        /// </summary>
        public string JoinUri { get; set; }

        /// <summary>
        /// Gets or Sets Activities.
        /// </summary>
        public List<Activity> Activities { get; set; }

        /// <summary>
        /// Gets or Sets Activity Ids.
        /// </summary>
        public string ActivityId { get; set; }

        /// <summary>
        /// Gets or Sets Conversation Ids.
        /// </summary>
        public string ConversationId { get; set; }

        /// <summary>
        /// Gets or Sets Notes.
        /// </summary>
        public List<Note> Notes { get; set; }

        /// <summary>
        /// Gets or Sets Attachments.
        /// </summary>
        public List<Attachment> Attachments { get; set; }
    }
}

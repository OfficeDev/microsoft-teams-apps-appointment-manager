// <copyright file="AssignConsultRequestBody.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models.Requests
{
    using System.ComponentModel.DataAnnotations;

    /// <summary>
    /// The model representing the request body of the AssignConsult endpoint.
    /// </summary>
    public class AssignConsultRequestBody
    {
        /// <summary>
        /// Gets or sets SelectedTimeBlock.
        /// </summary>
        [Required]
        public TimeBlock SelectedTimeBlock { get; set; }

        /// <summary>
        /// Gets or sets Comments.
        /// </summary>
        [Required(AllowEmptyStrings = true)]
        public string Comments { get; set; }

        /// <summary>
        /// Gets or sets Agent.
        /// </summary>
        public IdName Agent { get; set; }
    }
}

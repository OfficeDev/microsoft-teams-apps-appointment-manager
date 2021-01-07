// <copyright file="ReassignConsultRequestBody.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models.Requests
{
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;

    /// <summary>
    /// The model representing the request body of the ReassignConsult endpoint.
    /// </summary>
    public class ReassignConsultRequestBody
    {
        /// <summary>
        /// Gets or sets requested agents list.
        /// </summary>
        [Required]
        public List<IdName> Agents { get; set; }

        /// <summary>
        /// Gets or sets agent's Comments.
        /// </summary>
        public string Comments { get; set; }
    }
}

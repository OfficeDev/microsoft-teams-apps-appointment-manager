// <copyright file="GraphResponse.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Utils
{
    using System.ComponentModel.DataAnnotations;

    /// <summary>
    /// The model that represents the Graph call responses.
    /// </summary>
    /// <typeparam name="T">The result type from Graph.</typeparam>
    public class GraphResponse<T>
    {
        /// <summary>
        /// Gets or sets why the Graph call may have failed.
        /// </summary>
        [Required]
        public string FailureReason { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the Graph call response.
        /// </summary>
        [Required]
        public T Result { get; set; }
    }
}

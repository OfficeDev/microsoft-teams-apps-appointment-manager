// <copyright file="IdName.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System.ComponentModel.DataAnnotations;
    using System.Text;

    /// <summary>
    /// Model that describes a simple entity with string id and name.
    /// </summary>
    public class IdName
    {
        /// <summary>
        /// Gets or Sets Id.
        /// </summary>
        [Required]
        public string Id { get; set; }

        /// <summary>
        /// Gets or Sets DisplayName.
        /// </summary>
        [Required]
        public string DisplayName { get; set; }
    }
}

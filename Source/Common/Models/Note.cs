// <copyright file="Note.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System.ComponentModel.DataAnnotations;
    using System.Text;

    /// <summary>
    /// Model that describes the notes attached a consult request.
    /// </summary>
    public class Note : CreatedByUserBaseModel
    {
        /// <summary>
        /// Gets or Sets Text.
        /// </summary>
        [Required]
        public string Text { get; set; }
    }
}

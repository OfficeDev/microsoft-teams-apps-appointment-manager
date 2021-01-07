// <copyright file="Attachment.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System.ComponentModel.DataAnnotations;
    using System.Text;

    /// <summary>
    /// Model that describes the attachements created for a consult request.
    /// </summary>
    public class Attachment : CreatedByUserBaseModel
    {
        /// <summary>
        /// Gets or sets the name of the file attached.
        /// </summary>
        [Required]
        public string Filename { get; set; }

        /// <summary>
        /// Gets or sets the URI where the attached file lives.
        /// </summary>
        [Required]
        public string Uri { get; set; }

        /// <summary>
        /// Gets or sets the title of the attachment set by the agent.
        /// </summary>
        [Required]
        public string Title { get; set; }
    }
}

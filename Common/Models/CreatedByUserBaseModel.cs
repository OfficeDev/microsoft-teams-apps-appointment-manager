// <copyright file="CreatedByUserBaseModel.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System;

    /// <summary>
    /// Base class that provides model properties for objects that are created by users.
    /// </summary>
    public abstract class CreatedByUserBaseModel : BaseModel
    {
        /// <summary>
        /// Gets or sets the ID of the user that created the object.
        /// </summary>
        public Guid CreatedById { get; set; }

        /// <summary>
        /// Gets or sets the name of the user that created the object.
        /// </summary>
        /// <remarks>
        /// Corresponds to the same user as <see cref="CreatedById"/>.
        /// </remarks>
        public string CreatedByName { get; set; }
    }
}
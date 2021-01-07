// <copyright file="BaseModel.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System;

    /// <summary>
    /// Base class that provides basic model properties.
    /// </summary>
    public abstract class BaseModel
    {
        /// <summary>
        /// Gets or sets the ID of the object.
        /// </summary>
        /// <value>The identifier.</value>
        public Guid Id { get; set; }

        /// <summary>
        /// Gets or sets the datetime when the object was created.
        /// </summary>
        public DateTime CreatedDateTime { get; set; }
    }
}
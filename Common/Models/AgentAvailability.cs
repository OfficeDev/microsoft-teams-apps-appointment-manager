// <copyright file="AgentAvailability.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.Text;

    /// <summary>
    /// The model that represents the Agent availability.
    /// </summary>
    public class AgentAvailability
    {
        /// <summary>
        /// Gets or sets id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets displayName.
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets email.
        /// </summary>
        public string Emailaddress { get; set; }

        /// <summary>
        /// Gets or Sets StartDateTime.
        /// </summary>
        public List<TimeBlock> TimeBlocks { get; set; }

        /// <summary>
        /// Returns the string presentation of the object.
        /// </summary>
        /// <returns>String presentation of the object.</returns>
        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("class TimeBlock {\n");
            sb.Append("  Id: ").Append(this.Id).Append("\n");
            sb.Append("  DisplayName: ").Append(this.DisplayName).Append("\n");
            sb.Append("  Emailaddress: ").Append(this.Emailaddress).Append("\n");
            sb.Append("}\n");
            return sb.ToString();
        }
    }
}

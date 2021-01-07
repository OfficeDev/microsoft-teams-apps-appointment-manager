// <copyright file="Channel.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System.ComponentModel.DataAnnotations;
    using System.Text;
    using Newtonsoft.Json;

    /// <summary>
    /// Model that describes the Team channels where the consult bot is added to.
    /// </summary>
    public class Channel : BaseModel
    {
        /// <summary>
        /// Gets or Sets TenantId.
        /// </summary>
        [Required]
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or Sets ServiceUrl.
        /// </summary>
        [Required]
        public string ServiceUrl { get; set; }

        /// <summary>
        /// Gets or Sets TeamId.
        /// </summary>
        [Required]
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or Sets TeamAADObjectId.
        /// </summary>
        [Required]
        [JsonProperty("teamAadObjectId")]
        public string TeamAADObjectId { get; set; }

        /// <summary>
        /// Gets or Sets TeamName.
        /// </summary>
        [Required]
        public string TeamName { get; set; }

        /// <summary>
        /// Gets or Sets ChannelId.
        /// </summary>
        [Required]
        public string ChannelId { get; set; }

        /// <summary>
        /// Gets or Sets ChannelName.
        /// </summary>
        [Required]
        public string ChannelName { get; set; }
    }
}

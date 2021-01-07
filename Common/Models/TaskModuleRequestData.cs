// <copyright file="TaskModuleRequestData.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Defines model for opening task module.
    /// </summary>
    public class TaskModuleRequestData
    {
        /// <summary>
        /// Gets or sets action type for button.
        /// </summary>
        [JsonProperty("type")]
        public string Type { get; set; }

        /// <summary>
        /// Gets or sets bot command to be used by bot for processing user inputs.
        /// </summary>
        [JsonProperty("command")]
        public string Command { get; set; }

        /// <summary>
        /// Gets or sets unique GUID related to activity Id.
        /// </summary>
        [JsonProperty("contextId")]
        public string ContextId { get; set; }

        /// <summary>
        /// Gets or sets the status change choice toggle.
        /// </summary>
        [JsonProperty("statusChangeChoice")]
        public string StatusChangeChoice { get; set; }

        /// <summary>
        /// Gets or sets the agent's notes.
        /// </summary>
        [JsonProperty("notes")]
        public string Notes { get; set; }
    }
}
// <copyright file="ContainerNames.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Repositories
{
    /// <summary>
    /// Request data Container names.
    /// </summary>
    public static class ContainerNames
    {
        /// <summary>
        /// Container name for the Agents.
        /// </summary>
        public static readonly string AgentContainerName = "Agents";

        /// <summary>
        /// Agents partition key name.
        /// </summary>
        public static readonly string AgentDataPartition = "/aadObjectId";

        /// <summary>
        /// Container name for the Request.
        /// </summary>
        public static readonly string RequestContainerName = "ConsultRequests";

        /// <summary>
        /// Request partition key name.
        /// </summary>
        public static readonly string RequestDataPartition = "/id";

        /// <summary>
        /// Container name for the Channel Mapping.
        /// </summary>
        public static readonly string ChannelMappingContainerName = "ChannelMappings";

        /// <summary>
        /// ChannelMapping partition key name.
        /// </summary>
        public static readonly string ChannelMappingDataPartition = "/id";

        /// <summary>
        /// Container name for the Channels.
        /// </summary>
        public static readonly string ChannelContainerName = "Channels";

        /// <summary>
        /// Channel partition key name.
        /// </summary>
        public static readonly string ChannelDataPartition = "/channelId";
    }
}

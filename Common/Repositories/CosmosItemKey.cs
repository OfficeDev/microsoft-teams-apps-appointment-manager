// <copyright file="CosmosItemKey.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Repositories
{
    /// <summary>
    /// Model representing the key for an item in Cosmos.
    /// </summary>
    public readonly struct CosmosItemKey
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CosmosItemKey"/> struct.
        /// </summary>
        /// <param name="id">The ID of the Cosmos item.</param>
        /// <param name="partitionKey">The partition key of the Cosmos item.</param>
        public CosmosItemKey(string id, string partitionKey)
        {
            this.Id = id;
            this.PartitionKey = partitionKey;
        }

        /// <summary>
        /// Gets the ID of the Cosmos item.
        /// </summary>
        public string Id { get; }

        /// <summary>
        /// Gets the partition key of the Cosmos item.
        /// </summary>
        public string PartitionKey { get; }
    }
}

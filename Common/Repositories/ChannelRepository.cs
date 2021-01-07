// <copyright file="ChannelRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Repositories
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration;

    /// <summary>
    /// An channel repository backed by Cosmos DB.
    /// </summary>
    public class ChannelRepository : CosmosRepository<Channel>, IChannelRepository<CosmosItemKey>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ChannelRepository"/> class.
        /// </summary>
        /// <param name="cosmosClient">The CosmosClient instance to use.</param>
        /// <param name="logger">The logger instance to use.</param>
        /// <param name="cosmosDBOptions">Cosmos DB configuration options.</param>
        public ChannelRepository(CosmosClient cosmosClient, ILogger<ChannelRepository> logger, IOptions<CosmosDBSettings> cosmosDBOptions)
            : base(
                cosmosClient,
                logger,
                ContainerNames.ChannelContainerName,
                cosmosDBOptions)
        {
        }

        /// <inheritdoc/>
        protected override string ContainerPartitionKey => ContainerNames.ChannelDataPartition;

        /// <inheritdoc/>
        public async Task<Channel> GetByChannelIdAsync(string channelId)
        {
            _ = channelId ?? throw new ArgumentNullException(nameof(channelId));

            var query = new QueryDefinition("select * from c where c.channelId = @channelId")
                .WithParameter("@channelId", channelId);
            var channelsResult = await this.QueryAsync(query);

            return channelsResult.FirstOrDefault();
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<Channel>> GetByTeamIdAsync(string teamId)
        {
            _ = teamId ?? throw new ArgumentNullException(nameof(teamId));

            var query = new QueryDefinition("select * from c where c.teamId = @teamId")
                .WithParameter("@teamId", teamId);
            var channelsResult = await this.QueryAsync(query);

            return channelsResult;
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<Channel>> GetAllAsync()
        {
            var channelsResult = await this.QueryAsync(new QueryDefinition("select * from c"));

            return channelsResult;
        }

        /// <inheritdoc/>
        protected override string ResolvePartitionKey(Channel channel) => channel.ChannelId;
    }
}

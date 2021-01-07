// <copyright file="ChannelMappingRepository.cs" company="Microsoft">
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
    /// A channel mapping repository backed by Cosmos DB.
    /// </summary>
    public class ChannelMappingRepository : CosmosRepository<ChannelMapping>, IChannelMappingRepository<CosmosItemKey>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ChannelMappingRepository"/> class.
        /// </summary>
        /// <param name="cosmosClient">The CosmosClient instance to use.</param>
        /// <param name="logger">The logger instance to use.</param>
        /// <param name="cosmosDBOptions">Cosmos DB configuration options.</param>
        public ChannelMappingRepository(CosmosClient cosmosClient, ILogger<ChannelMappingRepository> logger, IOptions<CosmosDBSettings> cosmosDBOptions)
            : base(
                cosmosClient,
                logger,
                ContainerNames.ChannelMappingContainerName,
                cosmosDBOptions)
        {
        }

        /// <inheritdoc/>
        protected override string ContainerPartitionKey => ContainerNames.ChannelMappingDataPartition;

        /// <inheritdoc/>
        public async Task<ChannelMapping> GetByCategoryAsync(string category)
        {
            _ = category ?? throw new ArgumentNullException(nameof(category));

            var query = new QueryDefinition("select * from c where c.category = @category")
                .WithParameter("@category", category);
            var channelMappingsResult = await this.QueryAsync(query);

            return channelMappingsResult.FirstOrDefault();
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<ChannelMapping>> GetByChannelIds(IEnumerable<string> channelIds)
        {
            _ = channelIds ?? throw new ArgumentNullException(nameof(channelIds));

            if (!channelIds.Any())
            {
                return Enumerable.Empty<ChannelMapping>();
            }

            var channelMappingsResult = await this.QueryAsync(
                query => query.Where(m => channelIds.Contains(m.ChannelId)));

            return channelMappingsResult;
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<ChannelMapping>> GetAll()
        {
            var channelsResult = await this.QueryAsync(new QueryDefinition("select * from c"));

            return channelsResult;
        }

        /// <inheritdoc/>
        protected override string ResolvePartitionKey(ChannelMapping entity) => entity.Id.ToString();
    }
}

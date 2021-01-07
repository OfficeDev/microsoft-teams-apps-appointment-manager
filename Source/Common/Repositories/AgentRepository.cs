// <copyright file="AgentRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Repositories
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration;

    /// <summary>
    /// An agent repository backed by Cosmos DB.
    /// </summary>
    public class AgentRepository : CosmosRepository<Agent>, IAgentRepository<CosmosItemKey>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AgentRepository"/> class.
        /// </summary>
        /// <param name="cosmosClient">The CosmosClient instance to use.</param>
        /// <param name="logger">The logger instance to use.</param>
        /// <param name="cosmosDBOptions">Cosmos DB configuration options.</param>
        public AgentRepository(CosmosClient cosmosClient, ILogger<AgentRepository> logger, IOptions<CosmosDBSettings> cosmosDBOptions)
            : base(
                cosmosClient,
                logger,
                ContainerNames.AgentContainerName,
                cosmosDBOptions)
        {
        }

        /// <inheritdoc/>
        protected override string ContainerPartitionKey => ContainerNames.AgentDataPartition;

        /// <inheritdoc/>
        public async Task<Agent> GetByObjectIdAsync(string objectId)
        {
            _ = objectId ?? throw new ArgumentNullException(nameof(objectId));

            var query = new QueryDefinition("select * from a where a.aadObjectId = @aadObjectId")
                .WithParameter("@aadObjectId", objectId);
            var agentsResult = await this.QueryAsync(query);

            return agentsResult.FirstOrDefault();
        }

        /// <inheritdoc/>
        protected override string ResolvePartitionKey(Agent agent) => agent.AADObjectId;
    }
}

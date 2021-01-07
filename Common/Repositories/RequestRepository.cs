// <copyright file="RequestRepository.cs" company="Microsoft">
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
    /// An request repository backed by Cosmos DB.
    /// </summary>
    public class RequestRepository : CosmosRepository<Request>, IRequestRepository<CosmosItemKey>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RequestRepository"/> class.
        /// </summary>
        /// <param name="cosmosClient">The CosmosClient instance to use.</param>
        /// <param name="logger">The logger instance to use.</param>
        /// <param name="cosmosDBOptions">Cosmos DB configuration options.</param>
        public RequestRepository(CosmosClient cosmosClient, ILogger<RequestRepository> logger, IOptions<CosmosDBSettings> cosmosDBOptions)
            : base(
                cosmosClient,
                logger,
                ContainerNames.RequestContainerName,
                cosmosDBOptions)
        {
        }

        /// <inheritdoc/>
        protected override string ContainerPartitionKey => ContainerNames.RequestDataPartition;

        /// <inheritdoc/>
        public async Task<IEnumerable<Request>> GetByAssignedToId(string assignedToId)
        {
            _ = assignedToId ?? throw new ArgumentNullException(nameof(assignedToId));

            var query = new QueryDefinition("select * from r where r.assignedToId = @assignedToId")
                .WithParameter("@assignedToId", assignedToId);
            var requestsResult = await this.QueryAsync(query);

            return requestsResult;
        }

        /// <inheritdoc/>
        public async Task<Request> GetByConversationId(string conversationId)
        {
            _ = conversationId ?? throw new ArgumentNullException(nameof(conversationId));

            var query = new QueryDefinition("select * from r where contains (r.joinUri, @conversationId)")
                .WithParameter("@conversationId", conversationId);
            var requestsResult = await this.QueryAsync(query);

            return requestsResult.FirstOrDefault();
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<Request>> GetFiltered(IEnumerable<string> categories, IEnumerable<RequestStatus> statuses)
        {
            var requestsResult = await this.QueryAsync(query =>
            {
                var ret = query;

                if (categories != null && categories.Any())
                {
                    ret = ret.Where(r => categories.Contains(r.Category));
                }

                if (statuses != null && statuses.Any())
                {
                    ret = ret.Where(r => statuses.Contains(r.Status));
                }

                return ret;
            });

            return requestsResult;
        }

        /// <inheritdoc/>
        protected override string ResolvePartitionKey(Request request) => request.Id.ToString();
    }
}

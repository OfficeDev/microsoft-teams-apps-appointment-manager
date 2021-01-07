// <copyright file="CosmosRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Repositories
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos;
    using Microsoft.Azure.Cosmos.Linq;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration;

    /// <summary>
    /// Base for a repository backed by Cosmos DB.
    /// </summary>
    /// <typeparam name="T">The type of item stored in the repository.</typeparam>
    public abstract class CosmosRepository<T> : IRepository<T, CosmosItemKey>
        where T : BaseModel
    {
        private readonly Lazy<Task> initializeTask;

        private readonly IOptions<CosmosDBSettings> cosmosDBOptions;

        private Container cosmosContainer;

        /// <summary>
        /// Initializes a new instance of the <see cref="CosmosRepository{T}"/> class.
        /// </summary>
        /// <param name="cosmosClient">The CosmosClient instance to use.</param>
        /// <param name="logger">The logger instance to use.</param>
        /// <param name="containerName">The name of the Cosmos DB container.</param>
        /// <param name="cosmosDBOptions">Cosmos DB configuration options.</param>
        public CosmosRepository(
            CosmosClient cosmosClient,
            ILogger logger,
            string containerName,
            IOptions<CosmosDBSettings> cosmosDBOptions)
        {
            this.CosmosClient = cosmosClient;
            this.Logger = logger;
            this.ContainerName = containerName;
            this.cosmosDBOptions = cosmosDBOptions;

            this.initializeTask = new Lazy<Task>(() => this.InitializeDatabaseAsync());
        }

        /// <summary>
        /// Gets the logger.
        /// </summary>
        protected ILogger Logger { get; }

        /// <summary>
        /// Gets the Cosmos client.
        /// </summary>
        protected CosmosClient CosmosClient { get; }

        /// <summary>
        /// Gets the name of the Cosmos container.
        /// </summary>
        protected string ContainerName { get; }

        /// <summary>
        /// Gets the container's partition key.
        /// </summary>
        protected abstract string ContainerPartitionKey { get; }

        /// <summary>
        /// Gets an item from Cosmos DB that matches the given key.
        /// </summary>
        /// <inheritdoc/>
        public async Task<T> GetAsync(CosmosItemKey key)
        {
            await this.EnsureInitializedAsync();

            try
            {
                var response = await this.cosmosContainer.ReadItemAsync<T>(key.Id, new PartitionKey(key.PartitionKey));
                return response.Resource;
            }
            catch (CosmosException ex) when (ex.StatusCode == System.Net.HttpStatusCode.NotFound)
            {
                this.Logger.LogError(ex, "Failed to get item in Cosmos DB");
                return default;
            }
        }

        /// <summary>
        /// Adds the given item to Cosmos DB.
        /// </summary>
        /// <inheritdoc/>
        public async Task AddAsync(T entity)
        {
            await this.EnsureInitializedAsync();

            string partitionKey = this.ResolvePartitionKey(entity);
            await this.cosmosContainer.CreateItemAsync(entity, new PartitionKey(partitionKey));
        }

        /// <summary>
        /// Updates the given item in Cosmos DB, or creates it if it doesn't exist.
        /// </summary>
        /// <inheritdoc/>
        public async Task UpsertAsync(T entity)
        {
            await this.EnsureInitializedAsync();

            string partitionKey = this.ResolvePartitionKey(entity);
            try
            {
                await this.cosmosContainer.UpsertItemAsync<T>(entity, new PartitionKey(partitionKey));
            }
            catch (CosmosException ex) when (ex.StatusCode == System.Net.HttpStatusCode.NotFound)
            {
                this.Logger.LogError(ex, "Failed to upsert item in Cosmos DB");
            }
        }

        /// <summary>
        /// Deletes the item in Cosmos DB that matches the given key.
        /// </summary>
        /// <inheritdoc/>
        public async Task DeleteAsync(CosmosItemKey key)
        {
            await this.EnsureInitializedAsync();

            try
            {
                await this.cosmosContainer.DeleteItemAsync<T>(key.Id, new PartitionKey(key.PartitionKey));
            }
            catch (CosmosException ex) when (ex.StatusCode == System.Net.HttpStatusCode.NotFound)
            {
                this.Logger.LogError(ex, "Failed to delete item in Cosmos DB");
            }
        }

        /// <summary>
        /// Query for items in Cosmos DB by providing a query definition.
        /// </summary>
        /// <param name="query">The definition of the query to execute.</param>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the collection of items that match the query.</returns>
        protected async Task<IEnumerable<T>> QueryAsync(QueryDefinition query)
        {
            await this.EnsureInitializedAsync();

            List<T> results = new List<T>();

            using (var feedIterator = this.cosmosContainer.GetItemQueryIterator<T>(query))
            {
                while (feedIterator.HasMoreResults)
                {
                    FeedResponse<T> feedResponse = await feedIterator.ReadNextAsync();
                    results.AddRange(feedResponse);
                }
            }

            return results;
        }

        /// <summary>
        /// Query for items in Cosmos DB by defining IQueryable operations.
        /// </summary>
        /// <param name="filter">The IQueryable operations to apply.</param>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the collection of items that match the query.</returns>
        protected async Task<IEnumerable<T>> QueryAsync(Func<IQueryable<T>, IQueryable<T>> filter)
        {
            await this.EnsureInitializedAsync();

            IQueryable<T> query = this.cosmosContainer.GetItemLinqQueryable<T>();
            query = filter(query);

            return await this.QueryAsync(query.ToQueryDefinition());
        }

        /// <summary>
        /// Resolves the partition key for an entity.
        /// </summary>
        /// <param name="entity">The entity for which to get the partition key.</param>
        /// <returns>The partition key for the entity.</returns>
        protected abstract string ResolvePartitionKey(T entity);

        /// <summary>
        /// Initializes Cosmos DB by creating the necessary database and container if needed.
        /// </summary>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        private async Task InitializeDatabaseAsync()
        {
            var databaseName = this.cosmosDBOptions.Value.DatabaseName;

            // Create the database and container if they don't exist
            var databaseResponse = await this.CosmosClient.CreateDatabaseIfNotExistsAsync(databaseName);
            var containerResponse = await databaseResponse.Database.CreateContainerIfNotExistsAsync(this.ContainerName, this.ContainerPartitionKey, 400);

            this.cosmosContainer = containerResponse.Container;
        }

        private async Task EnsureInitializedAsync()
        {
            await this.initializeTask.Value;
        }
    }
}

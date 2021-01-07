// <copyright file="IAgentRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Repositories
{
    using System.Threading.Tasks;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;

    /// <summary>
    /// The base interface for an agent repository.
    /// </summary>
    /// <typeparam name="TKey">The type of the key that uniquely identifies an agent.</typeparam>
    public interface IAgentRepository<TKey> : IRepository<Agent, TKey>
    {
        /// <summary>
        /// Gets an agent by their AAD object ID.
        /// </summary>
        /// <param name="objectId">The object ID of the agent to get.</param>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the agent with the given AAD object ID.</returns>
        Task<Agent> GetByObjectIdAsync(string objectId);
    }
}

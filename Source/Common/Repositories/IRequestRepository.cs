// <copyright file="IRequestRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Repositories
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;

    /// <summary>
    /// The base interface for the request repository.
    /// </summary>
    /// <typeparam name="TKey">The type of the key that uniquely identifies a request.</typeparam>
    public interface IRequestRepository<TKey> : IRepository<Request, TKey>
    {
        /// <summary>
        /// Gets all requests assigned to the given user.
        /// </summary>
        /// <param name="assignedToId">The ID of the user.</param>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the collection of requests assigned to the given user.</returns>
        Task<IEnumerable<Request>> GetByAssignedToId(string assignedToId);

        /// <summary>
        /// Gets the request associated with the given conversation ID.
        /// </summary>
        /// <param name="conversationId">The conversation ID of the request to get.</param>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the request matching the given conversation ID.</returns>
        Task<Request> GetByConversationId(string conversationId);

        /// <summary>
        /// Gets all requests that match the given category and status filters.
        /// </summary>
        /// <param name="categories">The categories to include.</param>
        /// <param name="statuses">The statuses to include.</param>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the collection of requests satisfying the given filters.</returns>
        Task<IEnumerable<Request>> GetFiltered(IEnumerable<string> categories, IEnumerable<RequestStatus> statuses);
    }
}

// <copyright file="IChannelRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Repositories
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;

    /// <summary>
    /// The base interface for a channel repository.
    /// </summary>
    /// <typeparam name="TKey">The type of the key that uniquely identifies a channel.</typeparam>
    public interface IChannelRepository<TKey> : IRepository<Channel, TKey>
    {
        /// <summary>
        /// Gets a channel by its Teams channel ID.
        /// </summary>
        /// <param name="channelId">The Teams channel ID of the channel to get.</param>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the channel with the given Teams channel ID.</returns>
        Task<Channel> GetByChannelIdAsync(string channelId);

        /// <summary>
        /// Gets all channels for the given team ID.
        /// </summary>
        /// <param name="teamId">The team ID of the channels to get.</param>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the collection of channels for the given team ID.</returns>
        Task<IEnumerable<Channel>> GetByTeamIdAsync(string teamId);

        /// <summary>
        /// Gets all channels.
        /// </summary>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the collection of all channels.</returns>
        Task<IEnumerable<Channel>> GetAllAsync();
    }
}

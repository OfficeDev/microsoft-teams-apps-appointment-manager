// <copyright file="IChannelMappingRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Repositories
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;

    /// <summary>
    /// The base interface for a channel mapping repository.
    /// </summary>
    /// <typeparam name="TKey">The type of the key that uniquely identifies a channel mapping.</typeparam>
    public interface IChannelMappingRepository<TKey> : IRepository<ChannelMapping, TKey>
    {
        /// <summary>
        /// Gets the channel mapping for the given category.
        /// </summary>
        /// <param name="category">The category of the channel mapping to get.</param>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the channel mapping for the given category.</returns>
        Task<ChannelMapping> GetByCategoryAsync(string category);

        /// <summary>
        /// Gets all channel mappings for the given channel IDs.
        /// </summary>
        /// <param name="channelIds">The channel IDs of the channel mappings to get.</param>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the collection of channel mappings matching the given channel IDs.</returns>
        Task<IEnumerable<ChannelMapping>> GetByChannelIds(IEnumerable<string> channelIds);

        /// <summary>
        /// Gets all channel mappings.
        /// </summary>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the collection of all channel mappings.</returns>
        Task<IEnumerable<ChannelMapping>> GetAll();
    }
}

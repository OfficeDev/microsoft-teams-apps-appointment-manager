// <copyright file="IRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Repositories
{
    using System.Threading.Tasks;

    /// <summary>
    /// The base interface for a repository.
    /// </summary>
    /// <typeparam name="T">The type of item stored in the repository.</typeparam>
    /// <typeparam name="TKey">The type of the key that uniquely identifies an item.</typeparam>
    public interface IRepository<T, TKey>
    {
        /// <summary>
        /// Gets an item that matches the given key.
        /// </summary>
        /// <param name="key">The key that uniquely identifies the item.</param>
        /// <returns>A <see cref="Task{TResult}"/> whose result is the item that matches the given key.</returns>
        Task<T> GetAsync(TKey key);

        /// <summary>
        /// Adds the given item.
        /// </summary>
        /// <param name="entity">The entity to add.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        Task AddAsync(T entity);

        /// <summary>
        /// Updates the given item, or creates it if it doesn't exist.
        /// </summary>
        /// <param name="entity">The entity to upsert.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        Task UpsertAsync(T entity);

        /// <summary>
        /// Deletes the item that matches the given key.
        /// </summary>
        /// <param name="key">The key that uniquely identifies the item.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        Task DeleteAsync(TKey key);
    }
}

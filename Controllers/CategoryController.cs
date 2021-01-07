// <copyright file="CategoryController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Controllers
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.App.VirtualConsult.Common.Repositories;

    /// <summary>
    /// Web API for getting consumer categories.
    /// </summary>
    [ApiController]
    public class CategoryController : ControllerBase
    {
        private readonly IChannelMappingRepository<CosmosItemKey> channelMappingRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="CategoryController"/> class.
        /// </summary>
        /// <param name="channelMappingRepository">The channel mapping repository.</param>
        public CategoryController(IChannelMappingRepository<CosmosItemKey> channelMappingRepository)
        {
            this.channelMappingRepository = channelMappingRepository;
        }

        /// <summary>
        /// Gets a list of consumer categories.
        /// </summary>
        /// <response code="200">Successfully retrieved the list of consumer categories.</response>
        /// <response code="400">Error occured while fetching the list.</response>
        /// <returns>Enumerable list of string categories.</returns>
        [HttpGet]
        [Route("/api/category")]
        public async Task<IActionResult> GetAsync()
        {
            // Get all mappings in database
            var mappings = await this.channelMappingRepository.GetAll();
            var categories = mappings.Select(i => i.Category).OrderBy(i => i);

            return this.Ok(categories);
        }
    }
}
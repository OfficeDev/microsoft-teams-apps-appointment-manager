// <copyright file="ChannelController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Controllers
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;
    using Microsoft.Teams.App.VirtualConsult.Common.Repositories;

    /// <summary>
    /// Secure Web API controller for admins to manage channel mappings.
    /// </summary>
    [Authorize]
    [ApiController]
    public class ChannelController : ControllerBase
    {
        private readonly IChannelRepository<CosmosItemKey> channelRepository;
        private readonly IChannelMappingRepository<CosmosItemKey> channelMappingRepository;
        private readonly ILogger<ChannelController> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="ChannelController"/> class.
        /// </summary>
        /// <param name="channelRepository">The channel repository.</param>
        /// <param name="channelMappingRepository">The channel mapping repository.</param>
        /// <param name="logger">The logger.</param>
        public ChannelController(IChannelRepository<CosmosItemKey> channelRepository, IChannelMappingRepository<CosmosItemKey> channelMappingRepository, ILogger<ChannelController> logger)
        {
            this.channelRepository = channelRepository;
            this.channelMappingRepository = channelMappingRepository;
            this.logger = logger;
        }

        /// <summary>
        /// Gets collection of channel mappings.
        /// </summary>
        /// <response code="200">A list of channel mappings.</response>
        /// <returns>Enumerable list of ChannelMapping objects</returns>
        [HttpGet]
        [Route("/api/channel/channelmappings")]
        public async Task<IActionResult> GetChannelMappingsAsync()
        {
            // Get all mappings in database
            var mappings = await this.channelMappingRepository.GetAll();

            return this.Ok(mappings);
        }

        /// <summary>
        /// Updates a channel mapping.
        /// </summary>
        /// <param name="id">The ID of the channel mapping.</param>
        /// <param name="mapping">The channel mapping to update.</param>
        /// <response code="200">Successful update of the mapping.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <returns>Task.</returns>
        [HttpPatch]
        [Route("/api/channel/channelmappings/{id}")]
        public async Task<IActionResult> UpdateChannelMappingAsync(string id, [FromBody] ChannelMapping mapping)
        {
            // Get the channel mapping from database
            var originalMapping = await this.channelMappingRepository.GetAsync(new CosmosItemKey(id, id));
            if (originalMapping == null)
            {
                this.logger.LogError("Failed to update channel mapping for mapping {ChannelMappingId}. The mapping does not exist.", id);
                return this.NotFound(new UnsuccessfulResponse { Reason = "Channel mapping not found" });
            }

            // Update channel mapping properties that are allowed to be updated
            originalMapping.Category = mapping.Category;
            originalMapping.ChannelId = mapping.ChannelId;
            originalMapping.Supervisors = mapping.Supervisors;
            originalMapping.BookingsBusiness = mapping.BookingsBusiness;
            originalMapping.BookingsService = mapping.BookingsService;
            await this.channelMappingRepository.UpsertAsync(originalMapping);

            return this.NoContent();
        }

        /// <summary>
        /// Deletes a channel mapping.
        /// </summary>
        /// <remarks>Deletes a channel mapping using mapping id.</remarks>
        /// <param name="id">ID of the channel mapping to delete.</param>
        /// <response code="200">Channel mapping was successfully removed.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <returns>Task.</returns>
        [HttpDelete]
        [Route("/api/channel/channelmappings/{id}")]
        public async Task<IActionResult> DeleteChannelMappingAsync(string id)
        {
            // Delete the mapping
            await this.channelMappingRepository.DeleteAsync(new CosmosItemKey(id, id));

            return this.NoContent();
        }

        /// <summary>
        /// Creates a new channel mapping.
        /// </summary>
        /// <param name="mapping">The channel mapping to add.</param>
        /// <response code="200">Successful creation of the mapping.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <returns>ChannelMapping with new id</returns>
        [HttpPost]
        [Route("/api/channel/channelmappings")]
        public async virtual Task<IActionResult> CreateChannelMappingAsync([FromBody] ChannelMapping mapping)
        {
            // get the channelMapping for the request category
            var mappingDetails = await this.channelMappingRepository.GetByCategoryAsync(mapping.Category);
            if (mappingDetails != null)
            {
                return this.Conflict(new UnsuccessfulResponse { Reason = "Category already in use" });
            }

            mapping.Id = Guid.NewGuid();
            mapping.CreatedDateTime = DateTime.Now;
            await this.channelMappingRepository.AddAsync(mapping);
            return this.Ok(mapping);
        }

        /// <summary>
        /// Returns collection of channels the bot is added to.
        /// </summary>
        /// <remarks>Gets a list of Teams channels that the consult bot is added to.</remarks>
        /// <response code="200">A list of Teams channels that the consult bot is added to.</response>
        /// <response code="500">Unable to get the channel data.</response>
        /// <returns>Enumerable list of Channel objects.</returns>
        [HttpGet]
        [Route("/api/channel")]
        public async Task<IActionResult> GetChannelsAsync()
        {
            // Get all channels in database
            var items = await this.channelRepository.GetAllAsync();

            return this.Ok(items);
        }

        /// <summary>
        /// Returns collection of channel mappings for a specific team.
        /// </summary>
        /// <param name="id">The ID of the team.</param>
        /// <remarks>Gets a list of channel mappings for a specific team.</remarks>
        /// <response code="200">A list of channel mappings for a specific team.</response>
        /// <response code="500">Unable to get the channel mapping data.</response>
        /// <returns>Enumerable list of ChannelMapping objects.</returns>
        [HttpGet]
        [Route("/api/channel/channelmappings/{id}")]
        public async Task<IActionResult> GetChannelMappingsForTeamAsync(string id)
        {
            // First get all channels for a specific Team
            var channels = await this.channelRepository.GetByTeamIdAsync(id);

            // Next get the channel mappings for those channels within the team
            var channelIds = channels.Select(c => c.ChannelId);
            var channelMappings = await this.channelMappingRepository.GetByChannelIds(channelIds);

            return this.Ok(channelMappings);
        }
    }
}
// <copyright file="AgentController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Controllers
{
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration;
    using Microsoft.Teams.App.VirtualConsult.Common.Repositories;

    /// <summary>
    /// /// Web API for working with agents/users
    /// </summary>
    [Authorize]
    [ApiController]
    public class AgentController : ControllerBase
    {
        private readonly IAgentRepository<CosmosItemKey> agentRepository;
        private readonly ILogger<AgentController> logger;
        private readonly IOptions<AzureADSettings> azureADOptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="AgentController"/> class.
        /// </summary>
        /// <param name="agentRepository">The agent repository.</param>
        /// <param name="logger">The logger.</param>
        /// <param name="azureADOptions">Azure AD configuration options.</param>
        public AgentController(IAgentRepository<CosmosItemKey> agentRepository, ILogger<AgentController> logger, IOptions<AzureADSettings> azureADOptions)
        {
            this.agentRepository = agentRepository;
            this.logger = logger;
            this.azureADOptions = azureADOptions;
        }

        /// <summary>
        /// Get the agents in a Teams channel using Graph APIs.
        /// </summary>
        /// <param name="teamAadObjectId">The channel id where the agents are located.</param>
        /// <remarks>Gets the agents specified in the Teams team.</remarks>
        /// <response code="200">A list of Teams agents.</response>
        /// <response code="404">Unable to get the agents.</response>
        /// <returns>IActionResult.</returns>
        [HttpGet]
        [Route("/api/agents")]
        public async Task<IActionResult> GetAsync([FromHeader] string teamAadObjectId)
        {
            var userObjectId = this.GetUserObjectId();

            var graphTimeSlotsResponse = await Common.Utils.GraphUtil.GetMembersInTeamsChannelAsync(teamAadObjectId, this.azureADOptions.Value);

            if (!graphTimeSlotsResponse.FailureReason.Equals(string.Empty))
            {
                this.logger.LogError($"Failed to get members in Team {{TeamObjectId}}. The Graph call failed: {graphTimeSlotsResponse.FailureReason}", teamAadObjectId);
                return this.NotFound(new UnsuccessfulResponse { Reason = graphTimeSlotsResponse.FailureReason });
            }

            // Checks if requestor is in the list of agents in Teams.
            if (!graphTimeSlotsResponse.Result.Any(user => user.Id == userObjectId))
            {
                this.logger.LogWarning("Failed to get members in Team {TeamObjectId}. User {UserObjectId} is not a member of the Team.", teamAadObjectId, userObjectId);
                return this.Unauthorized();
            }

            return this.Ok(graphTimeSlotsResponse.Result);
        }

        /// <summary>
        /// Updates the agent.
        /// </summary>
        /// <param name="agent">Agent properties to update.</param>
        /// <param name="agentId">The object ID of the agent to update.</param>
        /// <response code="204">Successful update of agent.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <response code="403">Forbidden from updating the agent.</response>
        /// <response code="404">Agent not found.</response>
        /// <returns>An IActionResult indicating the result of the API call.</returns>
        [HttpPatch]
        [Route("/api/agent/{agentId}")]
        public async Task<IActionResult> PatchAsync([FromBody] Agent agent, string agentId)
        {
            // Ensure that calling agent matches the agent to update
            var userObjectId = this.GetUserObjectId();
            var userName = this.GetUserName();
            if (userObjectId != agentId)
            {
                this.logger.LogWarning("Failed to update agent data for user {AgentObjectId}. User {UserObjectId} is not allowed to update that agent.", agentId, userObjectId);
                return this.StatusCode(403, new UnsuccessfulResponse { Reason = "Insufficient permissions to update agent" });
            }

            // Get existing agent from DB
            var existingAgent = await this.agentRepository.GetByObjectIdAsync(agentId);
            if (existingAgent == null)
            {
                this.logger.LogError("Failed to update agent data for user {AgentObjectId}. The agent does not exist.", agentId);
                return this.NotFound(new UnsuccessfulResponse { Reason = "Agent not found" });
            }

            // Update agent properties that are allowed to be updated
            existingAgent.Name = userName;
            existingAgent.Locale = agent.Locale;

            // Update agent in DB
            await this.agentRepository.UpsertAsync(existingAgent);

            return this.NoContent();
        }

        /// <summary>
        /// Gets the calling user's object ID from the identity claims.
        /// </summary>
        /// <returns>The calling user's object ID.</returns>
        private string GetUserObjectId()
        {
            return this.User.Claims.FirstOrDefault(i => i.Type == "http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value;
        }

        /// <summary>
        /// Gets the calling user's name from the identity claims.
        /// </summary>
        /// <returns>The calling user's name.</returns>
        private string GetUserName()
        {
            var givenName = this.User.Claims.FirstOrDefault(i => i.Type == ClaimTypes.GivenName)?.Value;
            var surname = this.User.Claims.FirstOrDefault(i => i.Type == ClaimTypes.Surname)?.Value;

            return $"{givenName} {surname}";
        }
    }
}

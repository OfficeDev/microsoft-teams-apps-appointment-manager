// <copyright file="RequestController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Controllers
{
    extern alias GraphBetaLib;

    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Requests;
    using Microsoft.Teams.App.VirtualConsult.Common.Repositories;
    using Microsoft.Teams.App.VirtualConsult.Common.Utils;
    using GraphBeta = GraphBetaLib::Microsoft.Graph;

    /// <summary>
    /// Web API controller for consult requests.
    /// </summary>
    [Authorize]
    [ApiController]
    public class RequestController : ControllerBase
    {
        private readonly IRequestRepository<CosmosItemKey> requestRepository;
        private readonly IAgentRepository<CosmosItemKey> agentRepository;
        private readonly IChannelRepository<CosmosItemKey> channelRepository;
        private readonly IChannelMappingRepository<CosmosItemKey> channelMappingRepository;
        private readonly ILogger<RequestController> logger;
        private readonly IStringLocalizer<SharedResources> localizer;
        private readonly IOptions<AzureADSettings> azureADOptions;
        private readonly IOptions<BotSettings> botOptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="RequestController"/> class.
        /// </summary>
        /// <param name="requestRepository">The request repository.</param>
        /// <param name="agentRepository">The agent repository.</param>
        /// <param name="channelRepository">The channel repository.</param>
        /// <param name="channelMappingRepository">The channel mapping repository.</param>
        /// <param name="logger">The logger.</param>
        /// <param name="localizer">The localizer to use for strings.</param>
        /// <param name="azureADOptions">Azure AD configuration options.</param>
        /// <param name="botOptions">Bot configuration options.</param>
        public RequestController(
            IRequestRepository<CosmosItemKey> requestRepository,
            IAgentRepository<CosmosItemKey> agentRepository,
            IChannelRepository<CosmosItemKey> channelRepository,
            IChannelMappingRepository<CosmosItemKey> channelMappingRepository,
            ILogger<RequestController> logger,
            IStringLocalizer<SharedResources> localizer,
            IOptions<AzureADSettings> azureADOptions,
            IOptions<BotSettings> botOptions)
        {
            this.requestRepository = requestRepository;
            this.agentRepository = agentRepository;
            this.channelRepository = channelRepository;
            this.channelMappingRepository = channelMappingRepository;
            this.logger = logger;
            this.localizer = localizer;
            this.azureADOptions = azureADOptions;
            this.botOptions = botOptions;
        }

        /// <summary>
        /// Add an attachement to a consult request.
        /// </summary>
        /// <remarks>Adds an attachment to a request identified using its id.</remarks>
        /// <param name="requestId">The consult id to add the attachment to.</param>
        /// <param name="body">The attachments to add to the consult request.</param>
        /// <response code="200">Attachment was successfully added to the consult request.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <response code="404">Consult request not found.</response>
        /// <returns>IActionResult.</returns>
        [HttpPost]
        [Route("/api/request/attachment/{requestId}")]
        public virtual async Task<IActionResult> RequestAttachmentRequestIdPostAsync(string requestId, [FromBody] Attachment body)
        {
            if (string.IsNullOrEmpty(requestId) || body == null)
            {
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Invalid consult ID provided" });
            }

            var request = await this.requestRepository.GetAsync(new CosmosItemKey(requestId, requestId));

            if (request == null)
            {
                this.logger.LogError("Failed to add attachment to consult request {ConsultRequestId}. The consult request does not exist.", requestId);
                return this.NotFound(new UnsuccessfulResponse { Reason = "Consult request not found" });
            }

            List<Attachment> attachments = request.Attachments;

            if (attachments == null)
            {
                attachments = new List<Attachment>();
            }

            body.Id = Guid.NewGuid();
            body.CreatedDateTime = DateTime.Now;
            body.CreatedById = new Guid(this.GetUserObjectId());
            body.CreatedByName = this.GetUserName();
            attachments.Add(body);
            request.Attachments = attachments;

            await this.requestRepository.UpsertAsync(request);

            return this.Ok();
        }

        /// <summary>
        /// Gets a list of consult requests that are assigned to the calling agent.
        /// </summary>
        /// <response code="200">A list of requests assigned to the calling agent.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <returns>List of ConsultRequest objects.</returns>
        [HttpGet]
        [Route("/api/request")]
        public async Task<IActionResult> GetRequestsAsync()
        {
            // Use caller's identity to filter by agent
            var userObjectId = this.GetUserObjectId();
            if (string.IsNullOrEmpty(userObjectId))
            {
                this.logger.LogError("Failed to get requests for agent. The user's object ID was null or empty.");
                return this.Unauthorized();
            }

            // Look up the agents requests
            var requests = await this.requestRepository.GetByAssignedToId(userObjectId);

            return this.Ok(requests);
        }

        /// <summary>
        /// Gets a specific consult request that matches the specified conversation ID.
        /// </summary>
        /// <param name="conversationId">The conversation ID.</param>
        /// <response code="200">A consult request matching the specified conversationID.</response>
        /// <response code="404">Consult request not found.</response>
        /// <returns>A ConsultRequest object.</returns>
        [HttpGet]
        [Route("/api/request/lookup/{conversationId}")]
        public async Task<IActionResult> GetConversationToRequestAsync(string conversationId)
        {
            if (string.IsNullOrEmpty(conversationId))
            {
                return this.NotFound(new UnsuccessfulResponse { Reason = "Invalid conversation ID provided" });
            }

            var request = await this.requestRepository.GetByConversationId(conversationId);

            if (request == null)
            {
                this.logger.LogError("Failed to find consult request that matches conversation ID {ConversationId}.", conversationId);
                return this.NotFound(new UnsuccessfulResponse { Reason = "Consult request not found" });
            }

            return this.Ok(request);
        }

        /// <summary>
        /// Gets a specific consult request by ID.
        /// </summary>
        /// <param name="consultId">ID of the consult request to get.</param>
        /// <response code="200">A consult request matching the specified ID.</response>
        /// <response code="404">Consult request not found.</response>
        /// <returns>A ConsultRequest object.</returns>
        [HttpGet]
        [Route("/api/request/{consultId}")]
        public async Task<IActionResult> GetRequestAsync(string consultId)
        {
            if (string.IsNullOrEmpty(consultId))
            {
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Invalid consult ID provided" });
            }

            var request = await this.requestRepository.GetAsync(new CosmosItemKey(consultId, consultId));

            if (request == null)
            {
                this.logger.LogError("Failed to get consult request {ConsultRequestId}. The consult request does not exist.", consultId);
                return this.NotFound(new UnsuccessfulResponse { Reason = "Consult request not found" });
            }

            return this.Ok(request);
        }

        /// <summary>
        /// Checks if the user is a supervisor for the specified request.
        /// </summary>
        /// <param name="consultId">Request ID for the request to check.</param>
        /// <response code="200">Boolean value indicating if the calling user is a supervisor for the request.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <returns>A boolean value.</returns>
        [HttpGet]
        [Route("/api/request/{consultId}/issupervisor")]
        public async Task<IActionResult> IsSupervisorAsync(string consultId)
        {
            if (string.IsNullOrEmpty(consultId))
            {
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Invalid consult ID provided" });
            }

            return this.Ok(await this.CheckSupervisorAsync(consultId));
        }

        /// <summary>
        /// Gets a list of consult requests filtered by channel(s) and status(es).
        /// </summary>
        /// <param name="filter">RequestFilter object containing arrays of channelIds and statuses.</param>
        /// <response code="200">A list of requests for the specified filters.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <returns>List of ConsultRequest objects.</returns>
        [HttpPost]
        [Route("/api/request/filtered")]
        public async Task<IActionResult> GetRequestsFilteredAsync([FromBody] RequestFilter filter)
        {
            // Check if any filters are empty
            if (filter.Categories.Count == 0 || filter.Statuses.Count == 0)
            {
                return this.Ok(Enumerable.Empty<Request>());
            }

            // query requests for the categories and statuses requested
            var requests = await this.requestRepository.GetFiltered(filter.Categories, filter.Statuses);

            return this.Ok(requests);
        }

        /// <summary>
        /// Get the channel Id of where the request is posted.
        /// </summary>
        /// <param name="consultId">ID of the consult request to get.</param>
        /// <remarks> Gets the channel Id where the consult request is posted.</remarks>
        /// <response code="200">The channel Id was successfully identified.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <returns>IActionResult.</returns>
        [HttpGet]
        [Route("/api/request/{consultId}/channel")]
        public async Task<IActionResult> GetRequestChannelIdAsync(string consultId)
        {
            if (string.IsNullOrEmpty(consultId))
            {
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Invalid consult ID provided" });
            }

            var request = await this.requestRepository.GetAsync(new CosmosItemKey(consultId, consultId));

            if (request == null)
            {
                this.logger.LogError("Failed to get channel for consult request {ConsultRequestId}. The consult request does not exist.", consultId);
                return this.NotFound(new UnsuccessfulResponse { Reason = "Consult request not found" });
            }

            var mapping = await this.channelMappingRepository.GetByCategoryAsync(request.Category);

            if (mapping == null)
            {
                this.logger.LogError("Failed to get channel for consult request {ConsultRequestId}. No mapping exists for consult category {ConsultCategory}.", consultId, request.Category);
                return this.NotFound(new UnsuccessfulResponse { Reason = "Category not found" });
            }

            // Get channel from database
            var channel = await this.channelRepository.GetByChannelIdAsync(mapping.ChannelId);

            if (channel == null)
            {
                this.logger.LogError("Failed to get channel for consult request {ConsultRequestId}. Channel {ChannelId} does not exist, but a mapping to that channel exists.", consultId, mapping.ChannelId);
                return this.NotFound(new UnsuccessfulResponse { Reason = "Channel not found" });
            }

            return this.Ok(channel);
        }

        /// <summary>
        /// Add a note to a consult request.
        /// </summary>
        /// <param name="consultId">ID of the consult request to add the note to.</param>
        /// <param name="body">The note to add to the consult request.</param>
        /// <response code="200">Note was successfully added to the consult request.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <response code="404">Consult request not found.</response>
        /// <returns>An IActionResult indicating the result of the API call.</returns>
        [HttpPost]
        [Route("/api/request/{consultId}/notes")]
        public virtual async Task<IActionResult> CreateRequestNoteAsync(string consultId, [FromBody] Note body)
        {
            // Extract user's object ID from claims
            var userObjectId = this.GetUserObjectId();
            var userName = this.GetUserName();
            if (string.IsNullOrWhiteSpace(userObjectId))
            {
                this.logger.LogError("Failed to add note to consult request. The user's object ID was null or empty.");
                return this.Unauthorized();
            }

            if (string.IsNullOrEmpty(consultId))
            {
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Invalid consult ID provided" });
            }

            var request = await this.requestRepository.GetAsync(new CosmosItemKey(consultId, consultId));
            if (request == null)
            {
                this.logger.LogError("Failed to add note to consult request {ConsultRequestId}. The consult request does not exist.", consultId);
                return this.NotFound(new UnsuccessfulResponse { Reason = "Consult request not found" });
            }

            var currentTime = DateTime.UtcNow;
            var userObjectIdGuid = new Guid(userObjectId);

            var newNote = new Note
            {
                Id = Guid.NewGuid(),
                CreatedByName = userName,
                CreatedById = userObjectIdGuid,
                CreatedDateTime = currentTime,
                Text = body.Text,
            };

            request.Notes = request.Notes ?? new List<Note>();
            request.Notes.Add(newNote);

            await this.requestRepository.UpsertAsync(request);

            return this.Ok(newNote);
        }

        /// <summary>
        /// Create a new consult request.
        /// </summary>
        /// <param name="request">Consult request from consumer.</param>
        /// <response code="200">Successful creation of consult request.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <returns>Task.</returns>
        [AllowAnonymous]
        [HttpPost]
        [Route("/api/request")]
        public async Task<IActionResult> CreateRequestAsync([FromBody] Request request)
        {
            // Look up the channel to notify
            var mapping = await this.channelMappingRepository.GetByCategoryAsync(request.Category);

            if (mapping == null)
            {
                this.logger.LogError("Failed to create consult request. No mapping exists for consult category {ConsultCategory}.", request.Category);
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Consult category is not valid" });
            }

            // Lookup the Channel
            var channel = await this.channelRepository.GetByChannelIdAsync(mapping.ChannelId);

            if (channel == null)
            {
                this.logger.LogError("Failed to create consult request. Channel {ChannelId} does not exist, but a mapping to that channel exists.", mapping.ChannelId);
                return this.BadRequest(new UnsuccessfulResponse { Reason = "New consult requests cannot be made for this category" });
            }

            // Save the new request in the database
            request.Id = Guid.NewGuid();
            request.FriendlyId = this.GetFriendlyId();
            request.Status = RequestStatus.Unassigned;
            request.BookingsBusinessId = mapping.BookingsBusiness.Id;
            request.BookingsServiceId = mapping.BookingsService.Id;
            request.CreatedDateTime = DateTime.Now;
            await this.requestRepository.AddAsync(request);

            // Get the new request adaptive card
            var hostDomain = this.azureADOptions.Value.HostDomain;
            var consultAttachment = CardFactory.CreateConsultAttachment(request, $"https://{hostDomain}", this.localizer);

            // Send the proactive message to the channel
            var conversationParameter = await ProactiveUtil.SendChannelProactiveMessageAsync(consultAttachment, channel.ChannelId, channel.ServiceUrl, this.botOptions.Value);

            // rebuild local request with id from object
            request.ConversationId = conversationParameter.Id;
            request.ActivityId = conversationParameter.ActivityId;

            // Save the new request in the database
            await this.requestRepository.UpsertAsync(request);

            return this.NoContent();
        }

        /// <summary>
        /// Send a reassign adaptive card in agent's Team channel.
        /// </summary>
        /// <param name="authorization">The sso token.</param>
        /// <param name="consultId">ID of the consult request to get.</param>
        /// <param name="body">Details of the consult reassignment.</param>
        /// <response code="200">Successful reassignment request of consult.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <response code="403">Consult request cannot be assigned. </response>
        /// <response code="404">Consult request not found.</response>
        /// <returns>An IActionResult indicating the result of the API call.</returns>
        [HttpPost]
        [Route("/api/request/{consultId}/reassign")]
        public async Task<IActionResult> ReassignRequestAsync([FromHeader] string authorization, string consultId, [FromBody] ReassignConsultRequestBody body)
        {
            if (string.IsNullOrEmpty(consultId))
            {
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Invalid consult ID provided" });
            }

            // Extract user's object ID from claims
            var userObjectId = this.GetUserObjectId();
            var userName = this.GetUserName();
            var currentTime = DateTime.UtcNow;
            var userObjectIdGuid = new Guid(userObjectId);

            // Load existing consult from DB
            var request = await this.requestRepository.GetAsync(new CosmosItemKey(consultId, consultId));
            if (request == null)
            {
                this.logger.LogError("Failed to reassign consult request {ConsultRequestId}. The consult request does not exist.", consultId);
                return this.NotFound(new UnsuccessfulResponse { Reason = "Consult request not found" });
            }

            if (request.Status != RequestStatus.Assigned)
            {
                this.logger.LogWarning("Failed to reassign consult request {ConsultRequestId}. The consult request is not in the 'Assigned' state.", consultId);
                return this.StatusCode(403, new UnsuccessfulResponse { Reason = "The consult cannot be reassigned." });
            }

            request.Status = RequestStatus.ReassignRequested;

            // Add activity to consult request
            request.Activities = request.Activities ?? new List<Activity>();
            request.Activities.Add(new Activity
            {
                Id = Guid.NewGuid(),
                Type = ActivityType.ReassignRequested,
                ActivityForUserId = request.AssignedToId,
                ActivityForUserName = request.AssignedToName,
                CreatedByName = userName,
                CreatedById = userObjectIdGuid,
                CreatedDateTime = currentTime,
            });

            // Store updated consult request
            await this.requestRepository.UpsertAsync(request);

            // Look up the channel to notify
            var mapping = await this.channelMappingRepository.GetByCategoryAsync(request.Category);

            if (mapping == null)
            {
                this.logger.LogError("Failed to send channel message for reassignment of consult {ConsultRequestId}. Channel mapping does not exist for request category {RequestCategory}", consultId, request.Category);
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Consult requests in this category cannot be reassigned" });
            }

            // Lookup the Channel
            var channel = await this.channelRepository.GetByChannelIdAsync(mapping.ChannelId);

            if (channel == null)
            {
                // throw error...Category does not exists.
                this.logger.LogError("Failed to send channel message for reassignment of consult {ConsultRequestId}. Channel {ChannelId} does not exist, but a mapping to that channel exists", consultId, mapping.ChannelId);
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Consult requests in this category cannot be reassigned" });
            }

            // Get the assigned consult adaptive card for channel
            var baseUrl = $"https://{this.azureADOptions.Value.HostDomain}";

            List<object> mentionedAgents = new List<object>();
            string photo = null;
            string ssoToken = authorization.Substring("Bearer".Length + 1);

            var displayPicGraphResponse = await GraphUtil.GetUserDisplayPhotoAsync(ssoToken, userObjectId, this.azureADOptions.Value);
            if (displayPicGraphResponse.FailureReason == string.Empty)
            {
                photo = "data:image/jpeg;base64," + Convert.ToBase64String(displayPicGraphResponse.Result);
            }

            var comment = body.Comments == null ? string.Empty : body.Comments;

            foreach (var agentIdName in body.Agents)
            {
                var agent = await this.agentRepository.GetByObjectIdAsync(agentIdName.Id);

                mentionedAgents.Add(new
                {
                    agentUserId = agent.TeamsId,
                    agentName = agentIdName.DisplayName,
                });
            }

            var reassignConsultChannelCard = CardFactory.CreateReassignConsultAttachment(request, mentionedAgents, baseUrl, comment, userName, photo, this.localizer);

            // Update the channel message, and reply to it so that it gets pushed to the bottom
            var replyString = this.localizer.GetString("ConsultNeedsReassignment", userName);
            await ProactiveUtil.UpdateChannelProactiveMessageAsync(reassignConsultChannelCard, channel.ServiceUrl, request.ConversationId, request.ActivityId, this.botOptions.Value);
            await ProactiveUtil.ReplyToChannelMessageAsync(replyString, channel.ServiceUrl, request.ConversationId, request.ActivityId, this.botOptions.Value);

            return this.Ok(request);
        }

        /// <summary>
        /// Assign a consult to the calling agent.
        /// </summary>
        /// <param name="consultId">ID of the consult request to get.</param>
        /// <param name="requestBody">Details of the consult assignment.</param>
        /// <param name="authorization">The Authorization header.</param>
        /// <response code="200">Successful assignment of consult request.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <response code="403">Consult request cannot be assigned. </response>
        /// <response code="404">Consult request not found.</response>
        /// <returns>An IActionResult indicating the result of the API call.</returns>
        [HttpPost]
        [Route("/api/request/{consultId}/assign")]
        public async Task<IActionResult> AssignRequestAsync(string consultId, [FromBody] AssignConsultRequestBody requestBody, [FromHeader] string authorization)
        {
            string ssoToken = authorization.Substring("Bearer".Length + 1);

            // Extract user's object ID from claims
            var assignerObjectId = this.GetUserObjectId();
            var assignerUserName = this.GetUserName();
            if (string.IsNullOrWhiteSpace(assignerObjectId))
            {
                this.logger.LogError("Failed to assign consult request. The assigner's object ID was null or empty.");
                return this.Unauthorized();
            }

            if (string.IsNullOrEmpty(consultId))
            {
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Invalid consult ID provided" });
            }

            // Handle assign to another agent
            var assigneeObjectId = assignerObjectId;
            if (requestBody.Agent != null && !string.IsNullOrEmpty(requestBody.Agent.Id) && !string.IsNullOrEmpty(requestBody.Agent.DisplayName))
            {
                var canAssignOther = await this.CheckSupervisorAsync(consultId);
                if (!canAssignOther)
                {
                    return this.Unauthorized();
                }

                assigneeObjectId = requestBody.Agent.Id;
            }

            // Lookup the agent
            var assigneeAgent = await this.agentRepository.GetByObjectIdAsync(assigneeObjectId);
            if (assigneeAgent == null)
            {
                this.logger.LogError("Failed to assign consult request. Agent with ID {UserObjectId} does not exist", assigneeObjectId);
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Invalid agent provided" });
            }

            // Set assignee user name based on self-assign vs assign to other
            var assigneeUserName = assigneeObjectId == assignerObjectId ? assignerUserName : assigneeAgent.Name;

            // Load existing consult from DB
            var request = await this.requestRepository.GetAsync(new CosmosItemKey(consultId, consultId));
            if (request == null)
            {
                this.logger.LogError("Failed to assign consult request {ConsultRequestId}. The consult request does not exist.", consultId);
                return this.NotFound(new UnsuccessfulResponse { Reason = "Consult request not found" });
            }

            if (request.Status != RequestStatus.Unassigned && request.Status != RequestStatus.ReassignRequested)
            {
                this.logger.LogWarning("Failed to assign consult request {ConsultRequestId}. The consult request is already assigned.", consultId);
                return this.StatusCode(403, new UnsuccessfulResponse { Reason = "The consult is already assigned" });
            }

            // Update consult request fields from request body
            request.AssignedTimeBlock = requestBody.SelectedTimeBlock;

            string assigneeStaffMemberId = assigneeAgent.BookingsStaffMemberId;
            GraphBeta.BookingAppointment bookingsResult = null;
            bool didBookingsSucceed = false;
            bool didRetrieveNewStaffMemberId = false;
            while (!didBookingsSucceed && !didRetrieveNewStaffMemberId)
            {
                // If we don't have the agent's staff member ID, try getting it from Graph
                if (string.IsNullOrEmpty(assigneeStaffMemberId))
                {
                    try
                    {
                        assigneeStaffMemberId = await GraphUtil.GetBookingsStaffMemberId(ssoToken, request.BookingsBusinessId, assigneeAgent.UserPrincipalName, this.azureADOptions.Value);
                    }
                    catch (Graph.ServiceException)
                    {
                        // This failure could also be because the caller doesn't have permission
                        this.logger.LogError("Failed to assign consult request {ConsultRequestId}. Bookings staff member ID for agent {AgentUPN} could not be retrieved.", consultId, assigneeAgent.UserPrincipalName);
                        return this.BadRequest(new UnsuccessfulResponse { Reason = "The assignee is not a valid staff member in Bookings." });
                    }

                    if (string.IsNullOrEmpty(assigneeStaffMemberId))
                    {
                        this.logger.LogError("Failed to assign consult request {ConsultRequestId}. Bookings staff member ID for agent {AgentUPN} could not be retrieved.", consultId, assigneeAgent.UserPrincipalName);
                        return this.BadRequest(new UnsuccessfulResponse { Reason = "The assignee is not a valid staff member in Bookings." });
                    }

                    didRetrieveNewStaffMemberId = true;
                }

                // Create/update Bookings appointment
                try
                {
                    if (request.Status == RequestStatus.Unassigned)
                    {
                        bookingsResult = await GraphUtil.CreateBookingsAppointment(ssoToken, request, assigneeStaffMemberId, this.azureADOptions.Value);
                    }
                    else
                    {
                        await GraphUtil.UpdateBookingsAppointment(ssoToken, request, assigneeStaffMemberId, this.azureADOptions.Value);
                    }

                    didBookingsSucceed = true;
                }
                catch (Graph.ServiceException)
                {
                    // Will try retrieving a fresh staff member ID
                    assigneeStaffMemberId = null;
                }
            }

            // Ensure appointment creation succeeded
            if (!didBookingsSucceed)
            {
                this.logger.LogError("Failed to assign consult request {ConsultRequestId}. The Bookings appointment could not be created/updated for agent {AgentId}.", consultId, assigneeObjectId);
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Unable to create/update the Bookings appointment." });
            }

            // Update DB with agent's staff member ID, if needed
            if (assigneeAgent.BookingsStaffMemberId != assigneeStaffMemberId)
            {
                assigneeAgent.BookingsStaffMemberId = assigneeStaffMemberId;
                await this.agentRepository.UpsertAsync(assigneeAgent);
            }

            // Update consult request fields with Bookings info, if needed
            if (request.Status == RequestStatus.Unassigned)
            {
                request.BookingsAppointmentId = bookingsResult.Id;
                request.JoinUri = bookingsResult.OnlineMeetingUrl;
            }

            // Update other consult request fields
            var currentTime = DateTime.UtcNow;
            request.Status = RequestStatus.Assigned;
            request.AssignedToId = assigneeObjectId;
            request.AssignedToName = assigneeUserName;
            request.Activities = request.Activities ?? new List<Activity>();
            request.Activities.Add(new Activity
            {
                Id = Guid.NewGuid(),
                Type = ActivityType.Assigned,
                ActivityForUserId = assigneeObjectId,
                ActivityForUserName = assigneeUserName,
                Comment = !string.IsNullOrWhiteSpace(requestBody.Comments) ? requestBody.Comments : null,
                CreatedByName = assignerUserName,
                CreatedById = new Guid(assignerObjectId),
                CreatedDateTime = currentTime,
            });

            // Update consult request in DB
            await this.requestRepository.UpsertAsync(request);

            // Look up the channel to notify
            var mapping = await this.channelMappingRepository.GetByCategoryAsync(request.Category);

            if (mapping == null)
            {
                this.logger.LogError("Failed to send channel message for assignment of consult {ConsultRequestId}. Channel mapping does not exist for request category {RequestCategory}.", consultId, request.Category);
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Consult requests in this category cannot be assigned" });
            }

            // Lookup the Channel
            var channel = await this.channelRepository.GetByChannelIdAsync(mapping.ChannelId);

            if (channel == null)
            {
                this.logger.LogError("Failed to send channel message for assignment of consult {ConsultRequestId}. Channel {ChannelId} does not exist, but a mapping to that channel exists.", consultId, mapping.ChannelId);
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Consult requests in this category cannot be reassigned" });
            }

            // Get the assigned consult adaptive card for channel
            var baseUrl = $"https://{this.azureADOptions.Value.HostDomain}";
            var consultChannelCard = CardFactory.CreateAssignedConsultAttachment(request, baseUrl, assigneeUserName, false, this.localizer);

            // Update the proactive message to the channel
            await ProactiveUtil.UpdateChannelProactiveMessageAsync(consultChannelCard, channel.ServiceUrl, request.ConversationId, request.ActivityId, this.botOptions.Value);

            // Temporarily set agent's locale as current culture before using localizer
            Microsoft.Bot.Schema.Attachment consultPersonalCard;
            using (new CultureSwitcher(assigneeAgent.Locale, assigneeAgent.Locale))
            {
                // Get the assigned consult adaptive card for 1:1 chat
                consultPersonalCard = CardFactory.CreateAssignedConsultAttachment(request, baseUrl, assigneeUserName, true, this.localizer);
            }

            // Send the proactive message to the agent
            await ProactiveUtil.SendChatProactiveMessageAsync(consultPersonalCard, assigneeAgent.TeamsId, this.azureADOptions.Value.TenantId, assigneeAgent.ServiceUrl, this.botOptions.Value);

            return this.Ok(request);
        }

        /// <summary>
        /// Mark a consult as completed.
        /// </summary>
        /// <param name="consultId">ID of the consult request to complete.</param>
        /// <response code="200">Successfully marked consult request as complete.</response>
        /// <response code="400">Invalid input data provided.</response>
        /// <response code="403">Consult request cannot be completed.</response>
        /// <response code="404">Consult request not found.</response>
        /// <returns>An IActionResult indicating the result of the API call.</returns>
        [HttpPost]
        [Route("/api/request/{consultId}/complete")]
        public async Task<IActionResult> CompleteRequestAsync(string consultId)
        {
            var userObjectId = this.GetUserObjectId();
            var userName = this.GetUserName();
            if (string.IsNullOrWhiteSpace(userObjectId))
            {
                this.logger.LogError("Failed to complete consult request. The user's object ID was null or empty.");
                return this.Unauthorized();
            }

            if (string.IsNullOrEmpty(consultId))
            {
                return this.BadRequest(new UnsuccessfulResponse { Reason = "Invalid consult ID provided" });
            }

            var request = await this.requestRepository.GetAsync(new CosmosItemKey(consultId, consultId));
            if (request == null)
            {
                this.logger.LogError("Failed to complete consult request {ConsultRequestId}. The consult request does not exist.", consultId);
                return this.NotFound(new UnsuccessfulResponse { Reason = "Consult request not found" });
            }

            if (request.Status != RequestStatus.Assigned && request.Status != RequestStatus.ReassignRequested)
            {
                this.logger.LogWarning("Failed to complete consult request {ConsultRequestId}. The request is not assigned.", consultId);
                return this.StatusCode(403, new UnsuccessfulResponse { Reason = "The consult cannot be completed." });
            }

            var currentTime = DateTime.UtcNow;
            var userObjectIdGuid = new Guid(userObjectId);

            // Change request status to completed
            request.Status = RequestStatus.Completed;
            request.Activities = request.Activities ?? new List<Activity>();
            request.Activities.Add(new Activity
            {
                Id = Guid.NewGuid(),
                Type = ActivityType.Completed,
                CreatedByName = userName,
                CreatedById = userObjectIdGuid,
                CreatedDateTime = currentTime,
            });

            // Update request in DB
            await this.requestRepository.UpsertAsync(request);

            return this.Ok(request);
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

        /// <summary>
        /// Returns random generated friendly id (not globally unique)
        /// </summary>
        /// <returns>random 6 character string</returns>
        private string GetFriendlyId()
        {
            Random generator = new Random();
            var r = generator.Next(0, 999999).ToString("D6");
            return r;
        }

        /// <summary>
        /// Checks if the user is a supervisor for the specified request.
        /// </summary>
        /// <param name="consultId">Request ID for the request to check.</param>
        private async Task<bool> CheckSupervisorAsync(string consultId)
        {
            // start by getting the request
            var request = await this.requestRepository.GetAsync(new CosmosItemKey(consultId, consultId));
            if (request == null)
            {
                this.logger.LogError("Failed to check supervisors for {ConsultRequestId}. The consult request does not exist.", consultId);
                return false;
            }

            // get the channelMapping for the request category
            var mapping = await this.channelMappingRepository.GetByCategoryAsync(request.Category);
            if (mapping == null)
            {
                // mapping for this category was not found
                this.logger.LogError("Failed to check supervisors for {ConsultRequestId}. No mapping exists for consult category {ConsultCategory}.", consultId, request.Category);
                return false;
            }
            else
            {
                // check if the user is in the supervisors list for this category
                var oidClaim = this.User.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
                return mapping.Supervisors.Count == 0 || mapping.Supervisors.Exists(i => i.Id == oidClaim);
            }
        }
    }
}
// <copyright file="UserAvailabilityController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Graph.Extensions;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration;
    using Microsoft.Teams.App.VirtualConsult.Common.Utils;

    /// <summary>
    /// Web API to get user availability.
    /// </summary>
    [ApiController]
    [Authorize]
    public class UserAvailabilityController : ControllerBase
    {
        private readonly ILogger<UserAvailabilityController> logger;
        private readonly IOptions<AzureADSettings> azureADOptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserAvailabilityController"/> class.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="azureADOptions">Azure AD configuration options.</param>
        public UserAvailabilityController(ILogger<UserAvailabilityController> logger, IOptions<AzureADSettings> azureADOptions)
        {
            this.logger = logger;
            this.azureADOptions = azureADOptions;
        }

        /// <summary>
        /// Get user availability with Graph access token.
        /// </summary>
        /// <param name="authorization">The SSO token.</param>
        /// <param name="constrains">The meeting constrain times.</param>
        /// <response code="200">Successfully get user availability.</response>
        /// <response code="400">Error occured while fetching user availability.</response>
        /// <returns>Graph access token.</returns>
        [HttpPost]
        [Route("/api/availability")]
        public async Task<IActionResult> PostAsync([FromHeader] string authorization, [FromBody] IEnumerable<TimeBlock> constrains)
        {
            return await this.GetAvailabilityAsync(authorization, constrains, null);
        }

        /// <summary>
        /// Get user availability with Graph access token.
        /// </summary>
        /// <param name="authorization">The SSO token.</param>
        /// <param name="constrains">The meeting constrain times.</param>
        /// <param name="teamAadObjectId">The group AAD Id where the agents are located.</param>
        /// <response code="200">Successfully get user availability.</response>
        /// <response code="400">Error occured while fetching user availability.</response>
        /// <returns>Graph access token.</returns>
        [HttpPost]
        [Route("/api/availability/{teamAadObjectId}")]
        public async Task<IActionResult> PostAsync([FromHeader] string authorization, [FromBody] IEnumerable<TimeBlock> constrains, string teamAadObjectId)
        {
            return await this.GetAvailabilityAsync(authorization, constrains, teamAadObjectId);
        }

        /// <summary>
        /// Get user meeting details for Graph access token.
        /// </summary>
        /// <param name="authorization">The SSO token.</param>
        /// <param name="constraints">The meeting constrain times.</param>
        /// <response code="200">Successfully get meeting details for agent.</response>
        /// <response code="400">Error occured while fetching meeting deatils for agent.</response>
        /// <returns>Graph access token.</returns>
        [HttpPost]
        [Route("/api/meetingDetails")]
        public async Task<IActionResult> GetMeetingDetailsAsync([FromHeader] string authorization, [FromBody] TimeBlock constraints)
        {
            string ssoToken = authorization.Substring("Bearer".Length + 1);
            var graphMeetingDetailsResponse = await GraphUtil.GetCalendarViewAsync(ssoToken, constraints, this.azureADOptions.Value);

            if (!graphMeetingDetailsResponse.FailureReason.Equals(string.Empty))
            {
                this.logger.LogError($"Failed to get meeting details for agent. The Graph call failed: {graphMeetingDetailsResponse.FailureReason}");
                return this.StatusCode(400, new UnsuccessfulResponse { Reason = graphMeetingDetailsResponse.FailureReason });
            }

            // Graph /calendarView endpoint returns results that are inclusive of the given startDateTime
            // We filter out events that are adjacent to (but not inside of) the constraint
            var filteredResults = graphMeetingDetailsResponse.Result
                .Where(meeting => meeting.MeetingTime.StartDateTime < constraints.EndDateTime && meeting.MeetingTime.EndDateTime > constraints.StartDateTime);

            return this.Ok(filteredResults);
        }

        private async Task<IActionResult> GetAvailabilityAsync(string authorization, IEnumerable<TimeBlock> constrains, string teamAadObjectId)
        {
            string ssoToken = authorization.Substring("Bearer".Length + 1);

            // build graphConstraints based on timeblocks passed in
            IEnumerable<Graph.TimeSlot> graphConstraints = constrains.Select(constraint =>
            {
                return new Graph.TimeSlot
                {
                    Start = constraint.StartDateTime.UtcDateTime.ToDateTimeTimeZone(TimeZoneInfo.Utc),
                    End = constraint.EndDateTime.UtcDateTime.ToDateTimeTimeZone(TimeZoneInfo.Utc),
                };
            });

            // check if we are getting availability for the current user or within a team
            List<string> emailAddresses = new List<string>();
            List<Graph.User> teamMembers = new List<Graph.User>();
            if (!string.IsNullOrEmpty(teamAadObjectId))
            {
                // add all agents in the team
                teamMembers = (await GraphUtil.GetMembersInTeamsChannelAsync(teamAadObjectId, this.azureADOptions.Value)).Result.ToList();
                emailAddresses = teamMembers.Select(teammember => teammember.Mail).ToList();
            }

            // make graph call to get availability
            AgentAvailability organizer = new AgentAvailability()
            {
                // TODO: move these claim keys to a consts file
                Id = this.User.Claims.FirstOrDefault(i => i.Type == "http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value,
                DisplayName = this.User.Claims.FirstOrDefault(i => i.Type == "name")?.Value,
                Emailaddress = this.User.Claims.FirstOrDefault(i => i.Type == "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn")?.Value,
                TimeBlocks = new List<TimeBlock>(),
            };
            var graphTimeSlotsResponse = await GraphUtil.GetAvailableTimeSlotsAsync(ssoToken, graphConstraints, organizer, this.azureADOptions.Value, 30, emailAddresses);

            // process results into a list of AgentAvailability objects
            List<AgentAvailability> availability = new List<AgentAvailability>();
            foreach (var email in graphTimeSlotsResponse.Keys)
            {
                // look up the team member...if null this is the organizer
                var member = teamMembers.FirstOrDefault(i => i.Mail == email);
                if (member == null)
                {
                    organizer.TimeBlocks = graphTimeSlotsResponse[email];
                    availability.Add(organizer);
                }
                else
                {
                    availability.Add(new AgentAvailability()
                    {
                        Id = member.Id,
                        DisplayName = member.DisplayName,
                        Emailaddress = email,
                        TimeBlocks = graphTimeSlotsResponse[email],
                    });
                }
            }

            // if for self and no availability return the organizer
            if (string.IsNullOrEmpty(teamAadObjectId) && availability.Count == 0)
            {
                availability.Add(organizer);
            }

            return this.Ok(availability);
        }
    }
}
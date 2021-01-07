// <copyright file="GraphUtil.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Utils
{
    extern alias GraphBetaLib;

    using System;
    using System.Collections.Generic;
    using System.Collections.Immutable;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Microsoft.Graph.Auth;
    using Microsoft.Graph.Extensions;
    using Microsoft.Identity.Client;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration;
    using GraphBeta = GraphBetaLib::Microsoft.Graph;

    /// <summary>
    /// Token exchange class to exchange the SSO token with a graph access token.
    /// </summary>
    public static class GraphUtil
    {
        /// <summary>
        /// Gets the users in a teams channel defined by channel id.
        /// </summary>
        /// <param name="teamAadObjectId">The AadObjectId id of the channel to grab users from.</param>
        /// <param name="azureADSettings">Azure AD configuration settings.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public static async Task<GraphResponse<IEnumerable<User>>> GetMembersInTeamsChannelAsync(string teamAadObjectId, AzureADSettings azureADSettings)
        {
            var authProvider = CreateClientCredentialProvider(azureADSettings);
            GraphServiceClient graphServiceClient = new GraphServiceClient(authProvider);
            GraphResponse<IEnumerable<User>> response = new GraphResponse<IEnumerable<User>>();

            try
            {
                var page = await graphServiceClient.Groups[teamAadObjectId]
                    .Members
                    .Request()
                    .GetAsync();

                List<User> result = new List<User>();

                while (page != null)
                {
                    var usersList = page.CurrentPage.ToList();
                    foreach (var user in usersList)
                    {
                        var currentUserObj = user as User;
                        result.Add(currentUserObj);
                    }

                    if (page.NextPageRequest == null)
                    {
                        break;
                    }

                    page = await page.NextPageRequest.GetAsync();
                }

                response.Result = result.ToImmutableList();

                return response;
            }
            catch (Exception e)
            {
                response.FailureReason = e.Message;
                response.Result = new List<User>();
            }

            return response;
        }

        /// <summary>
        /// Gets the Agents Calendar view.
        /// </summary>
        /// <param name="ssoToken">The sso access token used to make request against graph apis.</param>
        /// <param name="constraints">The time constraint for finding calendar view.</param>
        /// /// <param name="azureADSettings">Azure AD configuration settings.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public static async Task<GraphResponse<IEnumerable<MeetingDetails>>> GetCalendarViewAsync(string ssoToken, TimeBlock constraints, AzureADSettings azureADSettings)
        {
            var authProvider = CreateOnBehalfOfProvider(azureADSettings, new[] { "Calendars.Read" });
            GraphServiceClient graphServiceClient = new GraphServiceClient(authProvider);

            var startDateTimeParam = Uri.EscapeDataString(constraints.StartDateTime.ToString("o"));
            var endDateTimeParam = Uri.EscapeDataString(constraints.EndDateTime.ToString("o"));
            var queryOptions = new List<QueryOption>()
                      {
                           new QueryOption("startdatetime", startDateTimeParam),
                           new QueryOption("enddatetime", endDateTimeParam),
                      };

            GraphResponse<IEnumerable<MeetingDetails>> response = new GraphResponse<IEnumerable<MeetingDetails>>();

            try
            {
                var page = await graphServiceClient.Me
                    .CalendarView.Request(queryOptions)
                    .WithUserAssertion(new UserAssertion(ssoToken))
                    .GetAsync();

                List<MeetingDetails> result = new List<MeetingDetails>();
                while (page != null)
                {
                    var meetingList = page.CurrentPage;
                    foreach (var meeting in meetingList)
                    {
                        result.Add(new MeetingDetails
                        {
                            Subject = meeting.Subject,
                            MeetingTime = new TimeBlock
                            {
                                StartDateTime = meeting.Start.ToDateTimeOffset().ToUniversalTime(),
                                EndDateTime = meeting.End.ToDateTimeOffset().ToUniversalTime(),
                            },
                        });
                    }

                    if (page.NextPageRequest == null)
                    {
                        break;
                    }

                    page = await page.NextPageRequest.GetAsync();
                }

                response.Result = result.ToImmutableList();
                return response;
            }
            catch (Exception e)
            {
                response.FailureReason = e.Message;
                response.Result = new List<MeetingDetails>();
            }

            return response;
        }

        /// <summary>
        /// Gets the user available times.
        /// </summary>
        /// <param name="ssoToken">The sso access token used to make request against graph apis.</param>
        /// <param name="constraints">The time constraint for finding free times.</param>
        /// <param name="organizer">The organizer the ssoToken belongs to.</param>
        /// <param name="azureADSettings">The Azure AD application settings.</param>
        /// <param name="maxTimeSlots">The maximum number of time slots to return.</param>
        /// <param name="emailAddresses">Emails.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public static async Task<Dictionary<string, List<TimeBlock>>> GetAvailableTimeSlotsAsync(string ssoToken, IEnumerable<TimeSlot> constraints, AgentAvailability organizer, AzureADSettings azureADSettings, int? maxTimeSlots = null, IEnumerable<string> emailAddresses = null)
        {
            // setup graph client
            var authProvider = CreateOnBehalfOfProvider(azureADSettings, new[] { "Calendars.Read" });
            GraphServiceClient graphServiceClient = new GraphServiceClient(authProvider);

            // define the time constraints
            var timeConstaint = new TimeConstraint
            {
                ActivityDomain = ActivityDomain.Work,
                TimeSlots = constraints,
            };

            // add the attendees
            var attendees = new List<AttendeeBase>();
            foreach (var address in emailAddresses)
            {
                var attendee = new AttendeeBase
                {
                    Type = AttendeeType.Required,
                    EmailAddress = new EmailAddress
                    {
                        Address = address,
                    },
                };

                attendees.Add(attendee);
            }

            bool isMe = emailAddresses.Count() == 0;

            try
            {
                // perform the graph call for find meeting times
                var result = await graphServiceClient.Me.FindMeetingTimes(timeConstraint: timeConstaint, maxCandidates: maxTimeSlots, attendees: attendees, isOrganizerOptional: !isMe, minimumAttendeePercentage: 10)
                    .Request()
                    .WithUserAssertion(new UserAssertion(ssoToken))
                    .PostAsync();

                // return an empty collection if EmptySuggestionsReason returned
                if (!result.EmptySuggestionsReason.Equals(string.Empty))
                {
                    return new Dictionary<string, List<TimeBlock>>();
                }

                // pivot the results from timeblock centric to user centric
                Dictionary<string, List<TimeBlock>> dictionary = new Dictionary<string, List<TimeBlock>>();
                foreach (var timeslot in result.MeetingTimeSuggestions.ToList())
                {
                    if (isMe)
                    {
                        if (timeslot.OrganizerAvailability != FreeBusyStatus.Free)
                        {
                            continue;
                        }

                        if (!dictionary.ContainsKey(organizer.Emailaddress))
                        {
                            dictionary.Add(organizer.Emailaddress, new List<TimeBlock>());
                        }

                        dictionary[organizer.Emailaddress].Add(new TimeBlock()
                        {
                            StartDateTime = timeslot.MeetingTimeSlot.Start.ToDateTimeOffset().ToUniversalTime(),
                            EndDateTime = timeslot.MeetingTimeSlot.End.ToDateTimeOffset().ToUniversalTime(),
                        });
                    }
                    else
                    {
                        foreach (var agent in timeslot.AttendeeAvailability)
                        {
                            if (agent.Availability.GetValueOrDefault(FreeBusyStatus.Unknown) != FreeBusyStatus.Free)
                            {
                                continue;
                            }

                            if (!dictionary.ContainsKey(agent.Attendee.EmailAddress.Address))
                            {
                                dictionary.Add(agent.Attendee.EmailAddress.Address, new List<TimeBlock>());
                            }

                            dictionary[agent.Attendee.EmailAddress.Address].Add(new TimeBlock()
                            {
                                StartDateTime = timeslot.MeetingTimeSlot.Start.ToDateTimeOffset().ToUniversalTime(),
                                EndDateTime = timeslot.MeetingTimeSlot.End.ToDateTimeOffset().ToUniversalTime(),
                            });
                        }
                    }
                }

                return dictionary;
            }
            catch (Exception)
            {
                return new Dictionary<string, List<TimeBlock>>();
            }
        }

        /// <summary>
        /// Gets the user's display photo.
        /// </summary>
        /// <param name="ssoToken">The sso access token used to make request against graph apis.</param>
        /// <param name="userId">The user to get the .</param>
        /// /// <param name="azureADSettings">Azure AD configuration settings.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public static async Task<GraphResponse<byte[]>> GetUserDisplayPhotoAsync(string ssoToken, string userId, AzureADSettings azureADSettings)
        {
            var authProvider = CreateOnBehalfOfProvider(azureADSettings, new[] { "User.ReadBasic.All" });
            GraphServiceClient graphServiceClient = new GraphServiceClient(authProvider);
            GraphResponse<byte[]> response = new GraphResponse<byte[]>();

            try
            {
                var result = await graphServiceClient.Users[userId].Photos["48x48"].Content.Request().WithUserAssertion(new UserAssertion(ssoToken)).GetAsync();
                if (result == null)
                {
                    response.FailureReason = "Unable to get the user's display photo";
                    response.Result = null;

                    return response;
                }

                byte[] bytes = new byte[result.Length];
                result.Read(bytes, 0, (int)result.Length);

                response.Result = bytes;
                return response;
            }
            catch (Exception e)
            {
                response.FailureReason = e.Message;
                response.Result = null;
            }

            return response;
        }

        /// <summary>
        /// Create a Bookings appointment for the given consult request.
        /// </summary>
        /// <param name="ssoToken">The SSO token of the calling agent.</param>
        /// <param name="request">The consult request to book.</param>
        /// <param name="assignedStaffMemberId">The Bookings staff member ID of the assigned agent.</param>
        /// <param name="azureADSettings">Azure AD configuration settings.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation. The task result contains the created <see cref="GraphBeta.BookingAppointment"/>.</returns>
        public static async Task<GraphBeta.BookingAppointment> CreateBookingsAppointment(string ssoToken, Request request, string assignedStaffMemberId, AzureADSettings azureADSettings)
        {
            var authProvider = CreateOnBehalfOfProvider(azureADSettings, new[] { "BookingsAppointment.ReadWrite.All" });
            var graphServiceClient = new GraphBeta.GraphServiceClient(authProvider);

            var bookingAppointment = new GraphBeta.BookingAppointment
            {
                CustomerEmailAddress = request.CustomerEmail,
                CustomerName = request.CustomerName,
                CustomerPhone = request.CustomerPhone,
                Start = GraphBeta.DateTimeTimeZone.FromDateTimeOffset(request.AssignedTimeBlock.StartDateTime.ToUniversalTime(), System.TimeZoneInfo.Utc),
                End = GraphBeta.DateTimeTimeZone.FromDateTimeOffset(request.AssignedTimeBlock.EndDateTime.ToUniversalTime(), System.TimeZoneInfo.Utc),
                OptOutOfCustomerEmail = false,
                ServiceId = request.BookingsServiceId,
                StaffMemberIds = new List<string> { assignedStaffMemberId },
                IsLocationOnline = true,
            };

            var bookingResult = await graphServiceClient.BookingBusinesses[request.BookingsBusinessId].Appointments
                .Request()
                .WithUserAssertion(new UserAssertion(ssoToken))
                .AddAsync(bookingAppointment);

            return bookingResult;
        }

        /// <summary>
        /// Update a Bookings appointment.
        /// </summary>
        /// <param name="ssoToken">The SSO token.</param>
        /// <param name="request">The consult request to book.</param>
        /// <param name="assignedStaffMemberId">The Bookings staff member ID of the assigned agent.</param>
        /// <param name="azureADSettings">Azure AD configuration settings.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        public static async Task UpdateBookingsAppointment(string ssoToken, Request request, string assignedStaffMemberId, AzureADSettings azureADSettings)
        {
            var authProvider = CreateOnBehalfOfProvider(azureADSettings, new[] { "BookingsAppointment.ReadWrite.All" });
            var graphServiceClient = new GraphBeta.GraphServiceClient(authProvider);

            var bookingAppointment = new GraphBeta.BookingAppointment
            {
                StaffMemberIds = new List<string> { assignedStaffMemberId },
            };

            await graphServiceClient.BookingBusinesses[request.BookingsBusinessId].Appointments[request.BookingsAppointmentId]
                .Request()
                .WithUserAssertion(new UserAssertion(ssoToken))
                .UpdateAsync(bookingAppointment);
        }

        /// <summary>
        /// Gets the Bookings staff member ID for the given user.
        /// </summary>
        /// <param name="ssoToken">The SSO token.</param>
        /// <param name="businessId">The ID for the Bookings business to use.</param>
        /// <param name="userPrincipalName">The UPN of the user.</param>
        /// <param name="azureADSettings">Azure AD configuration settings.</param>
        /// <returns>The Bookings staff member ID for the given user, or null if not found.</returns>
        public static async Task<string> GetBookingsStaffMemberId(string ssoToken, string businessId, string userPrincipalName, AzureADSettings azureADSettings)
        {
            var authProvider = CreateOnBehalfOfProvider(azureADSettings, new[] { "BookingsAppointment.ReadWrite.All" });
            var graphServiceClient = new GraphBeta.GraphServiceClient(authProvider);

            var staffMembersRequest = graphServiceClient.BookingBusinesses[businessId].StaffMembers
                .Request()
                .Select(staffMember => new { staffMember.Id })
                .Filter($"emailAddress eq '{userPrincipalName}'")
                .WithUserAssertion(new UserAssertion(ssoToken));

            var staffMembers = await staffMembersRequest.GetAsync();

            // We expect exactly one staff member to match
            if (staffMembers.Count != 1)
            {
                return null;
            }

            return staffMembers.FirstOrDefault()?.Id;
        }

        private static OnBehalfOfProvider CreateOnBehalfOfProvider(AzureADSettings aadSettings, string[] scopes)
        {
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(aadSettings.AppId)
                .WithClientSecret(aadSettings.AppPassword)
                .WithTenantId(aadSettings.TenantId)
                .Build();

            return new OnBehalfOfProvider(confidentialClientApplication, scopes);
        }

        private static ClientCredentialProvider CreateClientCredentialProvider(AzureADSettings aadSettings)
        {
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(aadSettings.AppId)
                .WithClientSecret(aadSettings.AppPassword)
                .WithTenantId(aadSettings.TenantId)
                .Build();

            return new ClientCredentialProvider(confidentialClientApplication);
        }
    }
}
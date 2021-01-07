// <copyright file="BotActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration;
    using Microsoft.Teams.App.VirtualConsult.Common.Repositories;
    using Microsoft.Teams.App.VirtualConsult.Common.Utils;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// The SourceActivityHandler is responsible for reacting to incoming events from Teams sent from BotFramework.
    /// </summary>
    public sealed class BotActivityHandler : TeamsActivityHandler
    {
        private readonly IRequestRepository<CosmosItemKey> requestRepository;

        private readonly IAgentRepository<CosmosItemKey> agentRepository;

        private readonly IChannelRepository<CosmosItemKey> channelRepository;

        private readonly IChannelMappingRepository<CosmosItemKey> channelMappingRepository;

        private readonly ILogger<BotActivityHandler> logger;

        private readonly IStringLocalizer<SharedResources> localizer;

        private readonly IOptions<AzureADSettings> azureADOptions;

        private readonly IOptions<TeamsSettings> teamsOptions;

        private readonly IOptions<BotSettings> botOptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="BotActivityHandler"/> class.
        /// </summary>
        /// <param name="requestRepository">The request repository.</param>
        /// <param name="agentRepository">The agent repository.</param>
        /// <param name="channelRepository">The channel repository.</param>
        /// <param name="channelMappingRepository">The channel mapping repository.</param>
        /// <param name="logger">The logger.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="azureADOptions">Azure AD configuration options.</param>
        /// <param name="teamsOptions">Teams configuration options.</param>
        /// <param name="botOptions">Bot configuration options.</param>
        public BotActivityHandler(
            IRequestRepository<CosmosItemKey> requestRepository,
            IAgentRepository<CosmosItemKey> agentRepository,
            IChannelRepository<CosmosItemKey> channelRepository,
            IChannelMappingRepository<CosmosItemKey> channelMappingRepository,
            ILogger<BotActivityHandler> logger,
            IStringLocalizer<SharedResources> localizer,
            IOptions<AzureADSettings> azureADOptions,
            IOptions<TeamsSettings> teamsOptions,
            IOptions<BotSettings> botOptions)
        {
            this.requestRepository = requestRepository;
            this.agentRepository = agentRepository;
            this.channelRepository = channelRepository;
            this.channelMappingRepository = channelMappingRepository;
            this.logger = logger;
            this.localizer = localizer;
            this.azureADOptions = azureADOptions;
            this.teamsOptions = teamsOptions;
            this.botOptions = botOptions;
        }

        /// <summary>
        /// Handle when a message is addressed to the bot.
        /// </summary>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A Task resolving to either a login card or the adaptive card of the Reddit post.</returns>
        /// <remarks>
        /// For more information on bot messaging in Teams, see the documentation
        /// https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/conversation-basics?tabs=dotnet#receive-a-message .
        /// </remarks>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

            // Switch to user's locale for personal messages
            CultureSwitcher cultureSwitcher = null;
            var conversationType = turnContext.Activity.Conversation.ConversationType.ToEnum<ConversationType>();
            if (conversationType == ConversationType.Personal)
            {
                cultureSwitcher = new CultureSwitcher(turnContext.Activity.Locale, turnContext.Activity.Locale);
            }

            // The bot doesn't respond to direct message activities, so send the user a friendly error message
            string message;
            using (cultureSwitcher)
            {
                message = this.localizer.GetString("MessageDirectMessageError");
            }

            var activity = MessageFactory.Text(message);
            await turnContext.SendActivityAsync(activity, cancellationToken);
        }

        /// <summary>
        /// Invoked to handle the composeExtension/fetchTask event.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="action">The requested action.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            if (action.MessagePayload == null)
            {
                this.logger.LogError("Failed to fetch messaging extension task module. The message payload was null");
                return default;
            }

            var locale = turnContext.Activity.GetLocale();
            var conversationId = this.GetConversationIdFromMessageLink(action.MessagePayload.LinkToMessage.AbsoluteUri);

            // Ensure that messaging extension is being invoked from a meeting conversation
            if (string.IsNullOrWhiteSpace(conversationId))
            {
                string outsideMeetingErrorTitle;
                string outsideMeetingErrorMsg;
                using (new CultureSwitcher(locale, locale))
                {
                    outsideMeetingErrorTitle = this.localizer.GetString("TaskModuleTitleAttach");
                    outsideMeetingErrorMsg = this.localizer.GetString("MessagingExtensionOutsideMeetingError");
                }

                var outsideMeetingErrorCard = CardFactory.CreateGenericMessageAttachment(outsideMeetingErrorMsg);
                return Task.FromResult(this.GetMessengingExtensionResponse(outsideMeetingErrorCard, outsideMeetingErrorTitle, 100, 500));
            }

            switch (action.CommandId)
            {
                case "attachToTicket":
                    var attachmentUris = action.MessagePayload.Attachments.Select(attachment => attachment.ContentUrl);
                    string attachmentStr = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(attachmentUris)));

                    string localizedTitle;
                    using (new CultureSwitcher(locale, locale))
                    {
                        localizedTitle = this.localizer.GetString("TaskModuleTitleAttach");
                    }

                    var taskModuleUrl = $"https://{this.azureADOptions.Value.HostDomain}/consult/attach/{conversationId}/{attachmentStr}";
                    return Task.FromResult(this.GetMessengingExtensionResponse(taskModuleUrl, localizedTitle, 650, 600));
                default:
                    this.logger.LogError("Failed to fetch messaging extension task module. Invoke command {CommandId} is not valid", action.CommandId);
                    return default;
            }
        }

        /// <summary>
        /// Invoked when task module fetch event is received from the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            if (taskModuleRequest.Data == null)
            {
                this.logger.LogError("Failed to fetch task module. The data payload was null.");
                return default;
            }

            // parse the data passed in the task/fetch and perform response based on the command
            var payload = JsonConvert.DeserializeObject<TaskModuleRequestData>(JObject.Parse(taskModuleRequest.Data.ToString()).ToString());
            if (string.IsNullOrWhiteSpace(payload.ContextId))
            {
                this.logger.LogError("Failed to fetch task module. The consult ID in the payload was null or empty.");
                return default;
            }

            // Temporarily set agent's locale as current culture before using localizer
            var locale = turnContext.Activity.GetLocale();
            using (new CultureSwitcher(locale, locale))
            {
                switch (payload.Command)
                {
                    case "assignMe":
                        return this.GetTaskModuleResponse($"https://{this.azureADOptions.Value.HostDomain}/consult/assign/{payload.ContextId}/self", this.localizer.GetString("TaskModuleTitleAssignSelf"), 600, 600);
                    case "fromAssignedCard":
                        if (payload.StatusChangeChoice == "reassign")
                        {
                            return this.GetTaskModuleResponse($"https://{this.azureADOptions.Value.HostDomain}/consult/reassign/{payload.ContextId}", this.localizer.GetString("TaskModuleTitleReassign"), 300, 600);
                        }
                        else
                        {
                            string markCompleteMessage;
                            try
                            {
                                var request = await this.MarkCompletedAsync(payload.ContextId, turnContext);
                                markCompleteMessage = this.localizer.GetString("MessageConsultCompleteSuccess", request.FriendlyId);
                            }
                            catch (InvalidOperationException)
                            {
                                markCompleteMessage = this.localizer.GetString("MessageConsultCompleteFail");
                            }

                            var markCompleteResultCard = CardFactory.CreateGenericMessageAttachment(markCompleteMessage);
                            return this.GetTaskModuleResponse(markCompleteResultCard, this.localizer.GetString("TaskModuleTitleCompleteConsult"), 100, 400);
                        }

                    case "assignOther":
                        return this.GetTaskModuleResponse($"https://{this.azureADOptions.Value.HostDomain}/consult/assign/{payload.ContextId}/other", this.localizer.GetString("TaskModuleTitleAssignOther"), 600, 600);
                    case "details":
                        return this.GetTaskModuleResponse($"https://{this.azureADOptions.Value.HostDomain}/consult/detail/{payload.ContextId}", this.localizer.GetString("TaskModuleTitleDetails"), 550, 600);
                    case "addNotes":
                        string addNotesMessage;
                        if (string.IsNullOrEmpty(payload.Notes))
                        {
                            addNotesMessage = this.localizer.GetString("MessageNotesEmptyError");
                        }
                        else
                        {
                            try
                            {
                                await this.AddNotesAsync(payload.ContextId, payload.Notes, turnContext);
                                addNotesMessage = this.localizer.GetString("MessageNotesAddSuccess");
                            }
                            catch (InvalidOperationException)
                            {
                                addNotesMessage = this.localizer.GetString("MessageNotesAddFail");
                            }
                        }

                        var addNotesResultCard = CardFactory.CreateGenericMessageAttachment(addNotesMessage);
                        return this.GetTaskModuleResponse(addNotesResultCard, this.localizer.GetString("TaskModuleTitleAddNotes"), 100, 400);
                    default:
                        this.logger.LogError("Task module invoke command {CommandId} is not valid", payload.Command);
                        return default;
                }
            }
        }

        /// <summary>
        /// Invoked when task module submit event is received from the user.
        /// </summary>
        /// <inheritdoc/>
        protected override Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            // In this app, task module results are only used by tabs to refresh tab content.
            // The bot does not use the task module result, so it ignores these events.
            return Task.FromResult<TaskModuleResponse>(null);
        }

        /// <summary>
        /// Handle when a member is added to channel or 1:1 message, including the bot
        /// </summary>
        /// <param name="membersAdded">Collection of ChannelAccounts added.</param>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A Task on completion of event.</returns>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            if (membersAdded == null || membersAdded.Count == 0)
            {
                return;
            }

            if (turnContext == null)
            {
                return;
            }

            // Try loading agent from DB
            var loadedAgent = await this.agentRepository.GetByObjectIdAsync(turnContext.Activity.From.AadObjectId);

            // Check the context for members added
            var conversationType = turnContext.Activity.Conversation.ConversationType.ToEnum<ConversationType>();
            if (conversationType == ConversationType.Personal)
            {
                Microsoft.Bot.Schema.Attachment cardAttachment;

                // Temporarily set agent's locale as current culture before using localizer
                using (new CultureSwitcher(loadedAgent?.Locale, loadedAgent?.Locale))
                {
                    var myConsultsDeepLink = DeepLinkUtil.GetStaticTabDeepLink(this.teamsOptions.Value.AgentAppId, TabEntityId.My, this.localizer["MyConsultsDeepLinkLabel"], $"https://{this.azureADOptions.Value.HostDomain}/consult/my");
                    cardAttachment = CardFactory.CreateWelcomePersonalAttachment(myConsultsDeepLink, this.localizer);
                }

                // App was just installed for a user...send personal welcome message
                await ProactiveUtil.SendChatProactiveMessageAsync(
                    cardAttachment,
                    turnContext.Activity.From.Id,
                    turnContext.Activity.Conversation.TenantId,
                    turnContext.Activity.ServiceUrl,
                    this.botOptions.Value);
            }

            if (conversationType == ConversationType.Channel)
            {
                // Check for app being installed into Team or if member was added on existing install
                if (membersAdded[0].Id == turnContext.Activity.Recipient.Id)
                {
                    // Bot was installed into a new Team...capture details and store in database
                    var details = await TeamsInfo.GetTeamDetailsAsync(turnContext, turnContext.Activity.TeamsGetTeamInfo().Id, cancellationToken);
                    var channels = await TeamsInfo.GetTeamChannelsAsync(turnContext, turnContext.Activity.TeamsGetTeamInfo().Id, cancellationToken);
                    var members = await TeamsInfo.GetMembersAsync(turnContext, cancellationToken);

                    // Process team channels
                    var existingItems = await this.channelRepository.GetByTeamIdAsync(details.Id);
                    foreach (var channel in channels)
                    {
                        // check if channel exists
                        if (existingItems.FirstOrDefault(i => i.ChannelId == channel.Id) == null)
                        {
                            var newChannel = new Channel()
                            {
                                Id = Guid.NewGuid(),
                                TenantId = turnContext.Activity.Conversation.TenantId,
                                ServiceUrl = turnContext.Activity.ServiceUrl,
                                TeamId = details.Id,
                                TeamAADObjectId = details.AadGroupId,
                                TeamName = details.Name,
                                ChannelId = channel.Id,
                                ChannelName = string.IsNullOrEmpty(channel.Name) ? "General" : channel.Name,
                                CreatedDateTime = DateTime.Now,
                            };

                            // Upsert the channel (could be re-install)
                            await this.channelRepository.AddAsync(newChannel);
                        }
                    }

                    // Process members
                    foreach (var member in members)
                    {
                        // check if agent exists
                        var existingAgent = await this.agentRepository.GetByObjectIdAsync(member.AadObjectId);
                        if (existingAgent == null)
                        {
                            var newAgent = new Agent()
                            {
                                Id = Guid.NewGuid(),
                                TeamsId = member.Id,
                                AADObjectId = member.AadObjectId,
                                UserPrincipalName = member.UserPrincipalName,
                                Name = member.Name,
                                ServiceUrl = turnContext.Activity.ServiceUrl,
                                CreatedDateTime = DateTime.Now,
                            };

                            // Upsert agent
                            await this.agentRepository.UpsertAsync(newAgent);
                        }
                    }

                    // Send proactive welcome message
                    var myConsultsDeepLink = DeepLinkUtil.GetStaticTabDeepLink(this.teamsOptions.Value.AgentAppId, TabEntityId.My, this.localizer["MyConsultsDeepLinkLabel"], $"https://{this.azureADOptions.Value.HostDomain}/consult/my");
                    var cardAttachment = CardFactory.CreateWelcomeTeamAttachment(myConsultsDeepLink, this.localizer);
                    await ProactiveUtil.SendChannelProactiveMessageAsync(
                        cardAttachment,
                        turnContext.Activity.Conversation.Id,
                        turnContext.Activity.ServiceUrl,
                        this.botOptions.Value);
                }
                else
                {
                    // Upsert the agents added
                    foreach (var member in membersAdded)
                    {
                        var existingAgent = await this.agentRepository.GetByObjectIdAsync(member.AadObjectId);
                        if (existingAgent == null)
                        {
                            var details = await TeamsInfo.GetMemberAsync(turnContext, member.Id, cancellationToken);
                            var newAgent = new Agent()
                            {
                                Id = Guid.NewGuid(),
                                TeamsId = member.Id,
                                AADObjectId = member.AadObjectId,
                                UserPrincipalName = details.UserPrincipalName,
                                Name = member.Name,
                                ServiceUrl = turnContext.Activity.ServiceUrl,
                                CreatedDateTime = DateTime.Now,
                            };

                            // Upsert agent
                            await this.agentRepository.UpsertAsync(newAgent);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Handles channels being added to team
        /// </summary>
        /// <param name="channelInfo">Information on the channel being added.</param>
        /// <param name="teamInfo">Information on the team the channel is being added to.</param>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A Task on completion of event.</returns>
        protected override async Task OnTeamsChannelCreatedAsync(Microsoft.Bot.Schema.Teams.ChannelInfo channelInfo, Microsoft.Bot.Schema.Teams.TeamInfo teamInfo, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            // Add the channel to database so we can route messages to it
            var details = await TeamsInfo.GetTeamDetailsAsync(turnContext, turnContext.Activity.TeamsGetTeamInfo().Id, cancellationToken);
            var newChannel = new Channel()
            {
                Id = Guid.NewGuid(),
                TenantId = turnContext.Activity.Conversation.TenantId,
                ServiceUrl = turnContext.Activity.ServiceUrl,
                TeamId = details.Id,
                TeamAADObjectId = details.AadGroupId,
                TeamName = details.Name,
                ChannelId = channelInfo.Id,
                ChannelName = string.IsNullOrEmpty(channelInfo.Name) ? "General" : channelInfo.Name,
                CreatedDateTime = DateTime.Now,
            };

            // Insert the channel
            await this.channelRepository.AddAsync(newChannel);
        }

        /// <summary>
        /// Handles channels being deleted from a team
        /// </summary>
        /// <param name="channelInfo">Information on the channel being deleted.</param>
        /// <param name="teamInfo">Information on the team the channel is being delete from.</param>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A Task on completion of event.</returns>
        protected override async Task OnTeamsChannelDeletedAsync(Microsoft.Bot.Schema.Teams.ChannelInfo channelInfo, Microsoft.Bot.Schema.Teams.TeamInfo teamInfo, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            // Remove the channel from the existing record in the database...WARNING: might orphan a consumer option
            var channelToDelete = await this.channelRepository.GetByChannelIdAsync(channelInfo.Id);
            if (channelToDelete != null)
            {
                await this.channelRepository.DeleteAsync(new CosmosItemKey(channelToDelete.Id.ToString(), channelToDelete.ChannelId));
            }

            // Delete the channel mapping related to this channel (if exists)
            var mappings = await this.channelMappingRepository.GetByChannelIds(new[] { channelInfo.Id });
            foreach (var mapping in mappings)
            {
                await this.channelMappingRepository.DeleteAsync(new CosmosItemKey(mapping.Id.ToString(), mapping.Id.ToString()));
            }
        }

        /// <summary>
        /// Handles when a channel is renamed
        /// </summary>
        /// <param name="channelInfo">Information on the channel being renamed.</param>
        /// <param name="teamInfo">Information on the team the channel is being renamed in.</param>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A Task on completion of event.</returns>
        protected override async Task OnTeamsChannelRenamedAsync(Microsoft.Bot.Schema.Teams.ChannelInfo channelInfo, Microsoft.Bot.Schema.Teams.TeamInfo teamInfo, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            // Update the name of the channel in the database
            var channel = await this.channelRepository.GetByChannelIdAsync(channelInfo.Id);
            if (channel != null)
            {
                // update the channel
                channel.ChannelName = channelInfo.Name;
                await this.channelRepository.UpsertAsync(channel);
            }
            else
            {
                // create the channel
                var details = await TeamsInfo.GetTeamDetailsAsync(turnContext, turnContext.Activity.TeamsGetTeamInfo().Id, cancellationToken);
                var newChannel = new Channel()
                {
                    Id = Guid.NewGuid(),
                    TenantId = turnContext.Activity.Conversation.TenantId,
                    ServiceUrl = turnContext.Activity.ServiceUrl,
                    TeamId = details.Id,
                    TeamAADObjectId = details.AadGroupId,
                    TeamName = details.Name,
                    ChannelId = channelInfo.Id,
                    ChannelName = string.IsNullOrEmpty(channelInfo.Name) ? "General" : channelInfo.Name,
                    CreatedDateTime = DateTime.Now,
                };

                // perform the insert
                await this.channelRepository.AddAsync(newChannel);
            }
        }

        /// <summary>
        /// Handles when a soft-deleted channel is restored.
        /// </summary>
        /// <param name="channelInfo">Information on the channel being restored.</param>
        /// <param name="teamInfo">Information on the team the channel is being restored in.</param>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A Task on completion of event.</returns>
        protected override async Task OnTeamsChannelRestoredAsync(ChannelInfo channelInfo, TeamInfo teamInfo, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            await this.OnTeamsChannelCreatedAsync(channelInfo, teamInfo, turnContext, cancellationToken);
        }

        /// <summary>
        /// Handles when a team is renamed
        /// </summary>
        /// <param name="teamInfo">Information on the team being renamed.</param>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A Task on completion of event.</returns>
        protected override async Task OnTeamsTeamRenamedAsync(Microsoft.Bot.Schema.Teams.TeamInfo teamInfo, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            // Update the name of the channel in the database
            var channels = await this.channelRepository.GetByTeamIdAsync(teamInfo.Id);
            foreach (var channel in channels)
            {
                channel.TeamName = teamInfo.Name;
                await this.channelRepository.UpsertAsync(channel);
            }
        }

        /// <summary>
        /// Handles members being removed from team.
        /// </summary>
        /// <param name="teamsMembersRemoved">Collection of members removed.</param>
        /// <param name="teamInfo">Info on team members removed from.</param>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A Task on completion of event.</returns>
        protected override async Task OnTeamsMembersRemovedAsync(IList<Microsoft.Bot.Schema.Teams.TeamsChannelAccount> teamsMembersRemoved, Microsoft.Bot.Schema.Teams.TeamInfo teamInfo, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            if (teamsMembersRemoved == null || teamsMembersRemoved.Count == 0 || turnContext == null)
            {
                return;
            }

            // Check if the app was removed from the team
            var conversationType = turnContext.Activity.Conversation.ConversationType.ToEnum<ConversationType>();
            if (conversationType == ConversationType.Channel && teamsMembersRemoved.Any(member => member.Id == turnContext.Activity.Recipient.Id))
            {
                // Delete all the channel records for this team
                var channels = await this.channelRepository.GetByTeamIdAsync(teamInfo.Id);
                foreach (var channel in channels)
                {
                    await this.channelRepository.DeleteAsync(new CosmosItemKey(channel.Id.ToString(), channel.ChannelId));
                }

                // Delete the channel mappings related to this team's channels
                var mappings = await this.channelMappingRepository.GetByChannelIds(channels.Select(channel => channel.ChannelId));
                foreach (var mapping in mappings)
                {
                    await this.channelMappingRepository.DeleteAsync(new CosmosItemKey(mapping.Id.ToString(), mapping.Id.ToString()));
                }
            }
        }

        /// <summary>
        /// Get task module response object for a task module that opens a URL.
        /// </summary>
        private TaskModuleResponse GetTaskModuleResponse(string url, string title, int height, int width)
        {
            return new TaskModuleResponse
            {
                Task = this.GetTaskModuleContinueResponse(url, title, height, width),
            };
        }

        /// <summary>
        /// Get task module response object for a task module that shows an Adaptive Card.
        /// </summary>
        private TaskModuleResponse GetTaskModuleResponse(Microsoft.Bot.Schema.Attachment card, string title, int height, int width)
        {
            return new TaskModuleResponse
            {
                Task = this.GetTaskModuleContinueResponse(card, title, height, width),
            };
        }

        /// <summary>
        /// Get messaging extension response object for a task module that opens a URL.
        /// </summary>
        private MessagingExtensionActionResponse GetMessengingExtensionResponse(string url, string title, int height, int width)
        {
            return new MessagingExtensionActionResponse
            {
                Task = this.GetTaskModuleContinueResponse(url, title, height, width),
            };
        }

        /// <summary>
        /// Get messaging extension response object for a task module that shows an Adaptive Card.
        /// </summary>
        private MessagingExtensionActionResponse GetMessengingExtensionResponse(Microsoft.Bot.Schema.Attachment card, string title, int height, int width)
        {
            return new MessagingExtensionActionResponse
            {
                Task = this.GetTaskModuleContinueResponse(card, title, height, width),
            };
        }

        private TaskModuleContinueResponse GetTaskModuleContinueResponse(string url, string title, int height, int width)
        {
            return new TaskModuleContinueResponse()
            {
                Type = "continue",
                Value = new TaskModuleTaskInfo()
                {
                    Url = url,
                    Height = height,
                    Width = width,
                    Title = title,
                },
            };
        }

        private TaskModuleContinueResponse GetTaskModuleContinueResponse(Microsoft.Bot.Schema.Attachment card, string title, int height, int width)
        {
            return new TaskModuleContinueResponse()
            {
                Type = "continue",
                Value = new TaskModuleTaskInfo()
                {
                    Card = card,
                    Height = height,
                    Width = width,
                    Title = title,
                },
            };
        }

        private string GetConversationIdFromMessageLink(string messageLink)
        {
            if (messageLink.Equals(string.Empty) || messageLink == null)
            {
                return string.Empty;
            }

            var startIndex = messageLink.IndexOf("meeting");
            var endIndex = messageLink.IndexOf("@thread.v2");

            if (startIndex < 0 || endIndex < 0 || endIndex < startIndex)
            {
                return string.Empty;
            }

            return messageLink.Substring(startIndex, endIndex - startIndex);
        }

        private async Task<Note> AddNotesAsync(string consultId, string notes, ITurnContext turnContext)
        {
            var userObjectId = turnContext.Activity.From.AadObjectId;
            var userName = turnContext.Activity.From.Name;
            if (string.IsNullOrWhiteSpace(userObjectId))
            {
                this.logger.LogError("Failed to add notes to the consult request. The incoming user object ID is missing from the activity.");
                throw new InvalidOperationException("The incoming user ID is invalid.");
            }

            var request = await this.requestRepository.GetAsync(new CosmosItemKey(consultId, consultId));
            if (request == null)
            {
                this.logger.LogError("Failed to add notes to the consult request. Consult request {ConsultId} does not exist.", consultId);
                throw new InvalidOperationException("The consult request does not exist.");
            }

            var currentTime = DateTime.UtcNow;
            var userObjectIdGuid = new Guid(userObjectId);

            var newNote = new Note
            {
                Id = Guid.NewGuid(),
                CreatedByName = userName,
                CreatedById = userObjectIdGuid,
                CreatedDateTime = currentTime,
                Text = notes,
            };

            request.Notes = request.Notes ?? new List<Note>();
            request.Notes.Add(newNote);

            await this.requestRepository.UpsertAsync(request);

            return newNote;
        }

        private async Task<Request> MarkCompletedAsync(string consultId, ITurnContext turnContext)
        {
            var userObjectId = turnContext.Activity.From.AadObjectId;
            var userName = turnContext.Activity.From.Name;
            if (string.IsNullOrWhiteSpace(userObjectId))
            {
                this.logger.LogError("Failed to complete consult request. The incoming user object ID is missing from the activity.");
                throw new InvalidOperationException("The incoming user ID is invalid.");
            }

            var request = await this.requestRepository.GetAsync(new CosmosItemKey(consultId, consultId));
            if (request == null)
            {
                this.logger.LogError("Failed to complete consult request. Consult request {ConsultId} does not exist.", consultId);
                throw new InvalidOperationException("The consult request does not exist.");
            }

            if (request.Status == RequestStatus.Completed)
            {
                throw new InvalidOperationException("The request is already completed.");
            }

            var currentTime = DateTime.UtcNow;
            var userObjectIdGuid = new Guid(userObjectId);

            // Change request status to completed
            request.Status = RequestStatus.Completed;
            request.Activities = request.Activities ?? new List<Microsoft.Teams.App.VirtualConsult.Common.Models.Activity>();
            request.Activities.Add(new Microsoft.Teams.App.VirtualConsult.Common.Models.Activity
            {
                Id = Guid.NewGuid(),
                Type = ActivityType.Completed,
                CreatedByName = userName,
                CreatedById = userObjectIdGuid,
                CreatedDateTime = currentTime,
            });

            // Update request in DB
            await this.requestRepository.UpsertAsync(request);
            return request;
        }
    }
}
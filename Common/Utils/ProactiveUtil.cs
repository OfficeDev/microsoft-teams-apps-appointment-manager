// <copyright file="ProactiveUtil.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Utils
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Bot.Connector;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration;

    /// <summary>
    /// Utility class for sending bot proactive messages.
    /// </summary>
    public static class ProactiveUtil
    {
        /// <summary>
        /// Sends a proactive message to a specific user.
        /// </summary>
        /// <param name="cardAttachment">The Bot Framework attachment to send.</param>
        /// <param name="userId">The teams user id for the user to message.</param>
        /// <param name="tenantId">Tenant Id where the message is going.</param>
        /// <param name="serviceUrl">ServiceUrl for the tenant.</param>
        /// <param name="botSettings">Bot configuration settings.</param>
        /// <returns>Async task.</returns>
        public static async Task SendChatProactiveMessageAsync(Attachment cardAttachment, string userId, string tenantId, string serviceUrl, BotSettings botSettings)
        {
            if (cardAttachment == null)
            {
                return;
            }

            var parameters = new ConversationParameters
            {
                Members = new[] { new ChannelAccount(userId) },
                ChannelData = new TeamsChannelData
                {
                    Tenant = new TenantInfo(tenantId),
                },
            };

            // Establish the conversation with proper access token
            MicrosoftAppCredentials.TrustServiceUrl(serviceUrl, DateTime.MaxValue);
            using (var connectorClient = new ConnectorClient(new Uri(serviceUrl), botSettings.Id, botSettings.Password))
            {
                var response = await connectorClient.Conversations.CreateConversationAsync(parameters);

                var message = Activity.CreateMessageActivity();
                message.Attachments.Add(cardAttachment);
                await connectorClient.Conversations.SendToConversationAsync(response.Id, (Activity)message);
            }
        }

        /// <summary>
        ///  Sends a proactive message to a specific Teams channel.
        /// </summary>
        /// <param name="cardAttachment">The Bot Framework attachment to send.</param>
        /// <param name="channelId">Teams channel id to send message to.</param>
        /// <param name="serviceUrl">ServiceUrl for the tenant.</param>
        /// <param name="botSettings">Bot configuration settings.</param>
        /// <returns>Async task.</returns>
        public static async Task<ConversationResourceResponse> SendChannelProactiveMessageAsync(Attachment cardAttachment, string channelId, string serviceUrl, BotSettings botSettings)
        {
            if (cardAttachment == null)
            {
                return null;
            }

            var message = Activity.CreateMessageActivity();
            message.Attachments.Add(cardAttachment);

            var conversationParameters = new ConversationParameters
            {
                IsGroup = true,
                ChannelData = new TeamsChannelData
                {
                    Channel = new ChannelInfo(channelId),
                },
                Activity = (Activity)message,
            };

            // Establish the conversation with proper access token
            MicrosoftAppCredentials.TrustServiceUrl(serviceUrl, DateTime.MaxValue);
            using (var connectorClient = new ConnectorClient(new Uri(serviceUrl), botSettings.Id, botSettings.Password))
            {
                var conversationResource = await connectorClient.Conversations.CreateConversationAsync(conversationParameters);

                // Returns conversation resource containing message and conversation Ids
                return conversationResource;
            }
        }

        /// <summary>
        ///  Update a specific message that was previously sent.
        /// </summary>
        /// <param name="cardAttachment">The Bot Framework attachment to send.</param>
        /// <param name="serviceUrl">ServiceUrl for the tenant.</param>
        /// <param name="conversationId">The conversation ID of the previously sent message.</param>
        /// <param name="activityId">The activity ID of the previously sent message.</param>
        /// <param name="botSettings">Bot configuration settings.</param>
        /// <returns>Async task.</returns>
        public static async Task<ResourceResponse> UpdateChannelProactiveMessageAsync(Attachment cardAttachment, string serviceUrl, string conversationId, string activityId, BotSettings botSettings)
        {
            if (cardAttachment == null)
            {
                return null;
            }

            var message = Activity.CreateMessageActivity();
            message.Attachments.Add(cardAttachment);

            // Establish the conversation with proper access token
            MicrosoftAppCredentials.TrustServiceUrl(serviceUrl, DateTime.MaxValue);
            using (var connectorClient = new ConnectorClient(new Uri(serviceUrl), botSettings.Id, botSettings.Password))
            {
                var resourceResponse = await connectorClient.Conversations.UpdateActivityAsync(conversationId, activityId, (Activity)message);

                // Returns resource response containing message ID
                return resourceResponse;
            }
        }

        /// <summary>
        ///  Send a reply to a specific message that was previously sent.
        /// </summary>
        /// <param name="text">The text to send.</param>
        /// <param name="serviceUrl">ServiceUrl for the tenant.</param>
        /// <param name="conversationId">The conversation ID of the previously sent message.</param>
        /// <param name="activityId">The activity ID of the previously sent message.</param>
        /// <param name="botSettings">Bot configuration settings.</param>
        /// <returns>Async task.</returns>
        public static async Task<ResourceResponse> ReplyToChannelMessageAsync(string text, string serviceUrl, string conversationId, string activityId, BotSettings botSettings)
        {
            var message = Activity.CreateMessageActivity();
            message.Text = text;

            MicrosoftAppCredentials.TrustServiceUrl(serviceUrl, DateTime.MaxValue);
            using (var connectorClient = new ConnectorClient(new Uri(serviceUrl), botSettings.Id, botSettings.Password))
            {
                var resourceResponse = await connectorClient.Conversations.ReplyToActivityAsync(conversationId, activityId, (Activity)message);

                // Returns resource response containing message ID
                return resourceResponse;
            }
        }
    }
}
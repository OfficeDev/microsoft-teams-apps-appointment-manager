// <copyright file="CardFactory.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Utils
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using AdaptiveCards;
    using AdaptiveCards.Templating;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.App.VirtualConsult.Common.Models;

    /// <summary>
    /// Helper class to generate Adaptive cards used throughout the solution.
    /// </summary>
    public static class CardFactory
    {
        /// <summary>
        /// Creates a Bot Framework <see cref="Bot.Schema.Attachment"/> containing the new consult adaptive card.
        /// </summary>
        /// <param name="consultRequest">The consult whose info will be shown in the card.</param>
        /// <param name="baseUrl">The base URL of the application. Used to reference hosted images.</param>
        /// <param name="localizer">The string localizer from which to get static strings.</param>
        /// <returns>An <see cref="Bot.Schema.Attachment"/> containing the new consult adaptive card.</returns>
        public static Bot.Schema.Attachment CreateConsultAttachment(Request consultRequest, string baseUrl, IStringLocalizer localizer)
        {
            _ = consultRequest ?? throw new ArgumentNullException(nameof(consultRequest));
            _ = baseUrl ?? throw new ArgumentNullException(nameof(baseUrl));

            var staticStrings = GetLocalizedStringsAsDict(localizer);
            var cardDataObject = new { consultRequest, baseUrl, staticStrings };
            var adaptiveCard = GetAdaptiveCard("NewRequest.json", cardDataObject);
            return new Bot.Schema.Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard,
            };
        }

        /// <summary>
        /// Creates a Bot Framework <see cref="Bot.Schema.Attachment"/> containing the reassign consult adaptive card.
        /// </summary>
        /// <param name="consultRequest">The consult whose info will be shown in the card.</param>
        /// <param name="agents">The agents that are being asked to pick up the request.</param>
        /// <param name="baseUrl">The base URL of the application. Used to reference hosted images.</param>
        /// <param name="comment">The comment placed by the agent who is requesting a reassign.</param>
        /// <param name="initiaterDisplayName">The name of the agent who requested the reassignment.</param>
        /// <param name="initiaterPhotoUrl">Data URL with Base64 string representing the initiater's photo.</param>
        /// <param name="localizer">The string localizer from which to get static strings.</param>
        /// <returns>An <see cref="Bot.Schema.Attachment"/> containing the new consult adaptive card.</returns>
        public static Bot.Schema.Attachment CreateReassignConsultAttachment(Request consultRequest, object agents, string baseUrl, string comment, string initiaterDisplayName, string initiaterPhotoUrl, IStringLocalizer localizer)
        {
            if (consultRequest == null)
            {
                throw new ArgumentNullException(nameof(consultRequest));
            }

            if (string.IsNullOrEmpty(baseUrl))
            {
                throw new ArgumentNullException(nameof(baseUrl));
            }

            var staticStrings = GetLocalizedStringsAsDict(localizer);
            var cardDataObject = new { consultRequest, baseUrl, agents, comment, initiaterDisplayName, initiaterPhotoUrl = initiaterPhotoUrl ?? string.Empty, staticStrings };
            var adaptiveCard = GetAdaptiveCard("RequestReassign.json", cardDataObject);
            return new Bot.Schema.Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard,
            };
        }

        /// <summary>
        /// Creates a Bot Framework <see cref="Bot.Schema.Attachment"/> containing the assigned consult adaptive card.
        /// </summary>
        /// <param name="consultRequest">The consult whose info will be shown in the card.</param>
        /// <param name="baseUrl">The base URL of the application. Used to reference hosted images.</param>
        /// <param name="assignedToName">The name of the agent to whom the consult is assigned.</param>
        /// <param name="isPersonalCard">Indicates whether the card will be sent in personal (1:1) chat or in a channel.</param>
        /// <param name="localizer">The string localizer from which to get static strings.</param>
        /// <returns>An <see cref="Bot.Schema.Attachment"/> containing the assigned consult adaptive card.</returns>
        public static Bot.Schema.Attachment CreateAssignedConsultAttachment(Request consultRequest, string baseUrl, string assignedToName, bool isPersonalCard, IStringLocalizer localizer)
        {
            _ = consultRequest ?? throw new ArgumentNullException(nameof(consultRequest));
            _ = baseUrl ?? throw new ArgumentNullException(nameof(baseUrl));
            _ = assignedToName ?? throw new ArgumentNullException(nameof(assignedToName));

            if (string.IsNullOrWhiteSpace(baseUrl))
            {
                throw new ArgumentException("Argument cannot be whitespace.", nameof(baseUrl));
            }

            if (string.IsNullOrWhiteSpace(assignedToName))
            {
                throw new ArgumentException("Argument cannot be whitespace.", nameof(assignedToName));
            }

            var staticStrings = GetLocalizedStringsAsDict(localizer);
            var cardDataObject = new { consultRequest, baseUrl, assignedToName, isPersonalCard, staticStrings };
            var adaptiveCard = GetAdaptiveCard("AssignedConsult.json", cardDataObject);
            return new Bot.Schema.Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard,
            };
        }

        /// <summary>
        /// Creates a Bot Framework <see cref="Bot.Schema.Attachment"/> containing the personal welcome adaptive card.
        /// </summary>
        /// <param name="deepLink">The deep link URL to the app's My Consults tab.</param>
        /// <param name="localizer">The string localizer from which to get static strings.</param>
        /// <returns>An <see cref="Bot.Schema.Attachment"/> containing the personal welcome adaptive card.</returns>
        public static Bot.Schema.Attachment CreateWelcomePersonalAttachment(string deepLink, IStringLocalizer localizer)
        {
            _ = deepLink ?? throw new ArgumentNullException(nameof(deepLink));

            if (string.IsNullOrWhiteSpace(deepLink))
            {
                throw new ArgumentException("Argument cannot be whitespace.", nameof(deepLink));
            }

            var staticStrings = GetLocalizedStringsAsDict(localizer);
            var cardDataObject = new { deepLink, staticStrings };
            var adaptiveCard = GetAdaptiveCard("WelcomeMessagePersonal.json", cardDataObject);
            return new Bot.Schema.Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard,
            };
        }

        /// <summary>
        /// Creates a Bot Framework <see cref="Bot.Schema.Attachment"/> containing the team welcome adaptive card.
        /// </summary>
        /// <param name="deepLink">The deep link URL to the app's My Consults tab.</param>
        /// <param name="localizer">The string localizer from which to get static strings.</param>
        /// <returns>An <see cref="Bot.Schema.Attachment"/> containing the team welcome adaptive card.</returns>
        public static Bot.Schema.Attachment CreateWelcomeTeamAttachment(string deepLink, IStringLocalizer localizer)
        {
            _ = deepLink ?? throw new ArgumentNullException(nameof(deepLink));

            if (string.IsNullOrWhiteSpace(deepLink))
            {
                throw new ArgumentException("Argument cannot be whitespace.", nameof(deepLink));
            }

            var staticStrings = GetLocalizedStringsAsDict(localizer);
            var cardDataObject = new { deepLink, staticStrings };
            var adaptiveCard = GetAdaptiveCard("WelcomeMessageTeam.json", cardDataObject);
            return new Bot.Schema.Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard,
            };
        }

        /// <summary>
        /// Creates an adaptive card for a generic success/failure task module.
        /// </summary>
        /// <param name="message">The message to show in the adaptive card.</param>
        /// <returns>An <see cref="Bot.Schema.Attachment"/> containing the adaptive card.</returns>
        public static Bot.Schema.Attachment CreateGenericMessageAttachment(string message)
        {
            _ = message ?? throw new ArgumentNullException(nameof(message));

            string templateName = "GenericMessage.json";
            var cardDataObject = new { message };
            var adaptiveCard = GetAdaptiveCard(templateName, cardDataObject);
            return new Bot.Schema.Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard,
            };
        }

        /// <summary>
        /// Creates an adaptive card for the given template and binds multiple object data to it.
        /// </summary>
        /// <param name="templateName">The file name of the adaptive card template.</param>
        /// <param name="dataObject">The object containing data to bind to the card. Optional.</param>
        /// <returns>An <see cref="AdaptiveCard"/> created from the template and data objects.</returns>
        private static AdaptiveCard GetAdaptiveCard(string templateName, object dataObject = null)
        {
            var cardPath = Path.Combine("Resources", "Cards", templateName);
            var consultCardJson = File.ReadAllText(cardPath, Encoding.UTF8);

            // Expand card using data object (if provided)
            if (dataObject != null)
            {
                var adaptiveCardTemplate = new AdaptiveCardTemplate(consultCardJson);
                consultCardJson = adaptiveCardTemplate.Expand(dataObject);
            }

            return AdaptiveCard.FromJson(consultCardJson).Card;
        }

        /// <summary>
        /// Returns a dictionary of all localized string key/values from the given localizer.
        /// </summary>
        /// <param name="localizer">The string localizer from which to get strings.</param>
        /// <returns>A <see cref="Dictionary{TKey, TValue}"/> containing localized string keys mapped to their string values.</returns>
        private static Dictionary<string, string> GetLocalizedStringsAsDict(IStringLocalizer localizer)
        {
            return localizer
                .GetAllStrings()
                .ToDictionary(ls => ls.Name, ls => ls.Value);
        }
    }
}

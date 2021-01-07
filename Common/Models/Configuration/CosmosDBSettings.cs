// <copyright file="CosmosDBSettings.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models.Configuration
{
    /// <summary>
    /// Class CosmosDBSettings.
    /// </summary>
    public class CosmosDBSettings
    {
        /// <summary>
        /// Gets or sets the Cosmos DB Connection string.
        /// </summary>
        /// <value>The endpoint URL.</value>
        public string ConnectionString { get; set; }

        /// <summary>
        /// Gets or sets the name of the database.
        /// </summary>
        /// <value>The name of the database.</value>
        public string DatabaseName { get; set; }
    }
}

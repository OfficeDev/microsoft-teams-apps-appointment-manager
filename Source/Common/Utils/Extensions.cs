// <copyright file="Extensions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Models
{
    using System;

    /// <summary>
    /// Extensions class for project.
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// Converts string to enumeration of type T.
        /// </summary>
        /// <typeparam name="T">A enumeration type.</typeparam>
        /// <param name="value">The string value to convert.</param>
        /// <returns>The enum value of type T.</returns>
        public static T ToEnum<T>(this string value)
        {
            return (T)Enum.Parse(typeof(T), value, true);
        }
    }
}
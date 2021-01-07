// <copyright file="CultureSwitcher.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Common.Utils
{
    using System;
    using System.Globalization;

    /// <summary>
    /// Utility class for temporarily switching the current culture. The previous culture is restored when the instance is disposed.
    /// </summary>
    /// <remarks>
    /// This class is useful when using <see cref="Microsoft.Extensions.Localization.IStringLocalizer"/> for a culture other than the current culture.
    /// The <see cref="Microsoft.Extensions.Localization.IStringLocalizer"/> uses the current culture to determine which resource file to load.
    /// In order to load a different resource file, the current culture must be changed.
    /// See <see href="https://github.com/dotnet/aspnetcore/issues/7756">dotnet/aspnetcore issue 7756</see> for more details.
    /// </remarks>
    public sealed class CultureSwitcher : IDisposable
    {
        private readonly CultureInfo originalCulture;
        private readonly CultureInfo originalUICulture;

        /// <summary>
        /// Initializes a new instance of the <see cref="CultureSwitcher"/> class.
        /// </summary>
        /// <param name="culture">The culture to switch to. An invalid culture (including null) will result in no culture change.</param>
        /// <param name="uiCulture">The UI culture to switch to. An invalid culture (including null) will result in no UI culture change.</param>
        public CultureSwitcher(string culture, string uiCulture)
        {
            this.originalCulture = CultureInfo.CurrentCulture;
            this.originalUICulture = CultureInfo.CurrentUICulture;

            var cultureInfo = this.TryLoadCulture(culture) ?? this.originalCulture;
            var uiCultureInfo = this.TryLoadCulture(uiCulture) ?? this.originalUICulture;

            this.SetCulture(cultureInfo, uiCultureInfo);
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            this.SetCulture(this.originalCulture, this.originalUICulture);
        }

        private CultureInfo TryLoadCulture(string culture)
        {
            if (string.IsNullOrWhiteSpace(culture))
            {
                return null;
            }

            try
            {
                return new CultureInfo(culture);
            }
            catch (CultureNotFoundException)
            {
                return null;
            }
        }

        private void SetCulture(CultureInfo culture, CultureInfo uiCulture)
        {
            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = uiCulture;
        }
    }
}

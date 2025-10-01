// Namespace for the entire Birthday Extractor application.
namespace BirthdayExtractor
{
    // Imports System for the Version class.
    using System;

    /// <summary>
    /// Centralizes the application's version information.
    /// This allows for easy updates and consistent version reporting across the application,
    /// including for update checks against remote releases.
    /// </summary>
    internal static class AppVersion
    {
        /// <summary>
        /// Gets the human-friendly version string to be displayed in the UI.
        /// This should be updated for each new release.
        /// </summary>
        public const string Display = "0.86";

        /// <summary>
        /// Gets the semantic <see cref="Version"/> representation of the display string.
        /// This is used for reliable version comparison, for example, when checking for updates.
        /// </summary>
        public static Version Semantic { get; } = Version.Parse(Display);
    }
}
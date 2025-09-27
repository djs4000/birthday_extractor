using System;

namespace BirthdayExtractor
{
    /// <summary>
    /// Centralizes the application version that should be displayed and
    /// compared against remote releases.
    /// </summary>
    internal static class AppVersion
    {
        /// <summary>
        /// Gets the human friendly version string.
        /// </summary>
        public const string Display = "0.55";

        /// <summary>
        /// Gets the semantic <see cref="Version"/> representation used when
        /// comparing with GitHub releases.
        /// </summary>
        public static Version Semantic { get; } = Version.Parse(Display);
    }
}

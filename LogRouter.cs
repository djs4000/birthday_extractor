using System;
using System.Collections.Generic;

namespace BirthdayExtractor
{
    /// <summary>
    /// Centralized helper that forwards log messages to the UI text box when available.
    /// Allows background services and non-UI classes to surface exceptions to the user.
    /// </summary>
    internal static class LogRouter
    {
        private static readonly object _sync = new();
        private static readonly Queue<string> _pending = new();
        private static Action<string>? _uiLog;
        private static bool _verboseEnabled = true;

        /// <summary>
        /// Registers the UI log sink so background components can forward messages to the text box.
        /// Any pending messages queued before registration will be flushed immediately.
        /// </summary>
        public static void RegisterUiLogger(Action<string> log)
        {
            if (log is null)
            {
                throw new ArgumentNullException(nameof(log));
            }

            List<string>? flush = null;
            var flushMessages = false;
            lock (_sync)
            {
                _uiLog = log;
                if (!_verboseEnabled)
                {
                    _pending.Clear();
                }
                else if (_pending.Count > 0)
                {
                    flush = new List<string>(_pending);
                    _pending.Clear();
                    flushMessages = true;
                }
            }

            if (flushMessages && flush is not null)
            {
                foreach (var message in flush)
                {
                    log(message);
                }
            }
        }

        /// <summary>
        /// Unregisters the UI logger so that disposed forms do not retain delegates.
        /// </summary>
        public static void UnregisterUiLogger(Action<string> log)
        {
            if (log is null)
            {
                throw new ArgumentNullException(nameof(log));
            }

            lock (_sync)
            {
                if (_uiLog == log)
                {
                    _uiLog = null;
                }
            }
        }

        /// <summary>
        /// Routes a simple log message to the UI, or stores it until the UI is ready.
        /// </summary>
        public static void LogMessage(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            Action<string>? logger;
            lock (_sync)
            {
                if (!_verboseEnabled)
                {
                    return;
                }

                logger = _uiLog;
                if (logger is null)
                {
                    _pending.Enqueue(message);
                    return;
                }
            }

            logger(message);
        }

        /// <summary>
        /// Determines whether the supplied delegate is the currently registered UI logger.
        /// </summary>
        public static bool IsRegisteredLogger(Action<string> log)
        {
            if (log is null)
            {
                return false;
            }

            lock (_sync)
            {
                return _uiLog == log;
            }
        }

        /// <summary>
        /// Formats and logs an exception message, preserving contextual prefixes when supplied.
        /// </summary>
        public static void LogException(Exception ex, string? context = null)
        {
            if (ex is null)
            {
                return;
            }

            var trimmedContext = string.IsNullOrWhiteSpace(context) ? null : context!.TrimEnd();
            var formatted = trimmedContext is null
                ? ex.Message
                : $"{trimmedContext}{(trimmedContext.EndsWith(':') ? string.Empty : ":")} {ex.Message}";
            LogMessage(formatted);
        }

        /// <summary>
        /// Enables or disables forwarding of log messages to the UI text box.
        /// When disabled, queued messages are discarded.
        /// </summary>
        public static void SetVerboseLoggingEnabled(bool enabled)
        {
            lock (_sync)
            {
                _verboseEnabled = enabled;
                if (!enabled)
                {
                    _pending.Clear();
                }
            }
        }
    }
}


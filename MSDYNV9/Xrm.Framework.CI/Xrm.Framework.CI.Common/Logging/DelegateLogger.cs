using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Xrm.Framework.CI.Common.Logging
{
    /// <summary>
    /// Provides a way to convert methods like  Action<string> logVerbose to ILogger
    /// </summary>
    public class DelegateLogger : ILogger
    {
        public Action<string> LogError { get; }
        public Action<string> LogWarning { get; }
        public Action<string> LogInformation { get; }
        public Action<string> LogVerbose { get; }

        /// <summary>
        /// You can pass nulls if you want - logging methods will be callable but will do nothing.
        /// </summary>
        /// <param name="logError"></param>
        /// <param name="logWarning"></param>
        /// <param name="logInformation"></param>
        /// <param name="logVerbose"></param>
        public DelegateLogger(Action<string> logError, Action<string> logWarning, Action<string> logInformation, Action<string> logVerbose)
        {
            LogError = logError ?? new Action<string>((msg) => { return; });
            LogWarning = logWarning ?? new Action<string>((msg) => { return; });
            LogInformation = logInformation ?? new Action<string>((msg) => { return; });
            LogVerbose = logVerbose ?? new Action<string>((msg) => { return; });
        }

        void ILogger.LogVerbose(string format, params object[] args) => LogVerbose(string.Format(format, args));
        void ILogger.LogInformation(string format, params object[] args) => LogInformation(string.Format(format, args));
        void ILogger.LogWarning(string format, params object[] args) => LogWarning(string.Format(format, args));
        void ILogger.LogError(string format, params object[] args) => LogError(string.Format(format, args));
    }
}

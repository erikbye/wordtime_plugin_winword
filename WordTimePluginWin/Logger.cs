using System;
using System.IO;
using WordTimePluginWin.Extensions;

namespace WordTimePluginWin {
    internal static class Logger {
        private static readonly StreamWriter _streamWriter = new StreamWriter("wordtime.log") {
            AutoFlush = true
        };

        public static void Log(string message) {
#if DEBUG
            _streamWriter.WriteLine(DateTime.Now.GetTimestamp() + " : " + message);
#endif
        }
    }
}
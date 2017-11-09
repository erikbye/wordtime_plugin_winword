using System;
using System.IO;
using WordTimePluginWin.Extensions;

namespace WordTimePluginWin
{
    internal static class Logger
    {
        private static readonly StreamWriter _streamWriter = new StreamWriter(Config.Homepath + @"\wordtime.log")
        {
            AutoFlush = true
        };

        public static void Log(string message)
        {        
            _streamWriter.WriteLine(message + ": " + DateTime.Now.GetTimestamp());
        }
    }
}
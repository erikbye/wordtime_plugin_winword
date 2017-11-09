using System;
using System.IO;

namespace WordTimePluginWin
{
    class Logger
    {
        private static readonly string _userHomePath = Environment.GetEnvironmentVariable("homepath");
        private readonly StreamWriter _streamWriter;
        public Logger()
        {
            _streamWriter = new StreamWriter(_userHomePath + @"\wordtime.log", true);
        }

        public void Log(string message)
        {        
            _streamWriter.WriteLine(message + ": " + DateTime.Now.GetTimestamp());
            _streamWriter.Flush();
        }
    }
}
using System;
using System.Configuration;
using System.IO;
using WordTimePluginWin.Extensions;

namespace WordTimePluginWin
{
    internal class Logger
    {
        private readonly StreamWriter _streamWriter;

        public Logger()
        {
            _streamWriter = new StreamWriter(Config.Homepath + @"\wordtime.log", true);
        }

        public void Log(string message)
        {        
            _streamWriter.WriteLine(message + ": " + DateTime.Now.GetTimestamp());
            _streamWriter.Flush();
        }
    }
}
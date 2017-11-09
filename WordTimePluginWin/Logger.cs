using System;
using System.IO;
using WordTimePluginWin.Extensions;

namespace WordTimePluginWin
{
    internal class Logger
    {
        private readonly StreamWriter _streamWriter;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath">filepath root is user's homefolder</param>
        public Logger(string filepath)
        {
            _streamWriter = new StreamWriter(Config.Homepath + filepath, true)
            {
                AutoFlush = true                
            };            
        }

        public void Log(string message)
        {        
            _streamWriter.WriteLine(message + ": " + DateTime.Now.GetTimestamp());
        }
    }
}
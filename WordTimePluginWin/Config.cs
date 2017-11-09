using System;

namespace WordTimePluginWin
{
    internal static class Config
    {
        public static string Homepath { get; } = Environment.GetEnvironmentVariable("homepath");        
    }
}
using System;

namespace WordTimePluginWin
{
    internal static class Extensions {        
        public static string GetTimestamp(this DateTime value)
        {
            return value.ToString("G");
        }        
    }
}
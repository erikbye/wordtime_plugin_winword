using System;

namespace WordTimePluginWin
{
    static class Extensions {        
        public static string GetTimestamp(this DateTime value)
        {
            return value.ToString("G");
        }        
    }
}
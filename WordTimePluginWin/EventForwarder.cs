using System;
using Microsoft.Office.Interop.Word;

namespace WordTimePluginWin {
    internal class EventForwarder {        
        
        private DateTime _start;

        public EventForwarder() {
            
        }
        /// <summary>
        /// All we want to do here is check if we are going to send a message or not.
        /// I.e., do nothing unless 30 seconds has passed.
        /// </summary>
        /// <param name="document">The document we will forward and create a JSON object from</param>
        /// 
        internal void Forward(ref Document document) {            
            if (_start == DateTime.MinValue) {
                _start = DateTime.UtcNow;
            } else {
                if (DateTime.UtcNow - _start > TimeSpan.FromSeconds(30))
                {                    
                    Logger.Log("More than 30 seconds passed");

                    /* @TODO
                    If connection to server/ API:
                        Send heartbeats from local DB
                        Drop local DB rows
                        Send heartbeats directly
                    else
                        Store heartbeats in local DB
                    */


                    Logger.Log(document.Name);

                    _start = DateTime.MinValue;
                    return;
                }
                Logger.Log("Less than 30 seconds passed");
            }
        }
    }
}
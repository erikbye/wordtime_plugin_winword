using System;
using Microsoft.Office.Interop.Word;

namespace WordTimePluginWin {
    internal class EventForwarder {        
        private Messager _messager;
        private DateTime _start;

        public EventForwarder() {
            _messager = new Messager();
        }
        /// <summary>
        /// All we want to do here is check if we are going to send a message or not.
        /// I.e., do nothing unless 30 seconds has passed.
        /// </summary>
        /// <param name="document">The document we will forward and create a JSON object from</param>
        internal void Forward(ref Document document) {            
            if (_start == DateTime.MinValue) {
                _start = DateTime.UtcNow;
            } else {
                if (DateTime.UtcNow - _start > TimeSpan.FromSeconds(30)) {
                    Logger.Log("More than 30 seconds passed");

                    var message = _messager.CreateMessage(ref document);
                    _messager.Send(message);

                    _start = DateTime.MinValue;
                    return;
                }             
                Logger.Log("Less than 30 seconds passed");
            }
        }
    }
}
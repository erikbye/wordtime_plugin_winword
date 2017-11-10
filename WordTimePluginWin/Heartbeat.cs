using System;
using Microsoft.Office.Interop.Word;
using WordTimePluginWin.Extensions;

namespace WordTimePluginWin {
    internal class Heartbeat {

        private string _documentName;
        private string _fullName;
        private DateTime _start;

        public Heartbeat() {
            _documentName = "";
            _fullName = "";
        }

        public void Send(ref Document document) {
            // TODO: send heartbeat to API
            // include: timestamp, document filename, editor (e.g. Word Windows)
            // document attributes: Total editing time, Content created (not same as file created), 
            // Word count, Character count, Line count, Paragraph count
            if (_start == DateTime.MinValue) {
                _start = DateTime.UtcNow;
            } else {
                if (DateTime.UtcNow - _start > TimeSpan.FromSeconds(30)) {

                    Logger.Log("more than 30 seconds passed");

                    // Send our heartbeat to the server here
                    this._documentName = document.Name;
                    this._fullName = document.FullName;

                    Logger.Log("_documentName: " + _documentName);
                    Logger.Log("_fullName: " + _fullName);

                    _start = DateTime.MinValue;
                    return;
                }
                Logger.Log("less than 30 seconds passed");
            }
        }
    }
}
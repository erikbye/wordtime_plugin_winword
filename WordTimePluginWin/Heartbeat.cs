using Microsoft.Office.Interop.Word;

namespace WordTimePluginWin
{
    internal class Heartbeat
    {
        private string _documentName;
        private string _fullName;

        public Heartbeat()
        {
            _documentName = "";
            _fullName = "";
        }

        public void Send(ref Document document)
        {
            // TODO: send heartbeat to API
            // include: timestamp, document filename, editor (e.g. Word Windows)
            // document attributes: Total editing time, Content created (not same as file created), 
            // Word count, Character count, Line count, Paragraph count

            this._documentName = document.Name;
            this._fullName = document.FullName;

            // logging is just for testing
            Logger.Log("_documentName: " + _documentName);
            Logger.Log("_fullName: " + _fullName);
        }
    }
}
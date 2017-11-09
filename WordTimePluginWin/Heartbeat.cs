using Microsoft.Office.Interop.Word;

namespace WordTimePluginWin
{
    internal class Heartbeat
    {
        private string _documentName;
        private string _fullName;

        readonly Logger _logger = new Logger(@"\heartbeat.log");

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
            _logger.Log("_documentName: " + _documentName);
            _logger.Log("_fullName: " + _fullName);
        }
    }
}
using System;

namespace WordTimePluginWin
{
    class Heartbeat
    {
        private DateTime _timestamp;
        private string _document;
        private int _documentWordCount;
        
        // private Guid _id;

        public Heartbeat(string document, DateTime timestamp, int documentWordCount)
        {            
            _document = document;
            _timestamp = timestamp;
            _documentWordCount = documentWordCount;
            // _id = new Guid();
        }
    }
}

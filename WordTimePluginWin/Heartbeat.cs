using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Word;
using RabbitMQ.Client;

using WordTimePluginWin.Extensions;
using Task = System.Threading.Tasks.Task;

namespace WordTimePluginWin {
    internal class Heartbeat {

        private string _documentName;
        private string _fullName;
        private DateTime _start;
        private Messager _messager;

        public Heartbeat()
        {
            _messager = new Messager();

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

                    // TODO: Create a heartbeat object
                    this._documentName = document.Name;
                    this._fullName = document.FullName;
                   
                    // TODO: Send heartbeat to message broker
                    Task.Run(() => _messager.Send(_documentName));

                    Logger.Log("_documentName: " + _documentName);
                    Logger.Log("_fullName: " + _fullName);

                    _start = DateTime.MinValue;
                    return;
                }
                Task.Run(() => _messager.Send(_documentName));
                Logger.Log("less than 30 seconds passed");                
            }
        }
    }
}
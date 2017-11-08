using System;
using System.ComponentModel;
using System.IO;
using Microsoft.Office.Tools.Word;

using WordInterop = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using time = System.DateTime;

namespace WordTimePluginWin
{
    public partial class ThisAddIn
    {
        private static readonly string UserHomePath = Environment.GetEnvironmentVariable("homepath");
        private readonly StreamWriter streamWriter = new StreamWriter(UserHomePath + @"\wordtime.log", true);

        #region Add-in and document events

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // TODO: Check if user reaches server, if not, display a message in the statusbar
            // and let user know time stamps will be stored locally and sent 
            // to server when they are back online
            
            // Register event handlers    
            Application.DocumentOpen +=
                DocumentSelectionChange;

            Application.DocumentBeforeSave +=
                DocumentBeforeSave;

            Application.DocumentBeforeClose +=
                DocumentBeforeClose;
            
            ((WordInterop.ApplicationEvents4_Event)Application).NewDocument +=
                DocumentSelectionChange;

            streamWriter.WriteLine("WordTime loaded. " + time.Now.GetTimestamp());
            streamWriter.Flush();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {            
            streamWriter.WriteLine("WordTime shutdown. " + time.Now.GetTimestamp());
            streamWriter.Flush();
        }

        private void DocumentBeforeSave(WordInterop.Document doc, ref bool saveasui, ref bool cancel)
        {
            Document vstoDoc = Globals.Factory.GetVstoObject(Application.ActiveDocument);
            vstoDoc.BeforeSave += ThisDocument_BeforeSave;
        }
        
        // TODO: triggers twice on save, and not on save as
        void ThisDocument_BeforeSave(object sender, SaveEventArgs e)
        {
            var docProperties = (Office.DocumentProperties) Application.ActiveDocument.BuiltInDocumentProperties;
            var totalEditingTime = docProperties["Total editing time"];

            var documentName = Application.ActiveDocument.Name;

           // var fileName = Application.ActiveDocument.FullName;
           // var fileInfo = new FileInfo(fileName);
           // var creationTime = fileInfo.CreationTime.Date.ToString();
           
            streamWriter.WriteLine("Document: " + documentName + " was saved. " + time.Now.GetTimestamp());

            if (totalEditingTime.Value != null)
            {
                streamWriter.WriteLine("Total editing time (minutes): " + totalEditingTime.Value.ToString());
            }            
            streamWriter.Flush();            
        }

        private void DocumentSelectionChange(WordInterop.Document Doc)
        {
            Document vstoDoc = Globals.Factory.GetVstoObject(Application.ActiveDocument);
            vstoDoc.SelectionChange += ThisDocument_SelectionChange;
        }

        // Event fires on selection of text, deselection, return (newline), 
        // arrow keys, but not on backspace, nor if you just keep typing 
        // without using return
        void ThisDocument_SelectionChange(object sender, SelectionEventArgs e)
        {
            streamWriter.WriteLine("Selection changed. " + time.Now.GetTimestamp());
            streamWriter.Flush();
        }

        private void DocumentBeforeClose(WordInterop.Document doc, ref bool cancel)
        {
            Document vstoDoc = Globals.Factory.GetVstoObject(Application.ActiveDocument);
            vstoDoc.BeforeClose += ThisDocument_BeforeClose;
        }

        // TODO: this doesn't seem to trigger
        void ThisDocument_BeforeClose(object sender, CancelEventArgs e)
        {
            streamWriter.WriteLine("Documented closed. " + time.Now.GetTimestamp());
            streamWriter.Flush();
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

    }
}

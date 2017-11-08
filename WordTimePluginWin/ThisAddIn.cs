using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using WordInterop = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using time = System.DateTime;

namespace WordTimePluginWin
{
    public partial class ThisAddIn
    {
        private static readonly string UserHomePath = Environment.GetEnvironmentVariable("homepath");
        private readonly StreamWriter streamWriter = new StreamWriter(UserHomePath + @"\wordtime.log", true);

        #region Add-in and document events

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // TODO: Check if user reaches server, if not, display a message in the statusbar
            // and let user know time stamps will be stored locally and sent 
            // to server when they are back online
            
            // Register event handlers    
            this.Application.DocumentOpen +=
                new WordInterop.ApplicationEvents4_DocumentOpenEventHandler(DocumentSelectionChange);

            this.Application.DocumentBeforeSave +=
                new WordInterop.ApplicationEvents4_DocumentBeforeSaveEventHandler(DocumentBeforeSave);

            this.Application.DocumentBeforeClose +=
                new WordInterop.ApplicationEvents4_DocumentBeforeCloseEventHandler(DocumentBeforeClose);
            
            ((WordInterop.ApplicationEvents4_Event)this.Application).NewDocument +=
                new WordInterop.ApplicationEvents4_NewDocumentEventHandler(DocumentSelectionChange);

            streamWriter.WriteLine("WordTime loaded. " + time.Now.GetTimestamp());
            streamWriter.Flush();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {            
            streamWriter.WriteLine("WordTime shutdown. " + time.Now.GetTimestamp());
            streamWriter.Flush();
        }

        private void DocumentBeforeSave(WordInterop.Document doc, ref bool saveasui, ref bool cancel)
        {
            Document vstoDoc = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
            vstoDoc.BeforeSave += new Microsoft.Office.Tools.Word.SaveEventHandler(ThisDocument_BeforeSave);
        }
        
        // TODO: triggers twice on save, and not on save as
        void ThisDocument_BeforeSave(object sender, Microsoft.Office.Tools.Word.SaveEventArgs e)
        {
            var documentName = this.Application.ActiveDocument.Name;

            streamWriter.WriteLine("Document: " + documentName + " was saved. " + time.Now.GetTimestamp());
            streamWriter.Flush();            
        }

        private void DocumentSelectionChange(Microsoft.Office.Interop.Word.Document Doc)
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
            Document vstoDoc = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
            vstoDoc.BeforeClose += new System.ComponentModel.CancelEventHandler(ThisDocument_BeforeClose);
        }

        // TODO: this doesn't seem to trigger
        void ThisDocument_BeforeClose(object sender, System.ComponentModel.CancelEventArgs e)
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
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

    }
}

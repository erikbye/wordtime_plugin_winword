﻿using System;
using System.ComponentModel;
using Microsoft.Office.Tools.Word;
using WordInterop = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace WordTimePluginWin
{
    public partial class ThisAddIn
    {        
        private readonly Logger _logger = new Logger();

        #region Add-in and document events

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {            
            // TODO 1: Check if user reaches server, if not, display a message in the statusbar
            // and let user know time stamps will be stored locally and sent 
            // to server when they are back online
            // TODO 2: If online: check if user is logged in, if not, display LoginForm
            
            // Register event handlers                
            Application.DocumentOpen += DocumentSelectionChange;
            Application.DocumentBeforeSave += DocumentBeforeSave;
            Application.DocumentBeforeClose += DocumentBeforeClose;
            ((WordInterop.ApplicationEvents4_Event)Application).NewDocument += DocumentSelectionChange;

            _logger.Log("WordTime loaded");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _logger.Log("WordTime shutdown");
        }

        private void DocumentBeforeSave(WordInterop.Document doc, ref bool saveasui, ref bool cancel)
        {
            var vstoDoc = Globals.Factory.GetVstoObject(Application.ActiveDocument);
            vstoDoc.BeforeSave += ThisDocument_BeforeSave;
        }
        
        // TODO: triggers twice on save, and not on save as
        private void ThisDocument_BeforeSave(object sender, SaveEventArgs e)
        {
            var docProperties = (Office.DocumentProperties) Application.ActiveDocument.BuiltInDocumentProperties;
            var totalEditingTime = docProperties["Total editing time"];

            var documentName = Application.ActiveDocument.Name;

            // var fileName = Application.ActiveDocument.FullName;
            // var fileInfo = new FileInfo(fileName);
            // var creationTime = fileInfo.CreationTime.Date.ToString();

            _logger.Log("Document: " + documentName + " was saved. ");

            if (totalEditingTime.Value != null)
            {
                _logger.Log("Total editing time (minutes): " + totalEditingTime.Value.ToString());
            }            
        }

        private void DocumentSelectionChange(WordInterop.Document Doc)
        {            
            var vstoDoc = Globals.Factory.GetVstoObject(Application.ActiveDocument);
            vstoDoc.SelectionChange += ThisDocument_SelectionChange;
        }

        // Event fires on selection of text, deselection, return (newline), 
        // arrow keys, but not on backspace, nor if you just keep typing 
        // without using return
        private void ThisDocument_SelectionChange(object sender, SelectionEventArgs e)
        {
            var documentName = Application.ActiveDocument.Name;

            _logger.Log("Selection changed, working on document " + documentName);
        }

        private void DocumentBeforeClose(WordInterop.Document doc, ref bool cancel)
        {
            var vstoDoc = Globals.Factory.GetVstoObject(Application.ActiveDocument);
            vstoDoc.BeforeClose += ThisDocument_BeforeClose;
        }

        // TODO: this doesn't seem to trigger, and Word seems to use too much time closing the document
        private void ThisDocument_BeforeClose(object sender, CancelEventArgs e)
        {
            var documentName = Application.ActiveDocument.Name;
            _logger.Log("Document closed: " + documentName);
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
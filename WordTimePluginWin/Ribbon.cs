﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using WordTimePluginWin.Properties;
using WordTimePluginWin;
using WordTimePluginWin.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

namespace WordTimePluginWin
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordTimePluginWin.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        public stdole.IPictureDisp GetCustomImage(string name)
        {
            var image = (Bitmap) Resources.ResourceManager.GetObject(name);
            return image != null ? PictureConveter.ImageToPictureDisp(image) : null;
        }

        #region Button Callbacks

        public void OnWordTimeButton(Office.IRibbonControl control)
        {
            var wordTimeForm = new WordTimeForm();
            wordTimeForm.Show();
        }

        public void OnLoginButton(Office.IRibbonControl control)
        {
            var loginForm = new LoginForm();
            loginForm.Show();
        }

        public void OnSettingsButton(Office.IRibbonControl control)
        {
            var settingsForm = new SettingsForm();
            settingsForm.Show();
        }

        public void OnAboutButton(Office.IRibbonControl control)
        {
            var aboutForm = new AboutForm();
            aboutForm.Show();
        }

        public void OnFAQButton(Office.IRibbonControl control)
        {
            var faqForm = new FAQForm();
            faqForm.Show();
        }

        public void OnTourButton(Office.IRibbonControl control)
        {
            var tourForm = new TourForm();
            tourForm.Show();
        }

        #endregion
    }

    internal class PictureConveter : AxHost
    {
        public PictureConveter()
            : base(string.Empty)
        {
        }

        public static stdole.IPictureDisp ImageToPictureDisp(Image image) => (stdole.IPictureDisp) GetIPictureDispFromPicture(image);
        public static stdole.IPictureDisp IconToPictureDisp(Icon icon) => ImageToPictureDisp(icon.ToBitmap());
    }
}

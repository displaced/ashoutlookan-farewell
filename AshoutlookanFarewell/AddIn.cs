using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.IO;

namespace AshoutlookanFarewell
{
    public partial class AddIn
    {
        public static Outlook.Application app = null;

        private void AddIn_Startup(object sender, System.EventArgs e)
        {
            // Grab a reference to the Application object for use elsewhere
            app = this.Application;

            // Load default settings.  We'll add a customisation UI later.
            Settings.LoadDefaults();

            // Ensure we have a folder to store our data
            var dataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AshoutlookanFarewell");
            if (!Directory.Exists(dataDir)) { Directory.CreateDirectory(dataDir); }
            Settings.DataDir = dataDir;

            // Tell our music controller to prepare itself.
            MusicController.Init();

        }

        private void AddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(AddIn_Startup);
            this.Shutdown += new System.EventHandler(AddIn_Shutdown);
        }
        
        #endregion
    }
}

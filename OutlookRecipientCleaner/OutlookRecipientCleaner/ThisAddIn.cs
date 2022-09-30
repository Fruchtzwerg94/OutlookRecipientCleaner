using System;

using System.Diagnostics;

namespace OutlookRecipientCleaner
{
    public partial class ThisAddIn
    {
        public const string ADD_IN_NAME = "Outlook Recipient Cleaner";

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Debug.WriteLine("Starting up", ADD_IN_NAME);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
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
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

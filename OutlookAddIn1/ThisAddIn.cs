using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Globalization;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private const string ADXHTMLFileName = "MaPiFolderTemp.htm";
        private const string ADXHTMLFileName2 = "ADXOlFormGeneral.html";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            #region Add-in Express Regions generated code - do not modify
            this.FormsManager = AddinExpress.OL.ADXOlFormsManager.CurrentInstance;
            this.FormsManager.OnInitialize +=
                new AddinExpress.OL.ADXOlFormsManager.OnComponentInitialize_EventHandler(this.FormsManager_OnInitialize);
            this.FormsManager.ADXBeforeFolderSwitchEx += 
                 new AddinExpress.OL.ADXOlFormsManager.BeforeFolderSwitchEx_EventHandler(FormsManager_ADXBeforeFolderSwitchEx);
           
            this.FormsManager.Initialize(this);
            #endregion

        }
      
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

            #region Add-in Express Regions generated code - do not modify
            this.FormsManager.Finalize(this);
            #endregion
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

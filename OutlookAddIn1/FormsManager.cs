using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using AddinExpress.OL;
using System.Reflection;
using System.Globalization;
using System.Runtime.InteropServices;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {

        public ADXOlFormsManager FormsManager = null;
        public ADXOlFormsCollectionItem ADXOlForm1Item;

        /// <summary>
        /// Use this event to initialize regions and connect to the events of the ADXOlFormsManager class
        /// </summary>
        private void FormsManager_OnInitialize()
        {
            #region Events Initialization 

            // TODO: See the Class Reference for the complete list of events of the ADXOlFormsManager class

            this.FormsManager.ADXBeforeFolderSwitchEx
                += new ADXOlFormsManager.BeforeFolderSwitchEx_EventHandler(FormsManager_ADXBeforeFolderSwitchEx);
            this.FormsManager.ADXFolderSwitch
                += new ADXOlFormsManager.FolderSwitch_EventHandler(FormsManager_ADXFolderSwitch);
            this.FormsManager.ADXFolderSwitchEx
                += new ADXOlFormsManager.FolderSwitchEx_EventHandler(FormsManager_ADXFolderSwitchEx);
            this.FormsManager.ADXNewInspector
                += new ADXOlFormsManager.NewInspector_EventHandler(FormsManager_ADXNewInspector);
            this.FormsManager.OnError
                += new ADXOlFormsManager.Error_EventHandler(FormsManager_OnError);
            
            #endregion 

            #region ADXOlForm1

            // TODO: Use the ADXOlForm1Item properties to configure the region's location, appearance and behavior.
            // See the "The UI Mechanics" chapter of the Add-in Express Developer's Guide for more information.

            ADXOlForm1Item = new ADXOlFormsCollectionItem();
            ADXOlForm1Item.ExplorerLayout = ADXOlExplorerLayout.WebViewPane;
            ADXOlForm1Item.ExplorerItemTypes = ADXOlExplorerItemTypes.olMailItem;
            ADXOlForm1Item.UseOfficeThemeForBackground = true;
            ADXOlForm1Item.FormClassName = typeof(ADXOlFormWPFBrowser).FullName;
            ADXOlForm1Item.Enabled = false;
            this.FormsManager.Items.Add(ADXOlForm1Item);
            #endregion

        }
        
        #region ADXBeforeFolderSwitchEx
        /// <summary>
        /// Occurs before an Outlook explorer goes to a new folder,
        /// either as a result of user action or through program code.
        /// </summary>
        private void FormsManager_ADXBeforeFolderSwitchEx(object sender, AddinExpress.OL.BeforeFolderSwitchExEventArgs args)
        {

            //string folderNameDest = GetFullFolderName(args.DstFolder);

            //if (IsSharePointLinkFolder((Outlook.MAPIFolder)args.DstFolder))
            //{
            //    ADXOlForm1Item.Enabled = true;
            //}
            //else
            //{
            //   ADXOlForm1Item.Enabled = false;
            //}

        }
        #endregion


        private bool IsSharePointLinkFolder(Outlook.MAPIFolder folder)
        {
            if (folder == null || folder.WebViewURL == null || folder.WebViewURL.Length <= 0)
                return false;

            bool check1 = false, check2 = false;

            if (folder.WebViewOn)
            {
                check1 = (CultureInfo.InvariantCulture.CompareInfo.IndexOf(folder.WebViewURL,
                    ADXHTMLFileName, System.Globalization.CompareOptions.IgnoreCase) > 0);
                check2 = (CultureInfo.InvariantCulture.CompareInfo.IndexOf(folder.WebViewURL,
                    ADXHTMLFileName2, System.Globalization.CompareOptions.IgnoreCase) > 0);
                //bool check3 = folder.IsSharePointFolder;
            }

            return check1 || check2;
        }

        

        #region ADXFolderSwitch
        /// <summary>
        /// Occurs before an Outlook explorer goes to a new folder,
        /// either as a result of user action or through program code.
        /// </summary>
        private void FormsManager_ADXFolderSwitch(object sender, AddinExpress.OL.FolderSwitchEventArgs args)
        {
           
        }
        #endregion
        
        #region ADXFolderSwitchEx
        /// <summary>
        /// Occurs when an Outlook explorer goes to a new folder, 
        /// either as a result of user action or through program code.
        /// </summary>
        /// <remarks>
        /// Set args.ShowForm = False to prevent any ADXOlForm from display. This also prevents 
        /// the ADXFolderSwitch events from firing.
        /// <para>To prevent a given form instance from being shown, you set ADXOlFom.Visible = false
        /// in the ADXBeforeFormShow event of the corresponding ADXOlForm. </para>
        /// </remarks>
        private void FormsManager_ADXFolderSwitchEx(object sender, AddinExpress.OL.FolderSwitchExEventArgs args)
        {
        
        }
        #endregion
                  
        #region ADXNewInspector
        /// <summary>
        /// Occurs whenever a new inspector window is opened,
        /// either as a result of user action or through program code.
        /// </summary>
        private void FormsManager_ADXNewInspector(object inspectorObj)
        {
        
        }
        #endregion
        
        #region OnError
        /// <summary>
        /// Occurs when ADXOlFormaManager generates an exception.
        /// </summary>
        private void FormsManager_OnError(object sender, AddinExpress.OL.ErrorEventArgs args)
        {
        
        }
        #endregion

        #region RequestService
        /// <summary>
        /// Required method for DockRight, DockLeft, DockTop and DockBottom layout support.
        /// </summary>
        protected override object RequestService(Guid serviceGuid)
        {
            if (serviceGuid == typeof(Office.ICustomTaskPaneConsumer).GUID)
            {
                return AddinExpress.OL.CTPFactoryGettingTaskPane.Instance;
            }
            return base.RequestService(serviceGuid);
        }
        #endregion
    }
}

using System;
using Microsoft.Office.Tools.Ribbon;
using SharePoint_Link.UserModule;
using SharePoint_Link.Utility;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Xml;

namespace SharePoint_Link
{
    /// <summary>
    /// <c>SharePointRibbon</c>
    /// class implements the functionality to display itopia menu 
    /// </summary>
    public partial class SharePointRibbon
    {
        /// <summary>
        /// <c>frmSPSiteConfigurationObject</c> an object of window form 
        /// <c>frmSPSiteConfiguration</c>
        /// required to open new connection
        /// </summary>
        frmSPSiteConfiguration frmSPSiteConfigurationObject;

        /// <summary>
        /// <c>userOptions</c> XMLLogOptions class object declaration
        /// </summary>
        XMLLogOptions userOptions;

        /// <summary>
        /// <c>SharePointRibbon_Load</c> event handler
        /// get configuration properties from config file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SharePointRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            if (ThisAddIn.IsAuthorized)
            {
                userOptions = UserLogManagerUtility.GetUserConfigurationOptions();
            }
            else
            {
                btnConnectionProperties.Enabled = false;
                btnNewConnection.Enabled = false;
                btnOptions.Enabled = false;
            }
        }

        /// <summary>
        /// <c>btnConnectionProperties_Click</c> event handler
        /// display window form to display connection property window form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConnectionProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UpdateFoldersLocations();
                frmConnectionProperties objfrmConnectionProperties = new frmConnectionProperties();
                objfrmConnectionProperties.ShowDialog();

            }
            catch (Exception ex)
            {

            }
        }


        /// <summary>
        /// <c>UpdateFoldersLocations</c> member function
        /// update folder location
        /// </summary>
        private void UpdateFoldersLocations()
        {
            try
            {
                Outlook._Application outlookObj;
                outlookObj = Globals.ThisAddIn.Application;
                Microsoft.Office.Interop.Outlook.NameSpace outlookNameSpace = outlookObj.GetNamespace("MAPI");
                Outlook.MAPIFolder oInBox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Outlook.MAPIFolder parentFolder = (Outlook.MAPIFolder)oInBox;
                XmlNodeList xFolders = UserLogManagerUtility.GetAllFoldersDetails(UserStatus.Active);
                if (xFolders != null)
                {
                    string folderName = string.Empty, DocLibName = string.Empty;

                    foreach (XmlNode xNode in xFolders)
                    {
                        try
                        {
                            folderName = xNode.ChildNodes[2].InnerText;
                            //Get Doc Lib Name
                            DocLibName = xNode.ChildNodes[3].InnerText;
                            Outlook.MAPIFolder oChildFolder = MAPIFolderWrapper.GetFolder(parentFolder, folderName); // parentFolder.Folders[folderName];

                            if (oChildFolder != null && oChildFolder.Name == folderName)
                            {

                                if (oChildFolder.FolderPath.IndexOf("\\Deleted Items\\") == -1)
                                {
                                    UserLogManagerUtility.UpdateFolderConfigNodeDetails(oChildFolder.Name, "OutlookLocation", oChildFolder.FolderPath);
                                }
                            }
                        }
                        catch { }
                    }

                }

            }
            catch (Exception)
            {


            }
        }


        /// <summary>
        /// <c>btnNewConnection_Click</c> Event Handler
        /// display new connection window form to create new connection
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNewConnection_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.IsUrlIsTyped = true;
                frmSPSiteConfigurationObject = new frmSPSiteConfiguration();
                frmSPSiteConfigurationObject.ShowDialog();
            }
            catch (Exception ex)
            {

            }
        }


        /// <summary>
        /// <c>btnOptions_Click</c> event handler
        /// display Auto Delete Email option form to modify AdutDelete options.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOptions_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {

                frmOptions objfrmOptions = new frmOptions(userOptions);

                objfrmOptions.ShowDialog();
                if (objfrmOptions.ValuesUpdated)
                {
                    userOptions.AutoDeleteEmails = objfrmOptions.isAutoDeleteChecked;
                    UserLogManagerUtility.CreateXMLFileForStoringUserOptions(userOptions);
                }


            }
            catch (Exception ex)
            {

            }
        }

        private void Uploads_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.CustomTaskPanes[0].Visible = true;
                Globals.ThisAddIn.CustomTaskPanes[0].DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating;
            }
            catch (Exception ex)
            {
                
            }
        }




    }
}

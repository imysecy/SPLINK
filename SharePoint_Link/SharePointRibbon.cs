using System;
using Microsoft.Office.Tools.Ribbon;
using SharePoint_Link.UserModule;
using SharePoint_Link.Utility;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Xml;
using System.Data;
using System.IO;

namespace SharePoint_Link
{
    /// <summary>
    /// <c>SharePointRibbon</c>
    /// class implements the functionality to display itopia menu 
    /// </summary>
    public partial class SharePointRibbon
    {


        frmFolderOptions frmoption;
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

        bool folderOptionFrmIsOpen;
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
        /// <summary>
        /// code written by Joy
        /// this event fires when the Upload ribbon control is clicked
        /// makes visible the custom taskpane and reposition the dock position of the custom taskpane
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// code written by Joy
        /// fires when the copy ribbon control is clicked
        /// sets the mapped folder names to the dropdownlist
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Copy_Click(object sender, RibbonControlEventArgs e)
        {
            DataSet ds = new DataSet();

            if (File.Exists(UserLogManagerUtility.XMLFilePath))
            {
                ds.Tables.Clear();
                ds.ReadXml(UserLogManagerUtility.XMLFilePath);
                try
                {
                    if(ds.Tables[0].Rows.Count>0)
                    {
                        if (Globals.ThisAddIn.isuploadRunning == false && Globals.ThisAddIn.isTimerUploadRunning == false && Globals.ThisAddIn.isMoveRunning == false && Globals.ThisAddIn.isCopyRunninng == false)
                        {
                            Globals.ThisAddIn.copy_button_clicked = true;
                            Globals.ThisAddIn.move_button_clicked = false;

                            Outlook.Application myApplication = Globals.ThisAddIn.Application;
                            Outlook.Explorer myActiveExplorer = (Outlook.Explorer)myApplication.ActiveExplorer();
                            Globals.ThisAddIn.copySelected = myActiveExplorer.Selection;
                            if (Globals.ThisAddIn.copySelected.Count > 0)
                            {
                                
                                Globals.ThisAddIn.isCopyRunninng = true;
                                Globals.ThisAddIn.no_of_copied_item_uploaded = 0;
                                Globals.ThisAddIn.no_of_copied_item_to_be_uploaded = Globals.ThisAddIn.copySelected.Count;
                                frmoption = new frmFolderOptions();
                                frmoption.ShowDialog();
                            }
                            else
                            {
                                frmMessageWindow messagebox = new frmMessageWindow();
                                messagebox.DisplayMessage = "Please select some items";
                                messagebox.TopLevel = true;
                                messagebox.TopMost = true;
                                messagebox.ShowDialog();
                                messagebox.Dispose();

                                return;
                            }
                        }
                        else
                        {
                            frmMessageWindow objMessage = new frmMessageWindow();
                            objMessage.DisplayMessage = "Your uploads are still running.Please wait for sometime.";
                            objMessage.TopLevel = true;
                            objMessage.TopMost = true;
                            objMessage.ShowDialog();
                            objMessage.Dispose();

                            return;
                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }
            
        }
        /// <summary>
        /// code written by Joy
        /// this event is fired when the move button is clicked
        /// sets the mapped folder names to the dropdownlist
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Move_Click(object sender, RibbonControlEventArgs e)
        {
            DataSet ds = new DataSet();

            if (File.Exists(UserLogManagerUtility.XMLFilePath))
            {
                ds.Tables.Clear();
                ds.ReadXml(UserLogManagerUtility.XMLFilePath);
                try
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        if (Globals.ThisAddIn.isuploadRunning == false && Globals.ThisAddIn.isTimerUploadRunning == false && Globals.ThisAddIn.isMoveRunning == false && Globals.ThisAddIn.isCopyRunninng == false)
                        {
                            Globals.ThisAddIn.copy_button_clicked = false;
                            Globals.ThisAddIn.move_button_clicked = true;

                            Outlook.Application myApplication = Globals.ThisAddIn.Application;
                            Outlook.Explorer myActiveExplorer = (Outlook.Explorer)myApplication.ActiveExplorer();
                            Globals.ThisAddIn.moveSelected = myActiveExplorer.Selection;
                            if (Globals.ThisAddIn.moveSelected.Count > 0)
                            {
                                Globals.ThisAddIn.isMoveRunning = true;
                                Globals.ThisAddIn.no_of_moved_item_uploaded = 0;
                                Globals.ThisAddIn.no_of_moved_item_to_be_uploaded = Globals.ThisAddIn.moveSelected.Count;
                                frmoption = new frmFolderOptions();
                                frmoption.ShowDialog();

                            }
                            else
                            {
                                frmMessageWindow messagebox = new frmMessageWindow();
                                messagebox.DisplayMessage = "Please select some items";
                                messagebox.TopLevel = true;
                                messagebox.TopMost = true;
                                messagebox.ShowDialog();
                                messagebox.Dispose();

                                return;
                            }
                           
                        }
                        else
                        {
                            frmMessageWindow objMessage = new frmMessageWindow();
                            objMessage.DisplayMessage = "Your uploads are still running.Please wait for sometime.";
                            objMessage.TopLevel = true;
                            objMessage.TopMost = true;
                            objMessage.ShowDialog();
                            objMessage.Dispose();

                            return;

                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }
            

           
            
        }




    }
}

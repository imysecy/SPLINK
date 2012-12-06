using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.IO;
using SharePoint_Link.UserModule;
using SharePoint_Link.Utility;
using System.Text.RegularExpressions;
using Utility;
using System.Xml;
using System.ComponentModel;

namespace SharePoint_Link
{
    /// <summary>
    /// this wrappwer class written by Joy for timer upload process
    /// Wrapper class to Fire Item add event on outlook folders
    /// </summary>
    public class UploadBrokenUploads
    {
        #region Global Variables
        ///////////////////////Modified by Joy on 25.07.2012///////////////////////////////
       // private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        public static Outlook._Application OutlookObj;
        Outlook.NameSpace outlookNameSpace;
        Outlook.Folders oMailRootFolders;
        Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        /////////////////////////////////////////////////////////////////////////////////

        Outlook.MAPIFolder activeDroppingFolder;
        Outlook.Items activeDroppingFolderItems;
        Outlook.Explorer addinExplorer;
        Outlook.MailItem mailitem;
        bool isUserDroppedItemsCanUpload = false;
        Outlook.MAPIFolder parentfolder;
        private TypeOfMailItem ItemType = TypeOfMailItem.Mail;
        string mailitemEntryID = string.Empty;
        delegate void Add();
        /// <summary>
        /// code written by Joy
        /// decalaration for BackgoundWorker
        /// </summary>
        BackgroundWorker bwg;
        /// <summary>
        /// code written by Joy
        /// delegate for updating the progressbar's upload status
        /// </summary>
        delegate void updateProgresStatus();
        /// <summary>
        /// code written by Joy
        /// delegate for upadating the progressbar
        /// </summary>
        delegate void updateProgessBar();
        #endregion

        #region Methods

      /// <summary>
        ///  assigns properties to member fields
      /// </summary>
      /// <param name="outlookFolder"></param>
      /// <param name="outlookExplorer"></param>
      /// <param name="isFolderMappedWithDocLibrary"></param>
        public UploadBrokenUploads(Outlook.MAPIFolder outlookFolder, Outlook.Explorer outlookExplorer, bool isFolderMappedWithDocLibrary)
        {


            try
            {
                FolderName = outlookFolder.Name;
                //Get the details of the folder. Is it is mapped to SP DocLib or SPSIte(Pages)
                isUserDroppedItemsCanUpload = isFolderMappedWithDocLibrary;
                addinExplorer = outlookExplorer;
                activeDroppingFolder = outlookFolder;
                activeDroppingFolderItems = outlookFolder.Items;
                


            }
            catch (Exception ex)
            { }

            


        }

        #endregion
        /// <summary>
        /// code written by Joy
        /// invokes the frmUploadItemsList user control to the custom taskpane
        /// </summary>
        private void MyAddCustomTaskPane()
        {
            if (Globals.ThisAddIn.frmlistObject == null)
            {
                frmUploadItemsListObject = new frmUploadItemsList();
                Globals.ThisAddIn.frmlistObject = frmUploadItemsListObject;
                Globals.ThisAddIn.myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(frmUploadItemsListObject, "ITOPIA");
                Globals.ThisAddIn.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating;
                Globals.ThisAddIn.myCustomTaskPane.DockPositionChanged += new EventHandler(myCustomTaskPane_DockPositionChanged);
                Globals.ThisAddIn.myCustomTaskPane.Height = 500;
                Globals.ThisAddIn.myCustomTaskPane.Width = 600;
                frmUploadItemsListObject.ShowForm(folderName);
                frmUploadItemsListObject.Show();
            }
            else
            {
                frmUploadItemsListObject = Globals.ThisAddIn.frmlistObject;
                //frmUploadItemsListObject.Refresh();
                //Globals.ThisAddIn.myCustomTaskPane.Control.Refresh();
                frmUploadItemsListObject.lblPleaseWaitMessage.Text = "Please Wait - Uploading Items";
                frmUploadItemsListObject.ShowForm(folderName);
                frmUploadItemsListObject.Show();
            }
        }

/// <summary>
/// code wriiten by Joy
///  executes the delegate Add and invoking all of this to the main thread's form invoke method
/// </summary>

        void IAddCustomTaskPane()
        {

            Add add = new Add(MyAddCustomTaskPane);

            Globals.ThisAddIn.form.Invoke(add);

        }

        #region Events

        /// <summary>
        /// <c>frmUploadItemsListObject</c> class object of  <c>frmUploadItemsList</c> wondow form
        /// </summary>
        frmUploadItemsList frmUploadItemsListObject;
        string strMailSubjectReplcePattern = @"([{}\(\)\^$&_%#!@=<>:;,~`'\’ \*\?\/\+\|\[\\\\]|\]|\-)";
        string strAttachmentReplacePattern = @"([{}\(\)\^$&%#!@=<>:;,~`'\’ \*\?\/\+\|\[\\\\]|\]|\-)";


        /// <summary>
        /// fires the BackgroundWorker's DoWork event if backgroundwoker is not busy
        /// code written by Joy
        /// </summary>
     
      public  void uploadBrokenUploadsIfExists()
        {
            if (bwg == null)
            {
                bwg = new BackgroundWorker();
                bwg.DoWork += new DoWorkEventHandler(bwg_DoWork);
                //bwg.DoWork += delegate(object sender, DoWorkEventArgs e) { bwg_DoWork(sender, e, Item); };
                GC.KeepAlive(bwg);
            }
            if (!bwg.IsBusy)
            {
                bwg.RunWorkerAsync();
            }
        }

        /// <summary>
        /// code written by Joy
        /// excutes ansd the start the timer upload process when the backgoundworkers's do work event fires
        /// </summary>
        /// <param name="Item"></param>
      void doBackGroundUpload(object Item)
      {
          try
          {
             
              Globals.ThisAddIn.isTimerUploadRunning = true;
              OutlookObj = Globals.ThisAddIn.Application;
              outlookNameSpace = OutlookObj.GetNamespace("MAPI");
              Outlook.MAPIFolder oInBox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
              Outlook.MAPIFolder olMailRootFolder = (Outlook.MAPIFolder)oInBox.Parent;
              oMailRootFolders = olMailRootFolder.Folders;
              Outlook.MailItem moveMail = (Outlook.MailItem)Item;
              string newCatName = "Successfully Uploaded";
              if (Globals.ThisAddIn.Application.Session.Categories[newCatName] == null)
              {
                  outlookNameSpace.Categories.Add(newCatName, Outlook.OlCategoryColor.olCategoryColorDarkGreen, Outlook.OlCategoryShortcutKey.olCategoryShortcutKeyNone);
              }

              
              XmlNode uploadFolderNode = UserLogManagerUtility.GetSPSiteURLDetails("", folderName);

              if (uploadFolderNode != null)
              {
                  bool isDroppedItemUplaoded = false;

                  addinExplorer = ThisAddIn.OutlookObj.ActiveExplorer();

                  //Check the folder mapping with documnet library

                  if (isUserDroppedItemsCanUpload == false)
                  {
                      //Show message
                      try
                      {


                          Outlook.MailItem m = (Outlook.MailItem)Item;
                          mailitemEntryID = m.EntryID;

                          try
                          {
                              mailitem = m;

                              mailitemEntryID = m.EntryID;

                              string strsubject = m.EntryID;
                              if (string.IsNullOrEmpty(strsubject))
                              {
                                  strsubject = "tempomailcopy";
                              }

                              mailitemEntryID = strsubject;

                              string tempFilePath = UserLogManagerUtility.RootDirectory + "\\" + strsubject + ".msg";

                              if (Directory.Exists(UserLogManagerUtility.RootDirectory) == false)
                              {
                                  Directory.CreateDirectory(UserLogManagerUtility.RootDirectory);
                              }
                              m.SaveAs(tempFilePath, Outlook.OlSaveAsType.olMSG);


                          }
                          catch (Exception ex)
                          {


                          }

                          Outlook.MAPIFolder fp = (Outlook.MAPIFolder)m.Parent;
                          DoNotMoveInNonDocLib(mailitemEntryID, fp);




                      }
                      catch (Exception)
                      {
                          NonDocMoveReportItem(Item);
                      }


                      MessageBox.Show("You are attempting to move files to a non document library. This action is not supported.", "ITOPIA", MessageBoxButtons.OK, MessageBoxIcon.Information);

                      return;

                  }


                  if (frmUploadItemsListObject == null || (frmUploadItemsListObject != null && frmUploadItemsListObject.IsDisposed == true))
                  {
                      //frmUploadItemsListObject = new frmUploadItemsList();

                      // myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(frmUploadItemsListObject, "ITOPIA");
                      //myCustomTaskPane.Visible = true;

                      IAddCustomTaskPane();



                  }
                  //frmUploadItemsListObject.TopLevel = true;
                  //frmUploadItemsListObject.TopMost = true;

                  ////////////////////// frmUploadItemsListObject.Show();

                  try
                  {

                      //////
                      //////////
                      Outlook.MailItem oMailItem = (Outlook.MailItem)Item;
                      parentfolder = (Outlook.MAPIFolder)oMailItem.Parent;
                      try
                      {
                          mailitem = oMailItem;

                          mailitemEntryID = oMailItem.EntryID;

                          string strsubject = oMailItem.EntryID;
                          if (string.IsNullOrEmpty(strsubject))
                          {
                              strsubject = "tempomailcopy";
                          }


                          mailitemEntryID = strsubject;

                          string tempFilePath = UserLogManagerUtility.RootDirectory + "\\" + strsubject + ".msg";

                          if (Directory.Exists(UserLogManagerUtility.RootDirectory) == false)
                          {
                              Directory.CreateDirectory(UserLogManagerUtility.RootDirectory);
                          }
                          oMailItem.SaveAs(tempFilePath, Outlook.OlSaveAsType.olMSG);


                      }
                      catch (Exception ex)
                      {

                      }

                      string fileName = string.Empty;
                      if (!string.IsNullOrEmpty(oMailItem.Subject))
                      {
                          //Replce any specila characters in subject
                          fileName = Regex.Replace(oMailItem.Subject, strMailSubjectReplcePattern, " ");
                          fileName = fileName.Replace(".", "_");
                      }

                      if (string.IsNullOrEmpty(fileName))
                      {
                          DateTime dtReceivedDate = Convert.ToDateTime(oMailItem.ReceivedTime);
                          fileName = "Untitled_" + dtReceivedDate.Day + "_" + dtReceivedDate.Month + "_" + dtReceivedDate.Year + "_" + dtReceivedDate.Hour + "_" + dtReceivedDate.Minute + "_" + dtReceivedDate.Millisecond;
                      }

                      UploadItemsData newUploadData = new UploadItemsData();
                      newUploadData.ElapsedTime = DateTime.Now;
                      newUploadData.UploadFileName = fileName;// oMailItem.Subject;
                      newUploadData.UploadFileExtension = ".msg";
                      newUploadData.UploadingMailItem = oMailItem;
                      newUploadData.UploadType = TypeOfUploading.Mail;
                      newUploadData.DisplayFolderName = folderName;
                      frmUploadItemsListObject.UploadUsingDelegate(newUploadData);
                      //Set dropped items is uploaded
                      /////////////////////////updated by Joy on 25.07.2012/////////////////////////////////
                      bool uploadStatus = frmUploadItemsListObject.IsSuccessfullyUploaded;
                      XMLLogOptions userOption = UserLogManagerUtility.GetUserConfigurationOptions();
                      if (uploadStatus == true)
                      {
                          // Globals.ThisAddIn.isTimerUploaded = true;
                          isDroppedItemUplaoded = true;

                          for (int i = 0; i <= activeDroppingFolder.Items.Count; i++)
                          {
                              try
                              {
                                  Outlook.MailItem me = (Outlook.MailItem)activeDroppingFolder.Items[i];

                                  if (me.EntryID == mailitemEntryID)
                                  {
                                      me.Categories.Remove(0);
                                      me.Categories = newCatName;
                                      me.Save();

                                      if (userOption.AutoDeleteEmails == true)
                                      {
                                          UserMailDeleteOption(mailitemEntryID, parentfolder);
                                      }
                                  }
                              }
                              catch (Exception ex)
                              {


                              }
                          }

                          frmUploadItemsListObject.lblPRStatus.Invoke(new updateProgresStatus(() =>
                          {
                              frmUploadItemsListObject.lblPRStatus.Text = Globals.ThisAddIn.no_of_t_item_uploaded.ToString() + " " + "of" + " " + Globals.ThisAddIn.no_of_pending_items_to_be_uploaded.ToString() + " " + "Uploaded";
                          }));
                          frmUploadItemsListObject.progressBar1.Invoke(new updateProgessBar(() =>
                          {
                              frmUploadItemsListObject.progressBar1.Value = (((Globals.ThisAddIn.no_of_t_item_uploaded * 100 / Globals.ThisAddIn.no_of_pending_items_to_be_uploaded)));
                          }));

                      }
                      else
                      {
                          isDroppedItemUplaoded = false;
                      }

                      /////////////////////////updated by Joy on 25.07.2012/////////////////////////////////
                  }
                  catch (Exception ex)
                  {
                      isDroppedItemUplaoded = MoveItemIsReportItem(Item);
                  }

                  try
                  {
                      if (isDroppedItemUplaoded == false)
                      {
                          //string tempName = oDocItem.Subject;
                          string tempName = string.Empty;
                          Outlook.DocumentItem oDocItem = (Outlook.DocumentItem)Item;


                          try
                          {

                              Outlook._MailItem myMailItem = (Outlook.MailItem)addinExplorer.Selection[1];
                              foreach (Outlook.Attachment oAttachment in myMailItem.Attachments)
                              {
                                  if (oAttachment.FileName == oDocItem.Subject)
                                  {
                                      tempName = oAttachment.FileName;
                                      tempName = tempName.Substring(tempName.LastIndexOf("."));
                                      oAttachment.SaveAsFile(UserLogManagerUtility.RootDirectory + @"\tempattachment" + tempName);

                                      //Read file data to bytes
                                      //byte[] fileBytes = File.ReadAllBytes(UserLogManagerUtility.RootDirectory + @"\tempattachment" + tempName);
                                      System.IO.FileStream Strm = new System.IO.FileStream(UserLogManagerUtility.RootDirectory + @"\tempattachment" + tempName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                                      System.IO.BinaryReader reader = new System.IO.BinaryReader(Strm);
                                      byte[] fileBytes = reader.ReadBytes(Convert.ToInt32(Strm.Length));
                                      reader.Close();
                                      Strm.Close();

                                      //Replace any special characters are there in file name
                                      string fileName = Regex.Replace(oAttachment.FileName, strAttachmentReplacePattern, " ");

                                      //Add uplaod attachment item data to from list.
                                      UploadItemsData newUploadData = new UploadItemsData();
                                      newUploadData.UploadType = TypeOfUploading.Attachment;
                                      newUploadData.AttachmentData = fileBytes;
                                      newUploadData.DisplayFolderName = activeDroppingFolder.Name;


                                      if (fileName.Contains("."))
                                      {
                                          newUploadData.UploadFileName = fileName.Substring(0, fileName.LastIndexOf("."));
                                          newUploadData.UploadFileExtension = fileName.Substring(fileName.LastIndexOf("."));

                                          if (string.IsNullOrEmpty(newUploadData.UploadFileName.Trim()))
                                          {
                                              //check file name conatins empty add the date time 
                                              newUploadData.UploadFileName = "Untitled_" + DateTime.Now.ToFileTime();

                                          }
                                      }

                                      //Add to form
                                      frmUploadItemsListObject.UploadUsingDelegate(newUploadData);
                                      //Set dropped mail attachment items is uploaded.
                                      isDroppedItemUplaoded = true;
                                      newUploadData = null;
                                      //oDocItem.Delete();
                                      break;
                                  }
                              }
                          }
                          catch (InvalidCastException ex)
                          {
                              //Set dropped mail attachment items is uploaded to false
                              isDroppedItemUplaoded = false;
                          }

                          if (isDroppedItemUplaoded == false)
                          {
                              tempName = oDocItem.Subject;
                              tempName = tempName.Substring(tempName.LastIndexOf("."));
                              oDocItem.SaveAs(UserLogManagerUtility.RootDirectory + @"\tempattachment" + tempName, Type.Missing);

                              System.IO.FileStream Strm = new System.IO.FileStream(UserLogManagerUtility.RootDirectory + @"\tempattachment" + tempName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                              System.IO.BinaryReader reader = new System.IO.BinaryReader(Strm);
                              byte[] fileBytes = reader.ReadBytes(Convert.ToInt32(Strm.Length));
                              reader.Close();
                              Strm.Close();

                              //Replace any special characters are there in file name
                              string fileName = Regex.Replace(oDocItem.Subject, strAttachmentReplacePattern, " ");

                              //Add uplaod attachment item data to from list.
                              UploadItemsData newUploadData = new UploadItemsData();
                              newUploadData.UploadType = TypeOfUploading.Attachment;
                              newUploadData.AttachmentData = fileBytes;
                              newUploadData.DisplayFolderName = activeDroppingFolder.Name;


                              if (fileName.Contains("."))
                              {
                                  newUploadData.UploadFileName = fileName.Substring(0, fileName.LastIndexOf("."));
                                  newUploadData.UploadFileExtension = fileName.Substring(fileName.LastIndexOf("."));

                                  if (string.IsNullOrEmpty(newUploadData.UploadFileName.Trim()))
                                  {
                                      //check file name conatins empty add the date time 
                                      newUploadData.UploadFileName = "Untitled_" + DateTime.Now.ToFileTime();

                                  }
                              }

                              //Add to form
                              frmUploadItemsListObject.UploadUsingDelegate(newUploadData);
                              newUploadData = null;
                              //oDocItem.Delete();
                          }

                      }
                  }
                  catch (Exception ex)
                  {
                      //throw ex;
                      //////////////////////////////updated by Joy on 28.07.2012///////////////////////////////////
                      //  EncodingAndDecoding.ShowMessageBox("FolderItem Add Event_DocItem Conv", ex.Message, MessageBoxIcon.Error);
                      //////////////////////////////updated by Joy on 28.07.2012///////////////////////////////////
                  }




                  try
                  {
                      XMLLogOptions userOptions = UserLogManagerUtility.GetUserConfigurationOptions();

                      for (int i = 0; i <= parentfolder.Items.Count; i++)
                      {
                          try
                          {
                              Outlook.MailItem me = (Outlook.MailItem)parentfolder.Items[i];

                              if (me.EntryID == mailitemEntryID)
                              {
                                  ///////////////////////////modified by Joy on 10.08.2012////////////////////////////////////

                                  if (isDroppedItemUplaoded == true)
                                  {

                                      me.Categories.Remove(0);
                                      me.Categories = newCatName;
                                      me.Save();
                                      if (userOptions.AutoDeleteEmails == true)
                                      {
                                          UserMailDeleteOption(mailitemEntryID, parentfolder);
                                      }
                                      //parentfolder.Items.Remove(i);
                                  }
                                  ///////////////////////////modified by Joy on 10.08.2012////////////////////////////////////

                              }
                          }
                          catch (Exception)
                          {


                          }
                      }
                  }


                  catch (Exception)
                  {


                  }
                  if (!string.IsNullOrEmpty(mailitemEntryID))
                  {
                      if (ItemType == TypeOfMailItem.ReportItem)
                      {
                          UserReportItemDeleteOption(mailitemEntryID, parentfolder);
                      }
                      else
                      {
                          ///////////////////////////Updated by Joy on 16.08.2012....to be updated later///////////////////////////////
                          // UserMailDeleteOption(mailitemEntryID, parentfolder);
                          ///////////////////////////Updated by Joy on 16.08.2012....to be updated later///////////////////////////////
                      }
                  }

              }

          }
          catch (Exception ex)
          {
              EncodingAndDecoding.ShowMessageBox("Folder Item Add Event", ex.Message, MessageBoxIcon.Error);

          }

          //AddToUploadList(Item);
      }
        /// <summary>
        /// code written by Joy
        /// this event executed by the background worker calls the  doBackGroundUpload with a mail item
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

      void bwg_DoWork(object sender, DoWorkEventArgs e)
      {
         
          Globals.ThisAddIn.isTimerUploadRunning = true;
          foreach (Outlook.MailItem mItem in Globals.ThisAddIn.pendingList)
          {
              doBackGroundUpload(mItem);
          }
         
         
          Globals.ThisAddIn.isTimerUploadRunning = false;
      }
        /// <summary>
        /// code written by Joy
        /// reset the docking position of the custom taskpane 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void myCustomTaskPane_DockPositionChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating;
        }


        /// <summary>
        /// <c>MoveItemIsReportItem</c> member function
        /// move reportItem to the sourcefolder if autodelete option is  not checked
        /// </summary>
        /// <param name="Item"></param>
        /// <returns></returns>
        private bool MoveItemIsReportItem(object Item)
        {
            bool isDroppedItemUplaoded = false;
            try
            {

                //////////
                Outlook.ReportItem oMailItem = (Outlook.ReportItem)Item;
                ItemType = TypeOfMailItem.ReportItem;
                parentfolder = (Outlook.MAPIFolder)oMailItem.Parent;
                try
                {


                    mailitemEntryID = oMailItem.EntryID;

                    string strsubject = oMailItem.EntryID;
                    if (string.IsNullOrEmpty(strsubject))
                    {
                        strsubject = "tempomailcopy";
                    }

                    mailitemEntryID = strsubject;

                    string tempFilePath = UserLogManagerUtility.RootDirectory + "\\" + strsubject + ".msg";

                    if (Directory.Exists(UserLogManagerUtility.RootDirectory) == false)
                    {
                        Directory.CreateDirectory(UserLogManagerUtility.RootDirectory);
                    }
                    oMailItem.SaveAs(tempFilePath, Outlook.OlSaveAsType.olMSG);


                }
                catch (Exception ex)
                {

                }

                string fileName = string.Empty;
                if (!string.IsNullOrEmpty(oMailItem.Subject))
                {
                    //Replce any specila characters in subject
                    fileName = Regex.Replace(oMailItem.Subject, strMailSubjectReplcePattern, " ");
                    fileName = fileName.Replace(".", "_");
                }

                if (string.IsNullOrEmpty(fileName))
                {
                    DateTime dtReceivedDate = Convert.ToDateTime(oMailItem.CreationTime);
                    fileName = "Untitled_" + dtReceivedDate.Day + "_" + dtReceivedDate.Month + "_" + dtReceivedDate.Year + "_" + dtReceivedDate.Hour + "_" + dtReceivedDate.Minute + "_" + dtReceivedDate.Millisecond;
                }

                UploadItemsData newUploadData = new UploadItemsData();
                newUploadData.UploadFileName = fileName;// oMailItem.Subject;
                newUploadData.UploadFileExtension = ".msg";
                newUploadData.UploadingReportItem = oMailItem;
                newUploadData.UploadType = TypeOfUploading.Mail;
                newUploadData.TypeOfMailItem = TypeOfMailItem.ReportItem;
                newUploadData.DisplayFolderName = folderName;
                frmUploadItemsListObject.UploadUsingDelegate(newUploadData);
                //Set dropped items is uploaded
                isDroppedItemUplaoded = true;

            }
            catch (Exception ex) { }
            try
            {

                XMLLogOptions userOptions = UserLogManagerUtility.GetUserConfigurationOptions();
                if (userOptions.AutoDeleteEmails == true)
                {
                    for (int i = 0; i <= parentfolder.Items.Count; i++)
                    {
                        try
                        {
                            Outlook.ReportItem me = (Outlook.ReportItem)parentfolder.Items[i];

                            if (me.EntryID == mailitemEntryID)
                            {
                                parentfolder.Items.Remove(i);

                            }
                        }
                        catch (Exception)
                        {


                        }
                    }
                }

            }
            catch (Exception)
            {


            }

            return isDroppedItemUplaoded;
        }


        /// <summary>
        /// <c>NonDocMoveReportItem</c>  member function
        /// move the ReportItem back to source folder if the destination folder is not mapped to sharepoint
        /// document library
        /// </summary>
        /// <param name="Item"></param>

        private void NonDocMoveReportItem(object Item)
        {
            try
            {


                Outlook.ReportItem m = (Outlook.ReportItem)Item;
                mailitemEntryID = m.EntryID;

                try
                {


                    mailitemEntryID = m.EntryID;

                    string strsubject = m.EntryID;
                    if (string.IsNullOrEmpty(strsubject))
                    {
                        strsubject = "tempomailcopy";
                    }

                    mailitemEntryID = strsubject;

                    string tempFilePath = UserLogManagerUtility.RootDirectory + "\\" + strsubject + ".msg";

                    if (Directory.Exists(UserLogManagerUtility.RootDirectory) == false)
                    {
                        Directory.CreateDirectory(UserLogManagerUtility.RootDirectory);
                    }
                    m.SaveAs(tempFilePath, Outlook.OlSaveAsType.olMSG);


                }
                catch (Exception ex)
                {


                }

                Outlook.MAPIFolder fp = (Outlook.MAPIFolder)m.Parent;
                DoNotMoveInNonDocLib(mailitemEntryID, fp);
            }
            catch (Exception)
            {
            }

        }


        /// <summary>
        /// <c>UserMailDeleteOption</c> member function
        /// this member function deletes the mailitem if the Auto Delete option is checked
        /// </summary>
        /// <param name="relativepath"></param>
        /// <param name="fp"></param>
        private void UserMailDeleteOption(string relativepath, Outlook.MAPIFolder fp)
        {

            try
            {
                string tempFilePath = UserLogManagerUtility.RootDirectory + "\\" + relativepath + ".msg";
                Outlook._Application outlookObj = Globals.ThisAddIn.Application; ;
                Outlook.NameSpace outlookNameSpace = outlookObj.GetNamespace("MAPI");
                Outlook.MailItem mitem = (Outlook.MailItem)outlookNameSpace.OpenSharedItem(tempFilePath);
                XMLLogOptions userOptions = UserLogManagerUtility.GetUserConfigurationOptions();
                Outlook.MailItem oMailItem = (Outlook.MailItem)mitem;

                foreach (Outlook.Folder item in addinExplorer.Session.Folders)
                {
                    bool status = false;
                    foreach (Outlook.MAPIFolder fa in item.Folders)
                    {

                        if (fa.EntryID.Contains(ThisAddIn.FromFolderGuid))
                        {
                            if (userOptions.AutoDeleteEmails == true)
                            {
                                if (ThisAddIn.IsMailItemUploaded == true)
                                {
                                    foreach (Outlook.MAPIFolder df in item.Folders)
                                    {
                                        if (df.Name.ToLower().StartsWith("deleted"))
                                        {
                                            oMailItem.Categories = null;
                                            oMailItem.Move(df);
                                        }
                                    }
                                    //
                                    for (int i = 0; i <= fp.Items.Count; i++)
                                    {
                                        try
                                        {
                                            Outlook.MailItem me = (Outlook.MailItem)fp.Items[i];

                                            if (me.EntryID == mailitemEntryID)
                                            {
                                                fp.Items.Remove(i);

                                            }
                                        }
                                        catch (Exception)
                                        {


                                        }
                                    }
                                    //
                                }
                                else
                                {
                                    oMailItem.Move(fa);
                                }
                            }
                            else
                            {
                                oMailItem.Move(fa);
                            }
                            status = true;
                            break;
                        }

                    }
                    if (status == true)
                    {
                        break;
                    }
                }

                File.Delete(tempFilePath);
            }
            catch (Exception ex)
            {
                try
                {
                    string tempFilePath = UserLogManagerUtility.RootDirectory + "\\" + relativepath + ".msg";
                    if (File.Exists(tempFilePath))
                    {
                        File.Delete(tempFilePath);
                    }
                }
                catch (Exception)
                {
                }
            }

            mailitemEntryID = string.Empty;

        }

        /// <summary>
        /// <c>UserReportItemDeleteOption</c> member function
        /// this member function deletes the Reportitem if the Auto Delete option is checked
        /// </summary>
        /// <param name="relativepath"></param>
        /// <param name="fp"></param>
        private void UserReportItemDeleteOption(string relativepath, Outlook.MAPIFolder fp)
        {

            try
            {
                string tempFilePath = UserLogManagerUtility.RootDirectory + "\\" + relativepath + ".msg";
                Outlook._Application outlookObj = Globals.ThisAddIn.Application; ;
                Outlook.NameSpace outlookNameSpace = outlookObj.GetNamespace("MAPI");
                Outlook.ReportItem mitem = (Outlook.ReportItem)outlookNameSpace.OpenSharedItem(tempFilePath);
                XMLLogOptions userOptions = UserLogManagerUtility.GetUserConfigurationOptions();
                Outlook.ReportItem oMailItem = (Outlook.ReportItem)mitem;

                foreach (Outlook.Folder item in addinExplorer.Session.Folders)
                {
                    bool status = false;
                    foreach (Outlook.MAPIFolder fa in item.Folders)
                    {

                        if (fa.EntryID.Contains(ThisAddIn.FromFolderGuid))
                        {
                            if (userOptions.AutoDeleteEmails == true)
                            {
                                if (ThisAddIn.IsMailItemUploaded == true)
                                {
                                    foreach (Outlook.MAPIFolder df in item.Folders)
                                    {
                                        if (df.Name.ToLower().StartsWith("deleted"))
                                        {
                                            oMailItem.Move(df);
                                        }
                                    }
                                    //
                                    for (int i = 0; i <= fp.Items.Count; i++)
                                    {
                                        try
                                        {
                                            Outlook.ReportItem me = (Outlook.ReportItem)fp.Items[i];

                                            if (me.EntryID == mailitemEntryID)
                                            {
                                                fp.Items.Remove(i);

                                            }
                                        }
                                        catch (Exception)
                                        {


                                        }
                                    }
                                    //
                                }
                                else
                                {
                                    oMailItem.Move(fa);
                                }
                            }
                            else
                            {
                                oMailItem.Move(fa);
                            }
                            status = true;
                            break;
                        }

                    }
                    if (status == true)
                    {
                        break;
                    }
                }

                File.Delete(tempFilePath);
            }
            catch (Exception ex)
            {
                try
                {
                    string tempFilePath = UserLogManagerUtility.RootDirectory + "\\" + relativepath + ".msg";
                    if (File.Exists(tempFilePath))
                    {
                        File.Delete(tempFilePath);
                    }
                }
                catch (Exception)
                {
                }
            }

            mailitemEntryID = string.Empty;

        }

        /// <summary>
        /// <c>DoNotMoveInNonDocLib</c> member function
        /// this method moves the item back to source folder if the item is dragged to folder which is not mapped 
        /// to sharepoint document library. and stops moving in non document library
        /// </summary>
        /// <param name="relativepath"></param>
        /// <param name="fp"></param>
        private void DoNotMoveInNonDocLib(string relativepath, Outlook.MAPIFolder fp)
        {

            try
            {
                string tempFilePath = UserLogManagerUtility.RootDirectory + "\\" + relativepath + ".msg";
                Outlook._Application outlookObj = Globals.ThisAddIn.Application; ;
                Outlook.NameSpace outlookNameSpace = outlookObj.GetNamespace("MAPI");
                Outlook.MailItem oMailItem = (Outlook.MailItem)outlookNameSpace.OpenSharedItem(tempFilePath);

                XMLLogOptions userOptions = UserLogManagerUtility.GetUserConfigurationOptions();



                foreach (Outlook.Folder item in addinExplorer.Session.Folders)
                {

                    bool status = false;
                    foreach (Outlook.MAPIFolder fa in item.Folders)
                    {

                        if (fa.EntryID.Contains(ThisAddIn.FromFolderGuid))
                        {

                            string strpth = oMailItem.Subject;

                            oMailItem.Move(fa);
                            status = true;
                            break;
                        }

                    }
                    if (status == true)
                    {
                        break;
                    }
                }
                File.Delete(tempFilePath);

            }
            catch (Exception ex)
            {
                try
                {
                    string tempFilePath = UserLogManagerUtility.RootDirectory + "\\" + relativepath + ".msg";
                    if (File.Exists(tempFilePath))
                    {

                        File.Delete(tempFilePath);
                    }
                }
                catch (Exception)
                {


                }

            }

            for (int i = 0; i <= fp.Items.Count; i++)
            {
                try
                {
                    Outlook.MailItem me = (Outlook.MailItem)fp.Items[i];

                    if (me.EntryID == mailitemEntryID)
                    {
                        fp.Items.Remove(i);

                    }
                }
                catch (Exception)
                {


                }
            }

            mailitemEntryID = string.Empty;
        }


        #endregion

        # region Properties

        /// <summary>
        /// <c>folderName</c> member field
        /// holds the outlook folder name 
        /// </summary>
        private string folderName;

        /// <summary>
        /// <c>FolderName</c> member property
        /// encapsulates  folderName
        /// </summary>
        public string FolderName
        {
            get { return folderName; }
            set { folderName = value; }
        }

        /// <summary>
        /// <c>isFolderAuthenticated</c> member field of type bool
        /// holds the true/false value to
        /// </summary>
        private bool isFolderAuthenticated = false;


        /// <summary>
        /// <c>IsFolderAuthenticated</c> member property
        /// property to check the folder is authenticated or not
        /// encapsulates isFolderAuthenticated member field
        /// </summary>
        public bool IsFolderAuthenticated
        {
            get { return isFolderAuthenticated; }
            set { isFolderAuthenticated = value; }
        }


        /// <summary>
        /// <c>AttachedFolder</c> member field of type MAPIFolder
        /// it encapsulates  activeDroppingFolder member field
        /// </summary>
        public Outlook.MAPIFolder AttachedFolder
        {
            get { return activeDroppingFolder; }
            set { activeDroppingFolder = value; }
        }


        /// <summary>
        /// <c>GetFolder</c>
        /// it finds the outlook mapi folder within parent folder 
        /// </summary>
        /// <param name="parentfolder"></param>
        /// <param name="foldername"></param>
        /// <returns></returns>
        public static Outlook.MAPIFolder GetFolder(Outlook.MAPIFolder parentfolder, string foldername)
        {
            string foldname = foldername;
            Outlook.MAPIFolder returnedfolder = (Outlook.MAPIFolder)parentfolder;

            try
            {
                returnedfolder = (Outlook.MAPIFolder)parentfolder.Parent;
            }
            catch (Exception ex)
            {

                returnedfolder = (Outlook.MAPIFolder)parentfolder;
            }

            try
            {
                bool result = FolderFound(returnedfolder, foldname);
                if (result == true)
                {
                    return returnedfolder.Folders[foldname];

                }
                else
                {

                    bool found = false;
                    //First Level
                    foreach (Outlook.MAPIFolder item in returnedfolder.Folders)
                    {
                        found = FolderFound(item, foldname);
                        if (found == true)
                        {

                            if (item.Folders[foldname].FolderPath.Contains("\\Deleted Items\\"))
                            {
                                item.Folders[foldername].Delete();
                                // returnedfolder = null;
                                return null;
                            }
                            returnedfolder = item.Folders[foldname];

                            break;

                        }
                        else
                        {
                            //second level
                            foreach (Outlook.MAPIFolder secondlevel in item.Folders)
                            {
                                found = FolderFound(secondlevel, foldname);
                                if (found == true)
                                {

                                    if (secondlevel.Folders[foldname].FolderPath.Contains("\\Deleted Items\\"))
                                    {
                                        secondlevel.Folders[foldername].Delete();
                                        // returnedfolder = null;
                                        return null;
                                    }
                                    returnedfolder = secondlevel.Folders[foldname];
                                    break;
                                }
                                else
                                {
                                    //Third Level
                                    foreach (Outlook.MAPIFolder thirdlevel in secondlevel.Folders)
                                    {
                                        found = FolderFound(thirdlevel, foldname);
                                        if (found == true)
                                        {
                                            //returnedfolder = thirdlevel.Folders[foldname];
                                            if (thirdlevel.Folders[foldname].FolderPath.Contains("\\Deleted Items\\"))
                                            {
                                                thirdlevel.Folders[foldername].Delete();
                                                // returnedfolder = null;
                                                return null;
                                            }
                                            returnedfolder = thirdlevel.Folders[foldname];
                                            break;
                                        }
                                        else
                                        {

                                            // fourth
                                            foreach (Outlook.MAPIFolder fourthlevel in thirdlevel.Folders)
                                            {
                                                found = FolderFound(fourthlevel, foldname);
                                                if (found == true)
                                                {
                                                    //returnedfolder = fourthlevel.Folders[foldname];
                                                    if (fourthlevel.Folders[foldname].FolderPath.Contains("\\Deleted Items\\"))
                                                    {
                                                        fourthlevel.Folders[foldername].Delete();
                                                        // returnedfolder = null;
                                                        return null;
                                                    }
                                                    returnedfolder = fourthlevel.Folders[foldname];
                                                    break;
                                                }
                                                else
                                                {
                                                    //Fifth
                                                    foreach (Outlook.MAPIFolder fifthlevel in fourthlevel.Folders)
                                                    {
                                                        found = FolderFound(fifthlevel, foldname);
                                                        if (found == true)
                                                        {
                                                            // returnedfolder = fifthlevel.Folders[foldname];
                                                            if (fifthlevel.Folders[foldname].FolderPath.Contains("\\Deleted Items\\"))
                                                            {
                                                                fifthlevel.Folders[foldername].Delete();
                                                                // returnedfolder = null;
                                                                return null;
                                                            }
                                                            returnedfolder = fifthlevel.Folders[foldname];
                                                            break;
                                                        }
                                                    }

                                                    //Fifth
                                                }
                                                if (found == true)
                                                {
                                                    break;
                                                }
                                            }


                                            //fourth
                                        }
                                        if (found == true)
                                        {
                                            break;
                                        }
                                    }


                                    //End of thirdlevel
                                }
                                if (found == true)
                                {
                                    break;
                                }
                            }
                            //end of second level
                        }
                        if (found == true)
                        {
                            break;
                        }


                    }
                }




            }
            catch (Exception)
            {


            }


            return returnedfolder;

        }


        /// <summary>
        /// <c>FolderFound</c> member function 
        /// checks wether the folder is found or not based on folder name and parent folder
        /// </summary>
        /// <param name="parentfolder"></param>
        /// <param name="foldername"></param>
        /// <returns></returns>
        public static bool FolderFound(Outlook.MAPIFolder parentfolder, string foldername)
        {
            string folname = foldername;
            bool result = false;
            try
            {
                Outlook.MAPIFolder returnedfolder = parentfolder.Folders[folname];
                result = true;
                //foreach (Outlook.MAPIFolder item in parentfolder.Folders)
                //{
                //    if (item.Name.ToLower().Trim() == foldername.ToLower().Trim())
                //    {
                //        result = true;
                //        break;
                //    }
                //}

            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }


        /// <summary>
        /// <c>GetChildFolder</c> member function
        /// finds the immediate child folder based on folder name
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="foldername"></param>
        /// <returns></returns>
        public static Outlook.MAPIFolder GetChildFolder(Outlook.MAPIFolder parent, string foldername)
        {
            Outlook.MAPIFolder retfolder = null;
            try
            {
                foreach (Outlook.MAPIFolder item in parent.Folders)
                {
                    if (item.Name.Trim() == foldername.Trim())
                    {
                        retfolder = item;
                        break;
                    }
                }
            }
            catch (Exception)
            {


            }
            return retfolder;
        }

        #endregion
    }
}

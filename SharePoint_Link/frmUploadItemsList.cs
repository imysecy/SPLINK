using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Xml;
using System.Net;
using System.IO;
using SharePoint_Link.UserModule;
using SharePoint_Link.Utility;
using System.Drawing;
using Interfaces;
using Utility;

namespace SharePoint_Link
{
    /// <summary>
    ///many code of this class has been modfide and written by written by Joy
    /// it was previously a windows form,but replaced with an user control so to add it in custom taskpane
    /// </summary>
    public partial class frmUploadItemsList : UserControl
    {
        #region Constructor

        /// <summary>
        /// <c>upLoadErrorMessage</c> member field of type string
        /// holds the error message to be displayed
        /// </summary>
        private static string upLoadErrorMessage = string.Empty;
        /// <summary>
        /// code wriiten by Joy
        /// this variable is used to store the upload status
        /// </summary>
        bool isSuccessfullyCompleted;
        /// <summary>
        /// <c>UpLoadErrorMessage</c> member property
        /// encapsulates upLoadErrorMessage member field
        /// </summary>
        public static string UpLoadErrorMessage
        {
            get { return frmUploadItemsList.upLoadErrorMessage; }
            set { frmUploadItemsList.upLoadErrorMessage = value; }
        }

        /// <summary>
        /// <c>frmUploadItemsList</c> default constructor
        /// calls <c>InitializeComponent</c> and also <c>CheckItopiaDirectoryExits</c>
        /// method to check the existence of itopia directory . if it is not available creates it.
        /// </summary>
        
        public frmUploadItemsList()
        {
            InitializeComponent();
            UserLogManagerUtility.CheckItopiaDirectoryExits();
           
        }
        #endregion
        #region Delegate

        //delegate of process() function for the purpose of multithreading

        /// <summary>
        /// <c>DelegateHidePanel</c> delegate to hide/show progress bar
        /// </summary>
        /// <param name="isuploadingcompleted"></param>
        delegate void DelegateHidePanel(bool isuploadingcompleted);


        delegate void UIUpdaterDelegate();

        /// <summary>
        /// <c>DelegateUplaodData</c> 
        /// </summary>
        delegate void DelegateUplaodData();

        /// <summary>
        /// <c>DelegateUploadItemUsingCopyService</c> delegate 
        /// delegate to handle upload item event
        /// </summary>
        /// <param name="item"></param>
        delegate void DelegateUploadItemUsingCopyService(UploadItemsData item);

        /// <summary>
        /// <c>DelegateUpdateGridRow</c> delegate
        /// delegate to handle even for updating file uploading status in grid
        /// </summary>
        /// <param name="uploadResult"></param>
        /// <param name="outMessage"></param>
        /// <param name="outURL"></param>
        /// <param name="currentItem"></param>
        delegate void DelegateUpdateGridRow(bool uploadResult, string outMessage, string outURL, UploadItemsData currentItem);

        int currentRowIndex;
        #endregion

        #region Global Variables

        /// <summary>
        /// <c>uploadingItems</c> member field of type List collection
        /// holds UploadItemsData objects
        /// </summary>
        private List<UploadItemsData> uploadingItems = new List<UploadItemsData>();


        // initialize the web service

        // SPCopyService.Copy copyws;


        // ListWebService.Lists listService;
        /// <summary>
        /// <c>lstwebclass</c>  object  of ListWebClass class 
        /// required to upload files to sharepoint 2007 document library
        /// </summary>
        ListWebClass lstwebclass;

        /// <summary>
        /// <c>cmproperties</c> CommonProperties class object 
        /// holds the uploaded file properties and mapped library properties
        /// </summary>
        CommonProperties cmproperties;

        /// <summary>
        /// <c>m_WC</c> WebClient class object
        /// required toupload files 
        /// </summary>
        WebClient m_WC;


        /// <summary>
        /// <c>m_uploadDocLibraryName</c> member field of type string 
        /// holds document library name
        /// <c>m_sharepointLibraryURL</c> member field of type string 
        /// holds the sharepoint document library url.
        /// </summary>
        string m_uploadDocLibraryName = string.Empty, m_sharepointLibraryURL = string.Empty;

        /// <summary>
        /// <c>m_userName</c> member field of type string 
        /// holds the username of sharepoint site
        /// <c>m_password</c> member field of type string 
        /// holds the user password  required to connect with sharepoint document library
        /// </summary>
        string m_userName = string.Empty, m_password = string.Empty;

        /// <summary>
        /// <c>m_uploadingFileName</c> member field of type string
        /// holds the name of file  being uploaded
        /// <c>m_AuthenticationType</c> member field of type string
        /// holds the authentication type value(manual or domain credentials).
        /// </summary>
        string m_uploadingFileName = string.Empty, m_AuthenticationType = string.Empty;

        /// <summary>
        /// <c>m_LibSite</c> member field of type string
        /// holds the library site name
        /// </summary>
        string m_LibSite = string.Empty;

        /// <summary>
        /// <c>m_isUploadingCompleted</c> member field of type bool
        /// holds the status of the file whether it is uploaded or not.
        /// </summary>
        private bool m_isUploadingCompleted = false;

        /// <summary>
        /// <c>m_uploadFolderNode</c> member field of type  <c>XmlNode</c>
        /// holds the mapped folder connection properties as xml node
        /// </summary>
        XmlNode m_uploadFolderNode = null;

        /// <summary>
        /// <c>sendingRequestCount</c> member field of type int
        /// </summary>
        int sendingRequestCount = 0;

        /// <summary>
        /// <c>receivingRequestCount</c> member field of type int
        /// </summary>
        int receivingRequestCount = 0;

        /// <summary>
        /// <c>m_uploadLibNameFromAllLists</c> member field of type string
        /// not being used currently
        /// </summary>
        string m_uploadLibNameFromAllLists = string.Empty;

        /// <summary>
        /// <c>m_spSiteVersion</c> member field of type string
        /// holds the information whether it is sharepoint 2010 site or sharepoint 2007.
        /// </summary>
        string m_spSiteVersion = string.Empty;
        #endregion

        #region Methods


        /// <summary>
        /// <c>ShowForm</c> member functrion
        /// retrieve the connection properties form configuration file and assign them to member fields
        /// required to upload files to sharepoint mapped document library.
        /// </summary>
        /// <param name="folderName"></param>
        public void ShowForm(string folderName)
        {
            try
            {

                //Add trust security for SSL enable sites

                System.Net.ServicePointManager.CertificatePolicy = new TrustAllCertificatePolicy();

                m_isUploadingCompleted = false;
                HideProgressPanel(m_isUploadingCompleted);
                //code written by Joy]
                //invoke is used to avoid error-cross thread operation not valid
                this.Invoke(new MethodInvoker(delegate
                {
                    lblPleaseWaitMessage.Text = "Please Wait - Uploading Items";
                }));
                this.Refresh();
                //Create instance for the API
                //  copyws = new SPCopyService.Copy();

                //  listService = new SharePoint_Link.ListWebService.Lists();
                lstwebclass = new ListWebClass();
                cmproperties = new CommonProperties();
                ThisAddIn.IsUploadingFormIsOpen = true;
                //get the folder details from xml file 
                XmlNode uploadFolderNode = UserLogManagerUtility.GetSPSiteURLDetails("", folderName);
                if (uploadFolderNode != null)
                {

                    this.dgvUploadImages.Rows.Clear();

                    m_uploadFolderNode = uploadFolderNode;
                    cmproperties.UploadFolderNode = uploadFolderNode;
                    cmproperties.CompletedoclibraryURL = uploadFolderNode.ChildNodes[4].InnerText;

                    m_sharepointLibraryURL = uploadFolderNode.ChildNodes[4].InnerText;
                    m_userName = EncodingAndDecoding.Base64Decode(uploadFolderNode.ChildNodes[0].InnerText);
                    m_password = EncodingAndDecoding.Base64Decode(uploadFolderNode.ChildNodes[1].InnerText);

                    m_uploadDocLibraryName = uploadFolderNode.ChildNodes[3].InnerText; ;
                    m_AuthenticationType = uploadFolderNode.ChildNodes[5].InnerText;
                    try
                    {
                        m_spSiteVersion = uploadFolderNode.ChildNodes[12].InnerText;
                        SPVersionClass.SPSiteVersion = m_spSiteVersion;
                    }
                    catch (Exception)
                    {

                        SPVersionClass.SPSiteVersion = SPVersionClass.SiteVersion.SP2007.ToString();
                    }

                    /////////

                    ////////////////
                    //  copyws.CopyIntoItemsCompleted += new CopyIntoItemsCompletedEventHandler(copyws_CopyIntoItemsCompleted);
                    // listService.GetListItemsCompleted += new SharePoint_Link.ListWebService.GetListItemsCompletedEventHandler(listService_GetListItemsCompleted);


                    if (m_AuthenticationType == "Domain Credentials")
                    {
                        //set windows domain credentionals to service
                        System.Net.ICredentials credentionals = System.Net.CredentialCache.DefaultCredentials;
                        //  copyws.Credentials = credentionals;
                        // listService.Credentials = credentionals;

                        cmproperties.Credentionals = credentionals;
                    }
                    else
                    {
                        //Pass username/pwd to network credentionals
                        System.Net.NetworkCredential credentionals = new System.Net.NetworkCredential(m_userName, m_password);
                        // copyws.Credentials = credentionals;
                        // listService.Credentials = credentionals;

                        cmproperties.Credentionals = credentionals;
                        cmproperties.UserName = m_userName;
                        cmproperties.Password = m_password;

                    }
                    m_uploadDocLibraryName = m_uploadFolderNode.ChildNodes[3].InnerText;

                    //Get site URL
                    if (m_sharepointLibraryURL.Contains(m_uploadDocLibraryName) == true)
                    {
                        m_LibSite = m_sharepointLibraryURL.Substring(0, m_sharepointLibraryURL.IndexOf(m_uploadDocLibraryName));

                        //Assign URL to services
                        //  copyws.Url = m_LibSite + @"_vti_bin/Copy.asmx";
                        //  listService.Url = m_LibSite + @"_vti_bin/Lists.asmx";
                        cmproperties.CopyServiceURL = m_LibSite + @"_vti_bin/Lists.asmx";


                        //this.ParentForm.Opacity = 0.8;
                        this.Show();
                        this.Refresh();

                        cmproperties.UploadDocLibraryName = m_uploadDocLibraryName;
                        cmproperties.LibSite = m_LibSite;

                        bool result = lstwebclass.GetSPListName(cmproperties); // GetSPListName();
                        ////////////SP2010 written by mushtaq ahmad////
                        //CommonProperties cmnprop = new CommonProperties();
                        //cmnprop.UploadDocLibraryName = m_uploadDocLibraryName + "sip";
                        //cmnprop.SharepointLibraryURL = m_LibSite;
                        //bool test = Utility.ClientObjectClass.GetSPListName(cmnprop);
                        /////////////////////////End of SP2010/////////////////////////

                        if (result == false)
                        {
                            // this.Close();
                        }
                    }

                    //////////////


                    /////////////
                }
                else
                {
                    EncodingAndDecoding.ShowMessageBox("ShowForm", "Unable to find folder", MessageBoxIcon.Error);
                }

            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("401: Unauthorized") == true)
                {
                    string outMessage = ex.Message + "\r\nYou may get this error if the provided credentails to the folder are wrong.";
                    EncodingAndDecoding.ShowMessageBox("ShowForm", outMessage, MessageBoxIcon.Error);
                }
                else
                {
                    EncodingAndDecoding.ShowMessageBox("ShowForm", ex.Message, MessageBoxIcon.Error);
                }

            }

        }

        /// <summary>
        /// <c>UpdataGridRows</c> member function
        /// Method to add uploaded items to grid along with their status whether they are uploaded 
        /// to sharepoint mapped document library not not.
        /// </summary>
        /// <param name="uploadResult"></param>
        /// <param name="outMessage"></param>
        /// <param name="outURL"></param>
        /// <param name="currentItem"></param>
        public void UpdataGridRows(bool uploadResult, string outMessage, string outURL, UploadItemsData currentItem)
        {
            try
            {
                ThisAddIn.IsMailItemUploaded = uploadResult;
               
                DateTime dtStart = DateTime.Now;
                //Add item to grid
                //code written by Joy]
                //invoke is used to avoid error-cross thread operation not valid
                this.dgvUploadImages.Invoke(new UIUpdaterDelegate(()=>
               {
                      currentRowIndex = this.dgvUploadImages.Rows.Add(SharePoint_Link.Properties.Resources.PROCESS_READY, "", "Waiting....", currentItem.UploadFileName, 0, currentItem.DisplayFolderName);
                }));

                
               

                //simulate two states of current step: success or failure
                if (uploadResult == true)
                {
                    //code written by Joy]
                    //invoke is used to avoid error-cross thread operation not valid
                     this.Invoke(new MethodInvoker(delegate

                    {

                        lblPleaseWaitMessage.Text = "Upload completed for " + currentItem.UploadFileName;

                    }));

                     //code written by Joy]
                     //invoke is used to avoid error-cross thread operation not valid
                   
                         this.dgvUploadImages.Invoke(new UIUpdaterDelegate(()=>
               {
                    this.dgvUploadImages.Rows[currentRowIndex].Cells["colStatusImage"].Value = SharePoint_Link.Properties.Resources.ok_16;
               }));
                   if (string.IsNullOrEmpty(outMessage))
                    {
                        //code written by Joy]
                        //invoke is used to avoid error-cross thread operation not valid
                        this.dgvUploadImages.Invoke(new UIUpdaterDelegate(()=>
                            {
                        this.dgvUploadImages.Rows[currentRowIndex].Cells["colCurrentStatus"].Value = "" +
                            "Unable to read uploaded item inforamtion from list." +
                            "List Name contains special characters.";

                        //Uploading completed sucessfully.There is a problem with reading item from list.";
                        this.dgvUploadImages.Rows[currentRowIndex].Cells["colEdit"].Value = "View";
                        this.dgvUploadImages.Rows[currentRowIndex].Cells["colEdit"].Tag = outURL;
                            }));

                    }
                    else
                    {
                        //code written by Joy]
                        //invoke is used to avoid error-cross thread operation not valid
                         this.dgvUploadImages.Invoke(new UIUpdaterDelegate(()=>
                            {
                        this.dgvUploadImages.Rows[currentRowIndex].Cells["colCurrentStatus"].Value = "Upload Completed.";
                        this.dgvUploadImages.Rows[currentRowIndex].Cells["colEdit"].Value = "MetaTags";
                        this.dgvUploadImages.Rows[currentRowIndex].Cells["colEdit"].Tag = outURL;
                            }));
                    }
                    //Change the file name with modified file name
                    if (currentItem.UploadType == TypeOfUploading.Attachment)
                    {
                        //code written by Joy]
                        //invoke is used to avoid error-cross thread operation not valid
                         this.dgvUploadImages.Invoke(new UIUpdaterDelegate(()=>
                            {
                        this.dgvUploadImages.Rows[currentRowIndex].Cells["colMailSubject"].Value = currentItem.UploadFileName;// currentItem.UploadFileName;
                            }));
                         }
                    else
                    {
                        //code written by Joy]
                        //invoke is used to avoid error-cross thread operation not valid
                          this.dgvUploadImages.Invoke(new UIUpdaterDelegate(()=>
                            {
                        this.dgvUploadImages.Rows[currentRowIndex].Cells["colMailSubject"].Value = currentItem.MailSubject;// currentItem.UploadFileName;
                            }));
                            }

                    //Update uploaded time in xml file
                    UserLogManagerUtility.UpdateFolderConfigNodeDetails(currentItem.DisplayFolderName, "LastUpload", DateTime.Now.ToString());
                }
                else
                {
                    //code written by Joy]
                    //invoke is used to avoid error-cross thread operation not valid
                   // lblPleaseWaitMessage.Text = "Upload failed for " + currentItem.UploadFileName;
                    this.dgvUploadImages.Invoke(new UIUpdaterDelegate(()=>
                            {
                    this.dgvUploadImages.Rows[currentRowIndex].Cells["colStatusImage"].Value = SharePoint_Link.Properties.Resources.exclamation_red;
                    this.dgvUploadImages.Rows[currentRowIndex].Cells["colCurrentStatus"].Value = "Error ::" + outMessage;
                            }));
                }
                //code written by Joy]
                //invoke is used to avoid error-cross thread operation not valid
                this.dgvUploadImages.Invoke(new UIUpdaterDelegate(()=>
                            {
                this.dgvUploadImages.Rows[currentRowIndex].Cells["colElapsedTime"].Value = ((TimeSpan)(DateTime.Now - currentItem.ElapsedTime)).TotalSeconds.ToString("#,##0.00");// +" secs";// ((TimeSpan)(DateTime.Now - dtStart)).TotalSeconds.ToString("#,##0.00") + " secs";
                            }));
                 this.Invoke(new MethodInvoker(delegate
                  {
                this.Refresh();
                  }));
                float elapsedtime = 0.0F;
                //code written by Joy]
                //invoke is used to avoid error-cross thread operation not valid
                 this.Invoke(new MethodInvoker(delegate
                  {
                foreach (DataGridViewRow row in dgvUploadImages.Rows)
                {
                    try
                    {
                        elapsedtime += float.Parse(row.Cells["colElapsedTime"].FormattedValue.ToString());
                    }
                    catch (Exception)
                    {

                    }

                }
                  }));
                  this.Invoke(new MethodInvoker(delegate

                    {

                lblTimeElapsed.Text = "Time Elapsed: " + elapsedtime.ToString();
                    }));
            }

            catch (Exception ex)
            {
                EncodingAndDecoding.ShowMessageBox("UpdateGridRow", ex.Message, MessageBoxIcon.Error);
            }

            if (SPVersionClass.SPSiteVersion != SPVersionClass.SiteVersion.SP2010.ToString())
            {

                if (receivingRequestCount == sendingRequestCount && m_isUploadingCompleted == true)
                {
                    DelegateHidePanel phide = new DelegateHidePanel(HideProgressPanel);
                    this.Invoke(phide, new object[] { m_isUploadingCompleted });
                }

            }



        }

        /// <summary>
        /// <c>ti_tick</c> event handler
        /// calls <c>HideProgressPanel</c> member function to hide the progress bar
        /// Method to show/hide the control on forms
        /// </summary>
        /// <param name="value"></param>
       ////////////////////////////updated by Joy on 30.07.2012/////////////////////////////////
        
        protected void ti_tick(object sender, EventArgs e)
        {

            HideProgressPanel(true);
            Timer t_stop = sender as Timer;
            t_stop.Stop();
            
        }
        ////////////////////////////updated by Joy on 30.07.2012/////////////////////////////////
        /// <summary>
        /// <c>HideProgressPanel</c> member function 
        /// hide or display progress bar based on uploading status.
        /// hide if the uploading status is true otherwise  display
        /// code has been modified by Joy
        /// </summary>
        /// <param name="m_isUploadingCompleted"></param>
        public void HideProgressPanel(bool m_isUploadingCompleted)
        {


            if (m_isUploadingCompleted == true)
            {
                //code written by Joy
                //invoke is used to avoid error-cross thread operation not valid

                //  pictureBox1.Visible = false;
                //  lblPleaseWaitMessage.Visible = false;
                //  dgvUploadImages.Enabled = true;

                //this.Opacity = 1;

                //  dgvUploadImages.Visible = true;
                //code written by Joy]
                //invoke is used to avoid error-cross thread operation not valid
                 this.Invoke(new MethodInvoker(delegate

                    {
                pictureBox1.Visible = false;
                lblPleaseWaitMessage.Visible = false;
                dgvUploadImages.Enabled = true;
                ////this.Opacity = 1;
                //dgvUploadImages.Visible = true;
                this.Text = "Uploading Status: Completed ";
                    }));

            }
            else
            {
                //code written by Joy]
                //invoke is used to avoid error-cross thread operation not valid
                  this.Invoke(new MethodInvoker(delegate

                    {
                pictureBox1.Visible = true;
                lblPleaseWaitMessage.Visible = true;
                dgvUploadImages.Enabled = true; // false;
                dgvUploadImages.Visible = true; // false;

                //IntPtr Hicon = Properties.Resources.wait.GetHicon();
                //this.Icon = System.Drawing.Icon.FromHandle(Hicon);
                this.Text = "Uploading Status: In progress ... ";
                this.Text = "Uploading Status: Completed ";
                    }));

            }
        }



        //private Boolean GetSPListName()
        //{
        //    XmlNode listResponse = null;
        //    try
        //    {
        //        //Get the list details by its name
        //        listResponse = listService.GetList(m_uploadDocLibraryName);

        //    }
        //    catch (Exception ex)
        //    {
        //        m_uploadDocLibraryName = string.Empty;
        //        if (ex.Message.Contains("HTTP status 401: Unauthorized."))
        //        {
        //            return false;
        //        }
        //        else
        //        {
        //            try
        //            {

        //                //If it get exception get all lists information and then get the list 
        //                //using its comparision method.
        //                XmlNode allLists = listService.GetListCollection();
        //                XmlDocument allListsDoc = new XmlDocument();
        //                allListsDoc.LoadXml(allLists.OuterXml);
        //                // allListsDoc.Save(@"c:\allListsDoc.xml"); // for debug
        //                XmlNamespaceManager ns = new XmlNamespaceManager(allListsDoc.NameTable);
        //                ns.AddNamespace("d", allLists.NamespaceURI);

        //                // now get the GUID of the document library we are looking for
        //                //XmlNode dlNode = allListsDoc.SelectSingleNode("/d:Lists/d:List[@Title='" + documentLibraryName + "']", ns);
        //                //Employee[starts-with(FirstName,'Kl')]"
        //                XmlNodeList xlist = allListsDoc.SelectNodes("/d:Lists/d:List[starts-with(@DefaultViewUrl,'" + m_uploadFolderNode.ChildNodes[11].InnerText + "')]", ns);
        //                if (xlist.Count > 0)
        //                {
        //                    string viewURL = xlist[0].Attributes["DefaultViewUrl"].Value;
        //                    if (viewURL.StartsWith(m_uploadFolderNode.ChildNodes[11].InnerText))
        //                    {
        //                        m_uploadDocLibraryName = xlist[0].Attributes["Title"].Value;
        //                        //m_uploadLibNameFromAllLists = xlist[0].Attributes["Title"].Value;
        //                    }

        //                }

        //            }
        //            catch (Exception ex1)
        //            {

        //            }
        //        }
        //    }
        //    if (string.IsNullOrEmpty(m_uploadDocLibraryName))
        //    {
        //        //ShowMessageBox("Unable to find document library name.", MessageBoxIcon.Warning);
        //        return false;
        //    }
        //    return true;
        //}


        /// <summary>
        /// <c>UploadUsingDelegate</c> event handler 
        /// invokes <c>UploadItemUsingWebClientAPI</c> member function to start 
        /// uploading files to sharepoint mapped document library
        /// code written by Joy
        /// </summary>
        /// <param name="uploadData"></param>
        public void UploadUsingDelegate(UploadItemsData uploadData)
        {
            // HideProgressPanel(false);
            //BackgroundWorker bw = new BackgroundWorker();
            //bw.DoWork += delegate(object sender, DoWorkEventArgs e) { bw_DoWork(sender, e, uploadData); }; 
            //bw.RunWorkerAsync();
            bool successfullyuploaded;
           
            //DelegateUploadItemUsingCopyService pdSteps = new DelegateUploadItemUsingCopyService(UploadItemUsingWebClientAPI);
           
            //this.Invoke(pdSteps, new object[] { uploadData });   
            UploadItemUsingWebClientAPI(uploadData);
            
            //Timer ti = new Timer();
            //ti.Interval = 5000;
            //ti.Start();
            //ti.Tick += new EventHandler(ti_tick);

        }

        /// <summary>
        /// <c>UploadItemUsingWebClientAPI</c> member function
        ///finds uploading file properties and call methods to upload file to sharepoint mapped document library.
        /// Upload items using WEBClinet API
        /// </summary>
        /// <param name="uploadData"></param>
        private void UploadItemUsingWebClientAPI(UploadItemsData uploadData)
        {
            try
            {
              
                System.Net.ServicePointManager.CertificatePolicy = new TrustAllCertificatePolicy();
                
                m_WC = new WebClient();
                m_WC.UploadDataCompleted += new UploadDataCompletedEventHandler(m_WC_UploadDataCompleted);
                m_isUploadingCompleted = false;
                m_WC.Credentials = cmproperties.Credentionals;// listService.Credentials;
                byte[] fileBytes = null;


                //Replace "" with empty
                uploadData.UploadFileName = uploadData.UploadFileName.Replace("\"", " ");

                if (uploadData.UploadType == TypeOfUploading.Mail)
                {
                   
                    //Get mail item

                    string tempFilePath = UserLogManagerUtility.RootDirectory + @"\\temp.msg";

                    if (Directory.Exists(UserLogManagerUtility.RootDirectory) == false)
                    {
                        Directory.CreateDirectory(UserLogManagerUtility.RootDirectory);
                        
                    }
                    string msgbody = "";
                    if (uploadData.TypeOfMailItem == TypeOfMailItem.ReportItem)
                    {
                        Outlook.ReportItem omail = uploadData.UploadingReportItem;
                        omail.SaveAs(tempFilePath, Outlook.OlSaveAsType.olMSG);
                        msgbody = omail.CreationTime.ToString();
                    }
                    else
                    {
                       
                        Outlook.MailItem omail = uploadData.UploadingMailItem;
                        
                        omail.SaveAs(tempFilePath, Outlook.OlSaveAsType.olMSG);
                        msgbody = omail.SentOn.ToString();
                       
                    }


                    //// load the file into a file stream
                    //FileStream inStream = File.OpenRead(sourcePath);
                    //byte[] fileBytes = new byte[inStream.Length];
                    //inStream.Read(fileBytes, 0, fileBytes.Length);

                    //Read data to byte
                    fileBytes = File.ReadAllBytes(tempFilePath);
                    


                    ListWebClass.ComputeHashSP07(fileBytes, msgbody);
                   
                    uploadData.MailSubject = HashingClass.Mailsubject;
                   
                    uploadData.UploadFileName = HashingClass.Hashedemailbody.ToString().Trim() + uploadData.UploadFileExtension; // uploadData.UploadFileName.Trim()
                  

                }
                else
                {
                    //Set fullname to filename

                    fileBytes = uploadData.AttachmentData;


                    uploadData.UploadFileName = uploadData.UploadFileName.Trim() + uploadData.UploadFileExtension;



                }

                //string destinationUrls = m_LibSite + m_uploadDocLibraryName + "/" + uploadData.UploadFileName;
                string destinationUrls = m_sharepointLibraryURL.Substring(0, m_sharepointLibraryURL.LastIndexOf("Forms")) + uploadData.UploadFileName;
               
                cmproperties.LibSite = m_LibSite;
               
                cmproperties.UploadDocLibraryName = m_uploadDocLibraryName;
               
                cmproperties.FileBytes = fileBytes;
               
                cmproperties.UserName = m_userName;
                
                cmproperties.Password = m_password;
               
                ///Upload Files to Sharepoint site
                if (SPVersionClass.SPSiteVersion == SPVersionClass.SiteVersion.SP2010.ToString())
                {
                    
                    //for SharePoint2010
                    bool docuploaded = false;
                   
                    string libsitestr = cmproperties.LibSite;
                    
                    docuploaded = ListWebClass.uploadFilestoLibraryUsingClientOM(uploadData, cmproperties);
                    
                    //HideProgressPanel(false);
                    if (docuploaded == true)
                    {
                       
                        m_isUploadingCompleted = true;
                        string pth = libsitestr + cmproperties.UploadDocLibraryName + "/Forms/EditForm.aspx?ID=" + ListWebClass.FileID + "&IsDlg=1&UpFName=" + uploadData.UploadFileName;
                        //DelegateUpdateGridRow pdSteps = new DelegateUpdateGridRow();
                        UpdataGridRows(true, "Success", pth, uploadData);
                        //this.Invoke(pdSteps, new object[] { true, "Success", pth, uploadData });
                        //////////////////////////modified by joy on 27.07.2012////////////////////////////////////////
                        isSuccessfullyCompleted = true;
                        //code written by Joy
                        //increments the no of uploaded items
                        Globals.ThisAddIn.no_of_items_copied++;
                        Globals.ThisAddIn.no_of_moved_item_uploaded++;
                        Globals.ThisAddIn.no_of_copied_item_uploaded++;
                        Globals.ThisAddIn.no_of_t_item_uploaded++;
                        
                    }
                    else
                    {
                        //////////////////////////modified by joy on 27.07.2012////////////////////////////////////////
   
                        m_isUploadingCompleted = false;
                        UpdataGridRows(false, "Exception While Uploading." + UpLoadErrorMessage, " ", uploadData);
                       
                        //////////////////////////modified by joy on 27.07.2012////////////////////////////////////////
   
                        isSuccessfullyCompleted = false;

                    }
                    

                    //////////////////////////modified by joy on 30.07.2012////////////////////////////////////////
                    Timer ti = new Timer();
                    ti.Interval = 2000;
                    ti.Start();
                    //code written by Joy]
                    //invoke is used to avoid error-cross thread operation not valid
                     this.Invoke(new MethodInvoker(delegate

                    {
                    lblPleaseWaitMessage.Text = "  Upload completed  ";
                    }));
                    ti.Tick += new EventHandler(ti_tick);
                    HideProgressPanel(true);
                    //////////////////////////modified by joy on 30.07.2012////////////////////////////////////////

                }
                else
                {
                    if (SPVersionClass.SPSiteVersion == SPVersionClass.SiteVersion.SP2007.ToString())
                    {
                        // for sharepoint2007


                        uploadData.MailSubject = HashingClass.Mailsubject;
                        uploadData.ModifiedDate = HashingClass.Modifieddate;
                        m_WC.UploadDataAsync(new Uri(destinationUrls), "PUT", fileBytes, uploadData);





                    }
                    else
                    {
                        EncodingAndDecoding.ShowMessageBox("Version Conflick", "Cannot detect Sharepoint version", MessageBoxIcon.Error);
                        HideProgressPanel(true);
                    }
                }






                //try
                //{
                //    bool result = listService.CheckInFile(destinationUrls, "", "0");
                //}
                //catch { }

                sendingRequestCount++;


            }
            catch (Exception ex)
            {

                if (ex.Message.Contains("401: Unauthorized") == true)
                {
                    string outMessage = ex.Message + "\r\nYou may get this error if the provided credentails to the folder are wrong.";
                    EncodingAndDecoding.ShowMessageBox("WebClient API Calling", outMessage, MessageBoxIcon.Information);
                }
                else
                {
                    EncodingAndDecoding.ShowMessageBox("WebClient API Calling", ex.Message, MessageBoxIcon.Error);
                }

            }


        }

        /// <summary>
        /// code written by Joy
        /// this method returns the upload status for a currently uploading mail item
        /// </summary>
        public bool IsSuccessfullyUploaded
        {
            get
            {
              return  isSuccessfullyCompleted;
            }
        }
        
        
/// <summary>
        /// <c>UploadItemUsingCopyService</c>
        /// Upload Items using copy web service API
        /// calls <c>SPCopyClass.UploadItemUsingCopyService(</c> method to start uploading file to sharepoint 2007 mappped sharepoint
        /// document library.
        /// </summary>
        /// <param name="uploadData"></param>
        private void UploadItemUsingCopyService(UploadItemsData uploadData)
        {
            try
            {
                System.Net.ServicePointManager.CertificatePolicy = new TrustAllCertificatePolicy();

                CommonProperties comproperties = new CommonProperties();
                comproperties.LibSite = m_LibSite;
                comproperties.UploadDocLibraryName = m_uploadDocLibraryName;

                m_isUploadingCompleted = false;


                //oldcode//
                //byte[] fileBytes = null;
                //string[] destinationUrls = null;

                ////Replace "" with empty
                //uploadData.UploadFileName = uploadData.UploadFileName.Replace("\"", " ");
                //if (uploadData.UploadType == TypeOfUploading.Mail)
                //{

                //    //Get mail item

                //    string tempFilePath = UserLogManagerUtility.RootDirectory + @"\\temp.msg";
                //    if (Directory.Exists(UserLogManagerUtility.RootDirectory) == false)
                //    {
                //        Directory.CreateDirectory(UserLogManagerUtility.RootDirectory);
                //    }
                //    Outlook.MailItem omail = uploadData.UploadingMailItem;
                //    omail.SaveAs(tempFilePath, Outlook.OlSaveAsType.olMSG);

                //    //// load the file into a file stream

                //    //Read data to byte
                //    fileBytes = File.ReadAllBytes(tempFilePath);
                //    uploadData.UploadFileName = uploadData.UploadFileName.Trim() + uploadData.UploadFileExtension;
                //}
                //else
                //{
                //    //Set fullname to filename
                //    fileBytes = uploadData.AttachmentData;
                //    uploadData.UploadFileName = uploadData.UploadFileName.Trim() + uploadData.UploadFileExtension;
                //}


                //// format the destination URL
                //destinationUrls = new string[] { m_LibSite + m_uploadDocLibraryName + "/" + uploadData.UploadFileName };
                ////destinationUrls = new string[] { m_sharepointLibraryURL.Substring(0,m_sharepointLibraryURL.LastIndexOf("Forms")-1)  + "/" + uploadData.UploadFileName };

                //// to specify the content type
                //FieldInformation ctInformation = new FieldInformation();
                //ctInformation.DisplayName = "Content Type";
                //ctInformation.InternalName = "ContentType";
                //ctInformation.Type = FieldType.Choice;
                //ctInformation.Value = "Your content type";



                ////FieldInformation[] metadata = { titleInformation };
                //FieldInformation[] metadata = { };

                //// execute the CopyIntoItems method
                //copyws.CopyIntoItemsAsync("OutLook", destinationUrls, metadata, fileBytes, uploadData);
                //// End of oldcode//
                if (SPCopyClass.UploadItemUsingCopyService(uploadData, comproperties))
                {
                    sendingRequestCount++;
                }


            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("401: Unauthorized") == true)
                {
                    string outMessage = ex.Message + "\r\nYou may get this error if the provided credentails to the folder are wrong.";
                    EncodingAndDecoding.ShowMessageBox("UploadCopyCalling", outMessage, MessageBoxIcon.Information);
                }
                else
                {
                    EncodingAndDecoding.ShowMessageBox("UploadCopyCalling", ex.Message, MessageBoxIcon.Error);
                }

            }

        }


        /// <summary>
        /// <c>m_WC_UploadDataCompleted</c> Event handler
        /// WebCLinet Async completed event
        /// executes oncompleting uploading and call method to update grid row. and 
        /// GetListItemsAsyncWrapper method to complete the process of uploading in sharepoint 2007.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void m_WC_UploadDataCompleted(object sender, UploadDataCompletedEventArgs e)
        {
            try
            {
                int test = sendingRequestCount;
                string outMessage = string.Empty;
                string sucessURL = string.Empty;
                UploadItemsData uploadData = (UploadItemsData)e.UserState;
                try
                {
                    receivingRequestCount++;

                    // execute the CopyIntoItems method


                    //Set the file name to global variable
                    m_uploadingFileName = uploadData.UploadFileName;
                    if (e.Error != null)
                    {
                        Exception ex = e.Error;
                        if (Convert.ToString(ex.Message) == "Object reference not set to an instance of an object.")
                        {
                            outMessage = "Cause :: No prermission to access the List.(OR) List or Library is deleted or moved";
                        }
                        else if (Convert.ToString(ex.Message) == "InvalidUrl")
                        {
                            outMessage = "Upload Falied.Filename contains some special characters." + ex.Message;
                        }
                        else
                        {
                            outMessage = ex.Message;
                        }
                        m_isUploadingCompleted = true;
                        UpdataGridRows(false, outMessage, sucessURL, uploadData);
                      

                        //return false;
                    }
                    else
                    {
                        Byte[] result = e.Result;

                        XmlDocument xmlDoc = new XmlDocument();
                        XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");

                        XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");
                        XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");

                        ndQueryOptions.InnerXml = "<IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>" +
                            "<DateInUtc>TRUE</DateInUtc>";
                        //send required fields information
                        ndViewFields.InnerXml = "<FieldRef Name='ID'/> <FieldRef Name='LinkTitle'/>";

                        ndQuery.InnerXml = "<Where><Eq><FieldRef Name='FileLeafRef'></FieldRef><Value Type='Text'>" + uploadData.UploadFileName + "</Value></Eq></Where>";

                        //Call items details ASYNC
                        // listService.GetListItemsAsync(m_uploadDocLibraryName, null, ndQuery, ndViewFields, "2", ndQueryOptions, null, uploadData);
                        //  lstwebclass.GetListItemsAsync(m_uploadDocLibraryName, null, ndQuery, ndViewFields, "2", ndQueryOptions, null, uploadData, cmproperties);
                        if (SPVersionClass.SPSiteVersion == SPVersionClass.SiteVersion.SP2007.ToString())
                        {
                            GetListItemsAsyncWrapper(m_uploadDocLibraryName, null, ndQuery, ndViewFields, "2", ndQueryOptions, null, uploadData, cmproperties);
                        }
                        m_isUploadingCompleted = false;
                        // HideProgressPanel(true);


                    }

                }
                catch (Exception ex)
                {
                    outMessage = ex.Message;
                    if (ex.InnerException != null)
                    {
                        outMessage = ex.InnerException.Message;
                    }
                    m_isUploadingCompleted = true;
                    DelegateUpdateGridRow pdSteps = new DelegateUpdateGridRow(UpdataGridRows);
                    this.Invoke(pdSteps, new object[] { false, "Exception in CopyIntoItemsCompleted Event." + outMessage, sucessURL, uploadData });
                    // ShowMessageBox(ex.Message, MessageBoxIcon.Error);
                }


            }
            catch (Exception ex)
            {
                EncodingAndDecoding.ShowMessageBox("UploadDataAsync", ex.Message, MessageBoxIcon.Error);
            }

        }

        /// <summary>
        /// Copy web serice UplaoddataCOmpleted Async event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //void copyws_CopyIntoItemsCompleted(object sender, CopyIntoItemsCompletedEventArgs e)
        //{

        //    try
        //    {
        //        UploadItemsData uploaditemdata = (UploadItemsData)e.UserState;
        //        CopyResult[] result = e.Results;
        //        int resultlength = result.Length;
        //        string outMessage = null;
        //        if (result.Length > 0)
        //        {

        //            if (result[0].ErrorMessage != null)
        //            {

        //                if (Convert.ToString(result[0].ErrorMessage) == "Object reference not set to an instance of an object.")
        //                {
        //                    outMessage = "Cause :: No prermission to access the List.(OR) List or Library is deleted or moved";
        //                }
        //                else if (Convert.ToString(result[0].ErrorCode) == "InvalidUrl")
        //                {
        //                    outMessage = "Upload Falied.Filename contains some special characters." + result[0].DestinationUrl;
        //                }
        //                else
        //                {
        //                    outMessage = result[0].ErrorMessage;
        //                }

        //            }

        //        }
        //        CopyIntoItemCompleted(uploaditemdata, resultlength, outMessage);

        //    }
        //    catch (Exception)
        //    {


        //    }


        //}


        /// <summary>
        /// <c>CopyIntoItemCompleted</c> member function
        /// calls <c>UpdataGridRows</c> method to update rows in datagrid.
        /// and calls <c>GetListItemsAsyncWrapper</c> method to complete uploading process
        /// </summary>
        /// <param name="uploaditemdata"></param>
        /// <param name="resultlength"></param>
        /// <param name="oMessage"></param>
        public void CopyIntoItemCompleted(UploadItemsData uploaditemdata, int resultlength, string oMessage)
        {
            try
            {

                string outMessage = string.Empty;
                string sucessURL = string.Empty;
                UploadItemsData uploadData = uploaditemdata;
                try
                {
                    receivingRequestCount++;

                    // execute the CopyIntoItems method



                    if (resultlength > 0)
                    {
                        //Set the file name to global variable
                        m_uploadingFileName = uploadData.UploadFileName;
                        if (oMessage != null)
                        {

                            outMessage = oMessage;

                            m_isUploadingCompleted = true;
                            UpdataGridRows(false, outMessage, sucessURL, uploadData);
                           

                            //return false;
                        }
                        else
                        {

                            XmlDocument xmlDoc = new XmlDocument();
                            XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");

                            XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");
                            XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");

                            ndQueryOptions.InnerXml = "<IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>" +
                                "<DateInUtc>TRUE</DateInUtc>";
                            //send required fields information
                            ndViewFields.InnerXml = "<FieldRef Name='ID'/> <FieldRef Name='LinkTitle'/>";

                            ndQuery.InnerXml = "<Where><Eq><FieldRef Name='FileLeafRef'></FieldRef><Value Type='Text'>" + uploadData.UploadFileName + "</Value></Eq></Where>";


                            //   lstwebclass.GetListItemsAsync(m_uploadDocLibraryName, null, ndQuery, ndViewFields, "2", ndQueryOptions, null, uploadData, cmproperties);

                            if (SPVersionClass.SPSiteVersion == SPVersionClass.SiteVersion.SP2007.ToString())
                            {
                                GetListItemsAsyncWrapper(m_uploadDocLibraryName, null, ndQuery, ndViewFields, "2", ndQueryOptions, null, uploadData, cmproperties);
                            }
                            m_isUploadingCompleted = false;


                        }
                    }
                }
                catch (Exception ex)
                {
                    outMessage = ex.Message;
                    if (ex.InnerException != null)
                    {
                        outMessage = ex.InnerException.Message;
                    }
                    m_isUploadingCompleted = true;
                    UpdataGridRows(false, "Exception in CopyIntoItemsCompleted Event." + outMessage, sucessURL, uploadData);
                   
                    // ShowMessageBox(ex.Message, MessageBoxIcon.Error);
                }

            }
            catch (Exception ex)
            {
                EncodingAndDecoding.ShowMessageBox("CopyAsyncCompleted", ex.Message, MessageBoxIcon.Error);
            }
        }




        #endregion

        #region Form Events

        /// <summary>
        /// <c>frmProgress_FormClosing</c> event handler
        /// cancel the Itopia progress window form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        ////modified code on 20.07.2012-Joy
        ///
        ////////////////////////////////////////////////////////////////////////////
        //private void frmProgress_FormClosing(object sender, FormClosingEventArgs e)
        //{
        //    //Check is there any message is in process, cancel the form closing.
        //    if (m_isUploadingCompleted == false)
        //    {
        //        e.Cancel = true;
        //    }
        //    else
        //    {
        //        ThisAddIn.IsUploadingFormIsOpen = false;
        //    }
        //}
        //////////////////////////////////////////////////////////////////////////////
        ///
        ////modified code on 20.07.2012-Joy
        #endregion


        /// <summary>
        /// <c>dgvUploadImages_CellClick</c> Event Handler
        /// this opens the metatag window form for the file uploaded to sharepoint mapped document library.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvUploadImages_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 1 && e.RowIndex >= 0)
                {
                    string s = Convert.ToString(((System.Windows.Forms.DataGridView)(sender)).CurrentCell.Tag);
                    if (string.IsNullOrEmpty(s) == false)
                    {
                        frmEditUploadedItem obj = new frmEditUploadedItem();
                        obj.ShowWithBrowser(m_userName, m_password, s);
                        obj.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                EncodingAndDecoding.ShowMessageBox("Cell Click", ex.Message, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// <c>toolStripSplitButtonClose_ButtonClick</c> Event Handler
        /// close the window form and cancel uploading files
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        ////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Modified by joy on 21.07.2012
        /// </summary>

        ////////////////////////////////////////////////////////////////////////////
        private void toolStripSplitButtonClose_ButtonClick(object sender, EventArgs e)
        {
            if (m_isUploadingCompleted == false)
            {
                DialogResult result = MessageBox.Show("Closing window will run the upload in background.Do you want to close this window?", "ITOPIA", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (result == DialogResult.OK)
                {
                    m_isUploadingCompleted = true;
                    Globals.ThisAddIn.CustomTaskPanes[0].Visible = false;
                }
            }
            else
            {
                Globals.ThisAddIn.CustomTaskPanes[0].Visible = false;
            }

        }
        ////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Modified by joy on 21.07.2012
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// //////////////////////////////////////////////////////////////////////

        /// <summary>
        /// <c>frmUploadItemsList_Resize</c> Event Handler
        /// set height and width of Picture box and message location
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// ///////////////////////////////////////////////////////////////
        /// ///////////////Modified code by Joy on 20.07.2012////////////// 
        /// ///////////////////////////////////////////////////////////////
        private void frmUploadItemsList_Resize(object sender, EventArgs e)
        {
            try
            {
                Control control = (Control)sender;
                int height = control.Size.Height / 2;
                int width = control.Size.Width / 3;
                pictureBox1.Location = new Point(width + 60, height - 100);
                lblPleaseWaitMessage.Location = new Point(width, height);

            }
            catch (Exception ex)
            { }
        }
        /// ///////////////////////////////////////////////////////////////
        /// ///////////////Modified code by Joy on 20.07.2012////////////// 
        /// ///////////////////////////////////////////////////////////////
        

        /// <summary>
        /// <c>GetListItemsAsyncWrapper</c> wrapper function
        /// assign properties to  <c>ListWebService.Lists</c> class object
        /// and Register <c>GetListItemsCompleted</c> event of object
        /// it finally calls <c> listService.GetListItemsAsyn</c> function to sync uploading file
        /// </summary>
        /// <param name="m_uploadDocLibraryName"></param>
        /// <param name="viewname"></param>
        /// <param name="ndQuery"></param>
        /// <param name="ndViewFields"></param>
        /// <param name="rowlimit"></param>
        /// <param name="ndQueryOptions"></param>
        /// <param name="webid"></param>
        /// <param name="uploadData"></param>
        /// <param name="property"></param>
        private void GetListItemsAsyncWrapper(string m_uploadDocLibraryName, string viewname, XmlNode ndQuery, XmlNode ndViewFields, string rowlimit, XmlNode ndQueryOptions, string webid, object uploadData, CommonProperties property)
        {
            try
            {

                // for 2007//
                ListWebService.Lists listService = new ListWebService.Lists();
                listService.Credentials = property.Credentionals;
                listService.Url = property.CopyServiceURL;
                listService.GetListItemsCompleted += new ListWebService.GetListItemsCompletedEventHandler(listService_GetListItemsCompleted);
                // listService.GetListItemsAsync(m_uploadDocLibraryName, null, ndQuery, ndViewFields, "2", ndQueryOptions, null, uploadData);

                XmlNode n = listService.GetListCollection();


                string urlroot = property.LibSite.Remove(property.LibSite.Length - 1);
                foreach (XmlNode item in n.ChildNodes)
                {
                    string strcompare = urlroot + item.Attributes["DefaultViewUrl"].Value;
                    if (strcompare == property.CompletedoclibraryURL)
                    {
                        m_uploadDocLibraryName = item.Attributes["Title"].Value;
                        cmproperties.UploadDocLibraryName = m_uploadDocLibraryName;
                        break;

                    }

                }
                listService.GetListItemsAsync(m_uploadDocLibraryName, null, ndQuery, ndViewFields, "2", ndQueryOptions, null, uploadData);


            }
            catch (Exception ex)
            {
            }

        }


        /// <summary>
        /// <c>listService_GetListItemsCompleted</c> event handler
        /// calls <c>ListItemCompleted</c> method
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void listService_GetListItemsCompleted(object sender, SharePoint_Link.ListWebService.GetListItemsCompletedEventArgs e)
        {

            try
            {
                XmlNode xmlnode = e.Result;
                UploadItemsData uploaditemdata = (UploadItemsData)e.UserState;
                ListItemCompleted(xmlnode, uploaditemdata);

            }
            catch (Exception ex)
            {
                m_isUploadingCompleted = true;
                DelegateUpdateGridRow pdSteps = new DelegateUpdateGridRow(UpdataGridRows);
                this.Invoke(pdSteps, new object[] { false, "uploaded with error", "", null });

                EncodingAndDecoding.ShowMessageBox("Error occured, may be th document Library does not exists", "Error occured, may be th document Library does not exists", MessageBoxIcon.Error);
            }


        }

        /// <summary>
        /// <c>ListItemCompleted</c> function
        /// gets uploaded file  url and calls <c>UpdataGridRows</c> method to  add new row in grid.
        /// </summary>
        /// <param name="ndvolListItem"></param>
        /// <param name="uitemdata"></param>
        public void ListItemCompleted(XmlNode ndvolListItem, UploadItemsData uitemdata)
        {
            try
            {

                XmlNode ndVolunteerListItems = ndvolListItem;
                string test = uitemdata.UploadFileName;
                if (ndVolunteerListItems != null)
                {
                    if (ndVolunteerListItems.ChildNodes.Count == 3)
                    {
                        if (ndVolunteerListItems.ChildNodes[1].ChildNodes.Count > 1)
                        {
                            string id = Convert.ToString(ndVolunteerListItems.ChildNodes[1].ChildNodes[1].Attributes["ows_ID"].Value);
                            //////////

                            UpdateItemAttributes(id, uitemdata);

                            //////////

                            string sucessURL1 = m_LibSite + m_uploadFolderNode.ChildNodes[3].InnerText + "/Forms/EditForm.aspx?ID=" + id + "&UpFName=" + uitemdata.UploadFileName; ;
                            m_isUploadingCompleted = true;
                            DelegateUpdateGridRow pdSteps = new DelegateUpdateGridRow(UpdataGridRows);
                            this.Invoke(pdSteps, new object[] { true, "Success", sucessURL1, uitemdata });
                            UploadItemsData item = uitemdata; //(UploadItemsData)e.UserState;


                            //try
                            //{
                            //    string fileCheckin = m_sharepointLibraryURL.Substring(0, m_sharepointLibraryURL.LastIndexOf("Forms")) + item.UploadFileName;
                            //    bool b = listService.CheckInFile(fileCheckin, "Uplaod Completed.", "1");
                            //}
                            //catch (Exception ex)
                            //{ 
                            //}
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string outMessage1 = ex.Message;
                if (ex.InnerException != null)
                {
                    outMessage1 = ex.InnerException.Message;
                }
                m_isUploadingCompleted = true;
                DelegateUpdateGridRow pdSteps = new DelegateUpdateGridRow(UpdataGridRows);
                this.Invoke(pdSteps, new object[] { false, "Exception in GetListItemsCompleted Event." + outMessage1, " ", uitemdata });
            }
        }


        /// <summary>
        /// <c>UpdateItemAttributes</c> member function 
        /// updates metatags information of file uploaded to sharepoint 2007 mapped document library.
        /// </summary>
        /// <param name="id"></param>
        /// <param name="uitemdata"></param>
        private void UpdateItemAttributes(string id, UploadItemsData uitemdata)
        {

            try
            {

                ListWebService.Lists listService = new ListWebService.Lists();

                listService.Credentials = cmproperties.Credentionals;
                listService.Url = cmproperties.CopyServiceURL;
                XmlDocument doc = new XmlDocument();
                XmlNode docnode = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
                XmlElement batch_Element = doc.CreateElement("Batch");


                string itm = "<Method ID='1' Cmd='Update'>" + "<Field Name='ID'>" + id + "</Field>" + "<Field Name='Title'>" + uitemdata.MailSubject + "</Field>";
                itm += " </Method>";
                batch_Element.InnerXml = itm;
                try
                {
                    listService.UpdateListItems(cmproperties.UploadDocLibraryName, batch_Element);
                }
                catch (Exception)
                { }

                itm = "";
                itm = "<Method ID='1' Cmd='Update'>" + "<Field Name='ID'>" + id + "</Field>" + "<Field Name='Title'>" + uitemdata.MailSubject + "</Field>";
                itm += "<Field Name='ModifiedDate'>" + uitemdata.ModifiedDate + "</Field></Method>";

                XmlDocument docc = new XmlDocument();
                XmlElement bat_Element = docc.CreateElement("Batch");
                bat_Element.InnerXml = itm;
                listService.UpdateListItems(cmproperties.UploadDocLibraryName, bat_Element);




            }
            catch (Exception ex)
            {
                try
                {
                    ListWebService.Lists listService = new ListWebService.Lists();
                    listService.Credentials = cmproperties.Credentionals;
                    listService.Url = cmproperties.CopyServiceURL;
                    XmlDocument doc = new XmlDocument();
                    XmlNode docnode = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
                    XmlElement batch_Element = doc.CreateElement("Batch");

                    XmlNode FieldNode = doc.CreateElement("Method");
                    XmlAttribute productAttribute = doc.CreateAttribute("ID");
                    productAttribute.Value = "1";
                    FieldNode.Attributes.Append(productAttribute);
                    productAttribute = doc.CreateAttribute("Cmd");
                    productAttribute.Value = "Update";
                    FieldNode.Attributes.Append(productAttribute);

                    XmlNode fnode = doc.CreateElement("Field");
                    XmlAttribute fattribute = doc.CreateAttribute("Name");
                    fattribute.Value = "ID";
                    fnode.InnerText = id;
                    fnode.Attributes.Append(fattribute);
                    FieldNode.AppendChild(fnode);
                    /////////////
                    XmlNode fnodetitle = doc.CreateElement("Field");
                    XmlAttribute fattributetitle = doc.CreateAttribute("Name");
                    fattributetitle.Value = "Title";
                    fnodetitle.InnerText = uitemdata.MailSubject;
                    fnodetitle.Attributes.Append(fattributetitle);
                    FieldNode.AppendChild(fnodetitle);
                    /////////////
                    XmlNode fnodeDate = doc.CreateElement("Field");
                    XmlAttribute dateattribute = doc.CreateAttribute("Name");
                    dateattribute.Value = "ModifiedDate";
                    fnodeDate.InnerText = uitemdata.ModifiedDate;
                    fnodeDate.Attributes.Append(dateattribute);
                    FieldNode.AppendChild(fnodeDate);
                    ///////////

                    batch_Element.AppendChild(FieldNode);
                    doc.AppendChild(batch_Element);


                    try
                    {
                        listService.UpdateListItems(cmproperties.UploadDocLibraryName, batch_Element);
                    }
                    catch
                    { }



                }
                catch
                { }
            }

        }

      
     

       
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using SharePoint_Link.UserModule;
using SharePoint_Link.Utility;
using Utility;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using com.softwarekey.ClientLib.InstantPLUS;
using System.Data;
using System.Data.SqlClient;

namespace SharePoint_Link
{
    /// <summary>
    ///<c>ThisAddIn</c>outLook Addin 2010 main Class. This class handles events and functionalities
    /// through which this addin integrates with outlook and  perform its functionalities
    ///  
    /// </summary>

    public partial class ThisAddIn
    {
        #region Global declarations

        private static bool isAuthorized = false;
        public static bool IsAuthorized { get { return isAuthorized; } }

        public static String CurrentWebUrlLink = "";

        private const string ADXHTMLFileName = "MaPiFolderTemp.htm";
        private const string ADXHTMLFileName2 = "ADXOlFormGeneral.html";
        /// <summary>
        /// <c>OutlookObj</c> Interface member of  <c>Outlook._Application</c>
        /// this  is required  in the addin events.e.g outlook startup event
        /// , folder switching, paste etc
        /// </summary>
        public static Outlook._Application OutlookObj;


        public bool isuploadRunning=false;


        /// <summary>
        /// <c>outlookNameSpace</c> interface member of <c> Outlook.NameSpace</c>
        /// this  is required in in  aadin events
        /// </summary>
        Outlook.NameSpace outlookNameSpace;

        /// <summary>
        /// <c>addinExplorer</c>  Outlook.Explorer interface member required in addin events
        /// </summary>
        public static Outlook.Explorer addinExplorer;

        /// <summary>
        /// <c>oMailRootFolders</c> interface member of <c>Outlook.Folders</c> 
        /// this member is required to interacts with outlook default  folder e.g Inbox
        /// </summary>
        Outlook.Folders oMailRootFolders;


        /// <summary>
        /// <c>oCurrentSelectedFolder</c> interface member of <c>MAPIFolder</c>
        /// this member is required to interacts with the folder which is currently selected
        /// </summary>
        Outlook.MAPIFolder oCurrentSelectedFolder;


        /// <summary>
        /// <c>menuBarItopia</c> interface member of <c>CommandBar</c>
        /// this is required to add buttons on  top ribbon (outlook addin menu bar)
        /// </summary>
        private Office.CommandBar menuBarItopia;

        /// <summary>
        /// <c>menuBarSharePoint</c> interface member of <c>CommandBarButton</c>
        /// menuBarSharePoint  is required to create sharepoint addin tab in outlook menu
        /// </summary>
        private Office.CommandBarButton menuBarSharePoint;


        /// <summary>
        /// <c>btnNewConnection</c> interface member of <c>CommandBarButton</c>
        /// this button is shown as new button link on top menu
        /// </summary>
        private Office.CommandBarButton btnNewConnection;

        /// <summary>
        /// <c>btnConnectionProperties</c> interface member of <c>CommandBarButton</c>
        /// this button is shown as connection properties  link on top menu which will display  all connections
        /// </summary>
        private Office.CommandBarButton btnConnectionProperties;

        /// <summary>
        /// <c>btnOptions</c> interface member of <c>CommandBarButton</c>
        /// this  button is shown as option button link on top menu 
        /// </summary>
        private Office.CommandBarButton btnOptions;


        /// <summary>
        /// <c>strParentMenuName</c>  is a string variable to hold addin title
        /// </summary>
        string strParentMenuName = "SharePoint Link";

        /// <summary>
        /// <c>>strParentMenuTag</c> is a string variable which is not currently being used
        /// </summary>
        string strParentMenuTag = string.Empty;

        /// <summary>
        /// <c>frmSPSiteConfigurationObject</c>  an object of <c>frmSPSiteConfiguration</c> 
        /// this is used to set new connection settings 
        /// </summary>
        frmSPSiteConfiguration frmSPSiteConfigurationObject;

        /// <summary>
        /// <c>frmUploadEmailMessagesObject</c>   an object of  <c>frmUploadItemsList</c>
        /// this is  windows form which displays uploaded files status
        /// </summary>
        public static frmUploadItemsList frmUploadEmailMessagesObject;

        public  frmUploadItemsList frmlistObject;
        /// <summary>
        /// declares custom taskpane as global level
        /// code written by Joy
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        /// <summary>
        //code written by Joy
        /// </summary>
        public int no_of_items_to_be_uploaded=0;
        /// <summary>
        //code written by Joy
        /// </summary>
        public int no_of_items_copied = 0;
        /// <summary>
        /// selection variable used to store the no of selected mail items
        //code written by Joy
        /// </summary>
        public Outlook.Selection oselection;
        /// <summary>
        //code written by Joy
        /// </summary>
        public Form form = null;
        /// <summary>
        ///threading timer variable which is to start timer with 3 minutes interval 
        ///code written by Joy
        /// </summary>
        System.Threading.Timer timer;
        /// <summary>
        //code written by Joy
        /// </summary>
        public List<Outlook.MailItem> pendingList;

        /// <summary>
        //code written by Joy
        /// </summary>
        public bool isTimerUploadRunning;
        /// <summary>
        //code written by Joy
        /// </summary>
        public bool isCopyRunninng;
        /// <summary>
        //code written by Joy
        /// </summary>
        public bool isMoveRunning;
        /// <summary>
        //code written by Joy
        /// </summary>
     
        public int no_of_t_item_uploaded=0;
        /// <summary>
        //code written by Joy
        /// </summary>
        public int no_of_pending_items_to_be_uploaded = 0;
        /// <summary>
        //code written by Joy
        /// </summary>
        public int no_of_moved_item_to_be_uploaded = 0;
        public int no_of_moved_item_uploaded = 0;
        /// <summary>
        //code written by Joy
        /// </summary>
        public int no_of_copied_item_to_be_uploaded = 0;
        public int no_of_copied_item_uploaded = 0;
        /// <summary>
        //code written by Joy
        /// </summary>
        public bool copy_button_clicked = false;
        public bool move_button_clicked = false;
        public Outlook.Selection copySelected;
        public Outlook.Selection moveSelected;
        /// <summary>
        /// fires to upadte the progressbar
        /// code written by Joy
        /// </summary>
        delegate void progressUpdater();
        /// <summary>
        /// <c>myFolders</c> Collection to hold all the outlook folder names in memory
        /// </summary>
        List<MAPIFolderWrapper> myFolders = new List<MAPIFolderWrapper>();

        /// <summary>
        /// <c>OutlookWindow</c> object of class <c>OutlookExplorerWrapper</c>
        /// <c>OutlookWindow</c> object holds the reference to OutLook explorer window which is opened
        /// </summary>
        OutlookExplorerWrapper OutlookWindow;

        /// <summary>
        /// <c>IsUploadingFormIsOpen</c> member field of <c>Boolean</c> type 
        /// IsUploadingFormIsOpen member field holds the status of the ouloading form ("frmUploadItemsList") 
        /// status is true when form is opened otherwise it will be false
        /// </summary>
        public static Boolean IsUploadingFormIsOpen = false;

        /// <summary>
        /// <c>test</c>window form("frmtest")  object 
        /// this form display the message "Please wair" in outlook window during processing the request to open the library
        /// </summary>
        frmtest test = null;

        //For context menu Rename and Edit properties

        /// <summary>
        /// <c>oFolderMenuButtonViewProperties</c>  CommandBarButton interface member
        /// this is not used in outlook addin 2010
        /// </summary>
        Office.CommandBarButton oFolderMenuButtonViewProperties;

        /// <summary>
        /// <c>oFolderMenuButtonEditConnectionProperties</c>  CommandBarButton interface member
        /// this is not used in outlook addin 2010
        /// </summary>
        Office.CommandBarButton oFolderMenuButtonEditConnectionProperties;

        /// <summary>
        /// <c>oFolderMenuButtonSharePointView</c>  CommandBarButton interface member
        ///it displays command bar button in outlook menu bar. this is not used in outlook addin 2010
        /// </summary>
        Office.CommandBarButton oFolderMenuButtonSharePointView;

        /// <summary>
        /// <c>oFolderMenuButtonSharePointView</c>  CommandBarButton interface member
        /// this is not used in outlook addin 2010
        /// </summary>
        Office.CommandBarButton oFolderMenuButtonOutlookView;

        /// <summary>
        /// <c>frmRenameObject</c> windows form object of <c>frmRename</c> 
        /// this object is required to rename <c>MAPI</c> outlook folder
        /// </summary>
        frmRename frmRenameObject;


        /// <summary>
        /// <c>oContextMenuFolder</c>  MAPIFolder interface member
        /// holds the MAPI Folder Currently selected
        /// </summary>
        Outlook.MAPIFolder oContextMenuFolder; //Get selected context menu folder

        /// <summary>
        /// <c>renameButton</c>  CommandBarButton  Interface member
        /// it displays rename button in outlook menu bar.
        /// It is not currently used in outlook addin 2010
        /// </summary>
        Office.CommandBarButton renameButton;

        /// <summary>
        /// <c>folderNewName</c> member field of type <c>String</c>
        /// it Holds the name of newly created Outlook MAPI Folder
        /// </summary>
        string folderNewName = string.Empty;

        /// <summary>
        /// <c>userOptions</c> object of class <c>XMLLogOptions</c>
        /// it Holds the Choice whether to  Delete Email  after uploading or Not.
        /// It also updates Choice made by user to Configuration file 
        /// </summary>
        XMLLogOptions userOptions;

        /// <summary>
        /// <c>myTargetFolder</c>  MAPIFolder Interface member
        /// it holds the MAPI folder where the user drags and drop items
        /// </summary>
        Microsoft.Office.Interop.Outlook.MAPIFolder myTargetFolder;

        /// <summary>
        /// <c>currentFolderSelected</c> member field of Type <c>String</c>
        /// It Holds the name of default selected folder
        /// </summary>
        string currentFolderSelected;

        /// <summary>
        /// <c>FromFolderGuid</c> member field of type <c>String</c>
        /// Holds the Unique ID of Source Folder from where items are dragged and droped to another folder 
        /// </summary>
        public static string FromFolderGuid = "folderguid";


        /// <summary>
        /// <c>IsMailItemUploaded</c> member field of type <c>bool</c>
        /// Holds the true/false value  whether the item is uploaded or not to sharepoint document library
        /// </summary>
        public static bool IsMailItemUploaded = false;

        /// <summary>
        /// <c>IsUrlIsTyped</c> member field of type <c>bool</c>
        /// Holds the true/false value whether the url of sharepoint document library is typed of dragged
        /// </summary>
        public static bool IsUrlIsTyped = false;


        /// <summary>
        /// <c>currentFolderSelectedGuid</c> member field of type String
        /// Holds the unique id of Currently  selected folder
        /// </summary>
        string currentFolderSelectedGuid;


        /// <summary>
        /// <c>newFolders</c> member field List Collection 
        /// The collection holds the MapiFolders
        /// </summary>
        public static List<MAPIFolderWrapper> newFolders = new List<MAPIFolderWrapper>();
        #endregion

        #region StartUp & ShutDown Events

        /// <summary>
        ///<c>ThisAddIn_Startup</c>  Outlook startup event
        /// This Event is  executed when outlook starts(outlook is opened)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            try
            {


                //declare and initialize variables used for the Instant PLUS check 
                Int32 result = 0;

                string filePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\\ITOPIA\\SharePoint Link 2010\\", "SharePoint Link.xml");

                //Make the call to Instant PLUS and store the result
                result = Ip2LibManaged.CallInstantPLUS(Ip2LibManaged.FLAGS_NONE, "30820274020100300D06092A864886F70D01010105000482025E3082025A02010002818100BA81817A32C249909671428137ECFC2AEF45E5F746C218550C2191F525A15E65DCD87CAD8B46EB870E55897D1D185B9D88BCDDB6D44CB8B9DDFD7DB4948E8CF91743377F31DB733438828CA0EC3176D8650C8F4B77578E60285F55049D9A707C61FF75C7C626492415BBCBE8058E4F220826A356F1C50B29C92B354C61BEF21F0201110281800AF88F254E47A9F97242E5CB5DA4874DD1D6EF68E60B6AD7D389810E6BA0149C94853482ADD6FECBB58C8F9DF2A71472ADB0C1BF75E665381C1DF855EA9EF93BBA5432731B506EDFE944C217EB09F27AA8203C7D7310D78995A9549690836B313CDCD0B2FA6AD79576977EB44B33D46577E9EA6939DCF8388761E25C3FADF9B1024100F3F70346EEA01874427576DFF5136A9B3A1AB4522ADD2073669376540C6CE65A1A613D0F4754FAD4C8DA70D3F5F5F0D4DAF97B0CF49D9BDA0ECE0F8F46A19BBF024100C3B4DA9372E3FDE1787C322A5B74F21800CDD6A4A85C1DC9D18D40B0F8736BDD3CF45CD5DDB8FD626CD1F11B1127439036A4974D257AF38EBCDD1D9CE08FC1A1024100AC35E43211DA6B9D5C16AE43BC0DB4A9CEA9703A00239E6F93B36295AE6AFCF44EDB3A28E70ECF2CCA039AEFF8E9D72CD6CE38BDD9D8AA3F91FADDCE8C35D75902405095C369E40386A8228D7E1170F3EB370F63D0DA63713971382B1AA3392077B57373ADC1796A4A3796385438525B762C52BC3E4CF150BEA42FA6577CD4EFE65102405162E2024CE04279E92731B490BE2431758FA6D032B0DB85B2AC782956832095800B4A8AB7D6FC1DB905CA38508FC3CC49994A48940CF9BB5761C07A9289D492", filePath);

                if (11574 != result)
                {
                    //this.Close();
                    //this.Application.ActiveExplorer().Close();
                    //this.Application.Quit();
                    return;
                }

                isAuthorized = true;

                #region Add-in Express Regions generated code - do not modify
                this.FormsManager = AddinExpress.OL.ADXOlFormsManager.CurrentInstance;
                this.FormsManager.OnInitialize +=
                    new AddinExpress.OL.ADXOlFormsManager.OnComponentInitialize_EventHandler(this.FormsManager_OnInitialize);
                //this.FormsManager.ADXBeforeFolderSwitchEx +=
                //     new AddinExpress.OL.ADXOlFormsManager.BeforeFolderSwitchEx_EventHandler(FormsManager_ADXBeforeFolderSwitchEx);

                this.FormsManager.Initialize(this);
                #endregion

                //DateTime dtExpiredDate = new DateTime(2010, 08, 05);
                //DateTime dtWorkingDate = DateTime.Now;
                //TimeSpan t = new TimeSpan();
                //t = dtExpiredDate.Subtract(dtWorkingDate);

                //if (t.Days < 30 && t.Days >= 0)
                //{

                //outlookObj = new Outlook.Application();

                OutlookObj = Globals.ThisAddIn.Application;
                //Gte MAPI Name space
                outlookNameSpace = OutlookObj.GetNamespace("MAPI");


                //Get inbox folder
                Outlook.MAPIFolder oInBox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                //Get current user details to save the xml file based on user
                string userName = outlookNameSpace.DefaultStore.DisplayName;
                userName = userName.Replace("-", "_");
                userName = userName.Replace(" ", "");
                UserLogManagerUtility.UserXMLFileName = userName;

                //Get parent root folder
                Outlook.MAPIFolder olMailRootFolder = (Outlook.MAPIFolder)oInBox.Parent;
                //Get all folder
                oMailRootFolders = olMailRootFolder.Folders;
                //Create folder remove event
                oMailRootFolders.FolderRemove += new Microsoft.Office.Interop.Outlook.FoldersEvents_FolderRemoveEventHandler(oMailRootFolders_FolderRemove);

                //Set inbox folder as default


                try
                {



                    OutlookObj.ActiveExplorer().CurrentFolder = oInBox;
                    addinExplorer = this.Application.ActiveExplorer();
                    addinExplorer.BeforeItemPaste += new Microsoft.Office.Interop.Outlook.ExplorerEvents_10_BeforeItemPasteEventHandler(addinExplorer_BeforeItemPaste);
                    //Create folder Switch event
                    addinExplorer.FolderSwitch += new Microsoft.Office.Interop.Outlook.ExplorerEvents_10_FolderSwitchEventHandler(addinExplorer_FolderSwitch);


                    //crete folder context menu disply                 
                    this.Application.FolderContextMenuDisplay += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_FolderContextMenuDisplayEventHandler(Application_FolderContextMenuDisplay);

                    this.Application.ContextMenuClose += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ContextMenuCloseEventHandler(Application_ContextMenuClose);

                    strParentMenuTag = strParentMenuName;
                    // Removing the Existing Menu Bar
                    //  RemoveItopiaMenuBarIfExists(strParentMenuTag);
                    // Adding The Menu Bar Freshly
                    //  CreateParentMenu(strParentMenuTag, strParentMenuName);
                    myTargetFolder = oInBox;
                    CreateAddEventOnFolders();
                    //CreateDefaultAddEventOnFolders();
                    //Create outlook explorer wrapper class                   
                    OutlookWindow = new OutlookExplorerWrapper(OutlookObj.ActiveExplorer());
                    OutlookWindow.Close += new EventHandler(OutlookWindow_Close);

                    ((Outlook.ExplorerEvents_Event)addinExplorer).BeforeFolderSwitch += new Microsoft.Office.Interop.Outlook.ExplorerEvents_BeforeFolderSwitchEventHandler(ThisAddIn_BeforeFolderSwitch);

                    oMailRootFolders.FolderChange += new Outlook.FoldersEvents_FolderChangeEventHandler(oMailRootFolders_FolderChange);

                    foreach (Outlook.MAPIFolder item in oMailRootFolders)
                    {

                        try
                        {
                            item.Folders.FolderChange -= new Outlook.FoldersEvents_FolderChangeEventHandler(oMailRootFolders_FolderChange);

                        }
                        catch (Exception)
                        {
                        }
                        item.Folders.FolderChange += new Outlook.FoldersEvents_FolderChangeEventHandler(oMailRootFolders_FolderChange);

                    }


                    Outlook.Items activeDroppingFolderItems;

                    activeDroppingFolderItems = oInBox.Items;

                    userOptions = UserLogManagerUtility.GetUserConfigurationOptions();
                    currentFolderSelected = oInBox.Name;
                    currentFolderSelectedGuid = oInBox.EntryID;
                }
                catch (Exception ex)
                {
                    ListWebClass.Log(ex.Message, true);
                }

                // }



            }
            catch (Exception ex)
            {
                EncodingAndDecoding.ShowMessageBox("StartUP", ex.Message, MessageBoxIcon.Error);
            }
          
           

           
            /// <summary>
            //code written by Joy
            ///initializes the object of timer 
            /// </summary>
            
            System.Threading.AutoResetEvent reset = new System.Threading.AutoResetEvent(true);
            
            timer = new System.Threading.Timer(new System.Threading.TimerCallback(doBackgroundUploading), reset, 180000, 180000);
            
            GC.KeepAlive(timer);
            form = new Form();
            form.Opacity = 0.01;
            form.Show();
            form.Visible = false;     
             
          

        }
      /// <summary>
      /// this method executes by the timer's TimerCallback method
      /// finds the mapped folders and pending items,uploads each folder's pending items
      /// code wrritten by Joy
      /// </summary>
      /// <param name="state"></param>
        void doBackgroundUploading(object state)
        {
            try
            {
                // MessageBox.Show("Hello:I have been fired");
                var customCat = "Pending Uploads";
                DataSet ds = new DataSet();
                if (File.Exists(UserLogManagerUtility.XMLFilePath))
                {
                    ds.Tables.Clear();
                    ds.ReadXml(UserLogManagerUtility.XMLFilePath);

                }
                string mappedFolderName;
                Outlook.MAPIFolder mappedFolder;
                Outlook.MAPIFolder olInBox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Microsoft.Office.Interop.Outlook.MAPIFolder parentFolder = (Microsoft.Office.Interop.Outlook.MAPIFolder)olInBox.Parent;
                try
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {


                        mappedFolderName = ds.Tables[0].Rows[i]["DisplayName"].ToString();
                        mappedFolder = MAPIFolderWrapper.GetFolder(parentFolder, mappedFolderName);

                        UploadBrokenUploads BrokenUplaodObject = new UploadBrokenUploads(mappedFolder, this.Application.ActiveExplorer(), true);
                        if (isuploadRunning == false&&isTimerUploadRunning==false&&isMoveRunning==false&&isCopyRunninng==false)
                        {
                            pendingList = new List<Outlook.MailItem>();
                            foreach (Outlook.MailItem brokenItem in mappedFolder.Items)
                            {
                                try
                                {
                                    if (brokenItem.Categories.Contains(customCat))
                                    {


                                        // Use a timer to simulate an event in which the FakeMessageBox should be closed 
                                        //MessageBox.Show(mappedFolder.Items.Count.ToString());
                                        //MessageBox.Show("Please wait while we are uploading" + " " + brokenItem.Subject);
                                        pendingList.Add(brokenItem);
                                        addinExplorer_beforeMovingToMappedFolder(brokenItem, mappedFolder, false);

                                    }
                                }
                                catch (Exception ex)
                                {
                                }
                            }
                            if (pendingList != null)
                            {
                                no_of_t_item_uploaded = 0;
                                no_of_pending_items_to_be_uploaded = pendingList.Count;
                                BrokenUplaodObject.uploadBrokenUploadsIfExists();
                            }
                        }

                    }

                }
                catch (Exception ex)
                {
                }
            }
            catch (Exception ex)
            {

            }
 
            
        }
       
       

        /// <summary>
        /// <c>oMailRootFolders_FolderChange</c> Outlook Event
        /// this event is executed when any outlook folder is moved  to another  location or deleted
        /// </summary>
        /// <param name="folder"></param>
        void oMailRootFolders_FolderChange(Microsoft.Office.Interop.Outlook.MAPIFolder folder)
        {
            try
            {
                string f = folder.EntryID.ToString();
                string oldvalue = currentFolderSelected;
                string currentfolder = currentFolderSelectedGuid;
                AddfFolderinSessionMapi();
                MAPIFolderWrapper folderWrapper = myFolders.Find(delegate(MAPIFolderWrapper p) { return p.FolderName == oldvalue; });
                if (folderWrapper != null && !string.IsNullOrEmpty(folder.WebViewURL))
                {
                    if (currentFolderSelectedGuid.Contains(folder.EntryID))
                    {
                        UserLogManagerUtility.UpdateFolderConfigNodeDetails(oldvalue, "DisplayName", folder.Name);
                        MAPIFolderWrapper omapi = null;
                        //Doc name is empty means Folder is not mapped with Doc Lib
                        //omapi = new MAPIFolderWrapper(oChildFolder, addinExplorer, false);
                        if (oldvalue != folder.Name)
                        {
                            omapi = new MAPIFolderWrapper(ref  folder, this.Application.ActiveExplorer(), UserLogManagerUtility.IsDocumentLibrary(folder.Name));
                            try
                            {
                                foreach (MAPIFolderWrapper item in myFolders)
                                {
                                    if (item.FolderName == oldvalue)
                                    {
                                        Boolean result = myFolders.Remove(item);
                                    }
                                }
                            }
                            catch (Exception)
                            {
                            }

                            omapi.AttachedFolder.WebViewURL = UserLogManagerUtility.GetSPSiteURL(folder.Name);
                            myFolders.Add(omapi);
                        }
                    }

                }

            }
            catch (Exception ex)
            {
            }
            updatefolderlocationin();
        }


        /// <summary>
        /// <c>updatefolderlocationin</c> member function
        /// this member function calls <c>UserLogManagerUtility.UpdateFolderConfigNodeDetails</c> function to updates the configuration file when any folder is moved to another location or Deleted
        /// </summary>
        public void updatefolderlocationin()
        {
            try
            {
                string foldername = "";
                string folderpath = "";
                foreach (MAPIFolderWrapper item in myFolders)
                {
                    Outlook.MAPIFolder fol = item.AttachedFolder;
                    foldername = fol.Name; // item.AttachedFolder.Name;
                    folderpath = fol.FolderPath;

                    if (folderpath.IndexOf("\\Deleted Items\\") == -1)
                    {
                        UserLogManagerUtility.UpdateFolderConfigNodeDetails(foldername, "OutlookLocation", folderpath);
                    }



                }
            }
            catch (Exception)
            {


            }
        }

        /// <summary>
        /// <c>ThisAddIn_Shutdown</c> 
        /// Outlook shutdown event
        /// This event  is executed when outlook is closed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            #region Add-in Express Regions generated code - do not modify

            #endregion
            try
            {
                this.FormsManager.Finalize(this);

                SetOriginalUrls();
                if (ThisAddIn.frmUploadEmailMessagesObject != null)
                {
                    frmUploadEmailMessagesObject.ParentForm.Close();
                    frmUploadEmailMessagesObject.Dispose();
                }
                string path = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + "MaPiFolderTemp.htm";
                if (System.IO.File.Exists(path))
                {
                    System.IO.File.Delete(path);
                }
            }
            catch { }


        }

        #endregion

        #region Other Events

        /// <summary>
        /// <c>addinExplorer_BeforeItemPaste</c> Event Handler
        /// Fires when any item is dragged and dropped to an outlook Folder
        /// </summary>
        /// <param name="ClipboardContent"></param>
        /// <param name="Target"></param>
        /// <param name="Cancel"></param>
        void addinExplorer_BeforeItemPaste(ref object ClipboardContent, Microsoft.Office.Interop.Outlook.MAPIFolder Target, ref bool Cancel)
        {
          ///code written by Joy///
          ///checks if any upload is running
            if (isuploadRunning == true||isTimerUploadRunning==true||isMoveRunning==true||isCopyRunninng==true)
            {
                frmMessageWindow objMessage = new frmMessageWindow();
                objMessage.DisplayMessage = "Your uploads are still running.Please wait for sometime.";
                objMessage.TopLevel = true;
                objMessage.TopMost = true;
                objMessage.ShowDialog();
                objMessage.Dispose();
                //EncodingAndDecoding.ShowMessageBox("", "Please check the uploading form. Form is still open.", MessageBoxIcon.Warning);
                Cancel = true;
                return;


            }
            else
            {
                try
                {
                    ///all code in this section written by joy
                    ///retrieves the mail items from the Selection and set the progressbar value to default

                    Outlook.MailItem mailitem;
                    no_of_items_copied = 0;

                    Outlook.Application myApplication = Globals.ThisAddIn.Application;
                    Outlook.Explorer myActiveExplorer = (Outlook.Explorer)myApplication.ActiveExplorer();
                    
                    ///retrieves the mail items from the Selection 
                    oselection = myActiveExplorer.Selection;


                    no_of_items_to_be_uploaded = oselection.Count;


                    AddfFolderinSessionMapi();

                    IsUrlIsTyped = false;
                    currentFolderSelected = Target.Name;
                    currentFolderSelectedGuid = Target.EntryID;

                    myTargetFolder = Target;

                    if (IsUploadingFormIsOpen == true)
                    {
                        if (Globals.ThisAddIn.frmlistObject != null)
                        {
                            ///set the progressbar value to default
                            frmlistObject.progressBar1.Value = frmlistObject.progressBar1.Minimum;
                            frmlistObject.lblPRStatus.Text = "";
                        }
                        //frmUploadItemsList frmUplList = new frmUploadItemsList();
                        //frmUplList.progressBar1.Value = frmUplList.progressBar1.Minimum;
                        //frmUplList.lblPRStatus.Text = "";
                        //frmUplList.Refresh();
                        //frmUplList.Visible = false;
                        //frmUplList.Dispose();



                    }

                    //Check  dropping item is from browser or not

                    if (ClipboardContent.GetType().Name == "String")
                    {
                        //Set active folder as TargetFolder
                        this.Application.ActiveExplorer().SelectFolder(Target);

                        //'Cerate instance
                        frmSPSiteConfigurationObject = new frmSPSiteConfiguration();
                        //Get the drop url


                        frmSPSiteConfigurationObject.URL = Convert.ToString(ClipboardContent);

                        frmSPSiteConfigurationObject.ShowDialog();
                        if (frmSPSiteConfigurationObject.IsConfigureCompleted)
                        {
                            //Save the details in log proeprties object
                            XMLLogProperties xLogProperties = frmSPSiteConfigurationObject.FolderConfigProperties;


                            Outlook.MAPIFolder newFolder = null;
                            bool result = CreateFolderInOutLookSideMenu(xLogProperties.DisplayFolderName, xLogProperties.SiteURL, out newFolder, Target);

                            Cancel = true;
                            if (result == true && newFolder != null)
                            {


                                //Set new folder location
                                xLogProperties.OutlookFolderLocation = newFolder.FolderPath;
                                //Create node in xml file
                                UserLogManagerUtility.CreateXMLFileForStoringUserCredentials(xLogProperties);

                                MAPIFolderWrapper omapi = null;
                                if (string.IsNullOrEmpty(xLogProperties.DocumentLibraryName) == true)
                                {
                                    //Doc name is empty means Folder is not mapped with Doc Lib
                                    omapi = new MAPIFolderWrapper(ref  newFolder, addinExplorer, false);
                                }
                                else
                                {
                                    omapi = new MAPIFolderWrapper(ref newFolder, addinExplorer, true);
                                }
                                omapi.AttachedFolder.WebViewURL = ListWebClass.WebViewUrl(omapi.AttachedFolder.WebViewURL);
                                myFolders.Add(omapi);

                            }

                        }
                        else
                        {
                            frmSPSiteConfigurationObject.Close();
                            Cancel = true;
                        }
                    }
                    else
                    {

                    }
                }
                catch (Exception ex)
                {
                    EncodingAndDecoding.ShowMessageBox("BeforeItemPaste", ex.Message, MessageBoxIcon.Error);
                }
                finally
                {

                }
            }
        }


  
        
        
        /// <summary>
        /// code in this section written by Joy
        /// this method fires by the timer before uploading the mail items
        /// </summary>
        /// <param name="ClipboardContent"></param>
        /// <param name="Target"></param>
        /// <param name="Cancel"></param>
        void addinExplorer_beforeMovingToMappedFolder(object ClipboardContent, Microsoft.Office.Interop.Outlook.MAPIFolder Target, bool Cancel)
        {
            if (isTimerUploadRunning == false && isuploadRunning == false)
            {
                try
                {


                    AddfFolderinSessionMapi();

                    IsUrlIsTyped = false;
                    currentFolderSelected = Target.Name;
                    currentFolderSelectedGuid = Target.EntryID;

                    myTargetFolder = Target;
                    if (IsUploadingFormIsOpen == true)
                    {
                        if (Globals.ThisAddIn.frmlistObject != null)
                        {
                            Globals.ThisAddIn.frmlistObject.Invoke(new progressUpdater(()=>
                                 {
                            frmlistObject.progressBar1.Value = frmlistObject.progressBar1.Minimum;
                            frmlistObject.lblPRStatus.Text = "";
                                 }));
                        }

                        //frmMessageWindow objMessage = new frmMessageWindow();
                        //objMessage.DisplayMessage = "Please check the uploading form. Form is still open.";
                        //objMessage.TopLevel = true;
                        //objMessage.TopMost = true;
                        //objMessage.ShowDialog();
                        //objMessage.Dispose();
                        ////EncodingAndDecoding.ShowMessageBox("", "Please check the uploading form. Form is still open.", MessageBoxIcon.Warning);
                        //Cancel = true;
                        //return;


                    }

                    //Check  dropping item is from browser or not

                    if (ClipboardContent.GetType().Name == "String")
                    {
                        //Set active folder as TargetFolder
                        this.Application.ActiveExplorer().SelectFolder(Target);

                        //'Cerate instance
                        frmSPSiteConfigurationObject = new frmSPSiteConfiguration();
                        //Get the drop url

                        frmSPSiteConfigurationObject.URL = Convert.ToString(ClipboardContent);

                        frmSPSiteConfigurationObject.ShowDialog();
                        if (frmSPSiteConfigurationObject.IsConfigureCompleted)
                        {
                            //Save the details in log proeprties object
                            XMLLogProperties xLogProperties = frmSPSiteConfigurationObject.FolderConfigProperties;


                            Outlook.MAPIFolder newFolder = null;
                            bool result = CreateFolderInOutLookSideMenu(xLogProperties.DisplayFolderName, xLogProperties.SiteURL, out newFolder, Target);

                            Cancel = true;
                            if (result == true && newFolder != null)
                            {


                                //Set new folder location
                                xLogProperties.OutlookFolderLocation = newFolder.FolderPath;
                                //Create node in xml file
                                UserLogManagerUtility.CreateXMLFileForStoringUserCredentials(xLogProperties);

                                MAPIFolderWrapper omapi = null;
                                if (string.IsNullOrEmpty(xLogProperties.DocumentLibraryName) == true)
                                {
                                    //Doc name is empty means Folder is not mapped with Doc Lib
                                    omapi = new MAPIFolderWrapper(ref  newFolder, addinExplorer, false);
                                }
                                else
                                {
                                    omapi = new MAPIFolderWrapper(ref newFolder, addinExplorer, true);
                                }
                                omapi.AttachedFolder.WebViewURL = ListWebClass.WebViewUrl(omapi.AttachedFolder.WebViewURL);
                                myFolders.Add(omapi);

                            }

                        }
                        else
                        {
                            frmSPSiteConfigurationObject.Close();
                            Cancel = true;
                        }
                    }
                    else
                    {

                    }
                }
                catch (Exception ex)
                {
                    EncodingAndDecoding.ShowMessageBox("BeforeItemPaste", ex.Message, MessageBoxIcon.Error);
                }
                finally
                {

                }
            }
        }
        //////////////////////////////////////////////////////////////Updated by joy on 30.7.2012///////////////////////////////////
        
        
        
        /// <summary>
        /// <c>ThisAddIn_BeforeFolderSwitch</c> Event Handler
        /// Before folder swith event while selection of events
        /// This event is executed when user switches to another folder 
        /// </summary>
        /// <param name="NewFolder"></param>
        /// <param name="Cancel"></param>
        void ThisAddIn_BeforeFolderSwitch(object NewFolder, ref bool Cancel)
        {

            ADXOlForm1Item.Enabled = false;

            //  ListWebClass.Log("Folder is switched:", true);
            try
            {

                //  addinExplorer.Activate();
                Outlook.MAPIFolder ofolder = (Outlook.MAPIFolder)NewFolder;
                currentFolderSelected = ofolder.Name;
                currentFolderSelectedGuid = ofolder.EntryID;
                AddfFolderinSessionMapi();
                //Check selected folder is out Itopia Folder or not


                MAPIFolderWrapper myLocatedObject = myFolders.Find(delegate(MAPIFolderWrapper p) { return p.FolderName == ofolder.Name; });



                if (myLocatedObject != null)
                {
                    myLocatedObject.IsFolderAuthenticated = false;
                    if (ofolder.WebViewOn == true && myLocatedObject.IsFolderAuthenticated == false)
                    {


                        // ofolder.WebViewURL = UserLogManagerUtility.GetSPSiteURL(ofolder.Name);
                        string folderurl = UserLogManagerUtility.GetSPSiteURL(ofolder.Name);
                        string returnedvalue = ListWebClass.WebViewUrl(folderurl);



                        XmlNode folderNode = UserLogManagerUtility.GetSPSiteURLDetails("", ofolder.Name);
                        string m_strUser = EncodingAndDecoding.Base64Decode(folderNode.ChildNodes[0].InnerText);
                        string m_strPwd = EncodingAndDecoding.Base64Decode(folderNode.ChildNodes[1].InnerText);
                        // ofolder.WebViewURL = ListWebClass.WebViewUrl(ofolder.WebViewURL);
                        try
                        {
                            if (test != null)
                            {
                                test.Dispose();
                            }

                            //  test = new frmtest(m_strUser, m_strPwd, ofolder.WebViewURL);
                            test = new frmtest(m_strUser, m_strPwd, folderurl);
                            test.ShowDialog();


                            ThisAddIn.CurrentWebUrlLink = folderurl;
                            ADXOlForm1Item.Enabled = true;

                        }
                        catch (Exception)
                        { }

                        //MessageBox.Show(test.AuthenticationCompleted.ToString());
                        myLocatedObject.IsFolderAuthenticated = test.AuthenticationCompleted;
                        ///////////

                        try
                        {
                            Outlook.MAPIFolder myfolder = (Outlook.MAPIFolder)ofolder.Parent;

                            try
                            {
                                myfolder.Folders.FolderChange -= new Outlook.FoldersEvents_FolderChangeEventHandler(oMailRootFolders_FolderChange);

                            }
                            catch (Exception)
                            { }
                            myfolder.Folders.FolderChange += new Outlook.FoldersEvents_FolderChangeEventHandler(oMailRootFolders_FolderChange);

                        }
                        catch (Exception)
                        {
                        }

                        ////////////
                        // ListWebClass.Log("Normal switched:", true);
                        //  ofolder.WebViewURL = ListWebClass.WebViewUrl(ofolder.WebViewURL);


                    }
                    else
                    {
                        if (ofolder.WebViewOn == true)
                        {
                            // ofolder.WebViewURL = UserLogManagerUtility.GetSPSiteURL(ofolder.Name);


                            XmlNode folderNode = UserLogManagerUtility.GetSPSiteURLDetails("", ofolder.Name);
                            string m_strUser = EncodingAndDecoding.Base64Decode(folderNode.ChildNodes[0].InnerText);
                            string m_strPwd = EncodingAndDecoding.Base64Decode(folderNode.ChildNodes[1].InnerText);
                            try
                            {
                                if (test != null)
                                {
                                    test.Dispose();
                                }
                                test = new frmtest(m_strUser, m_strPwd, ofolder.WebViewURL);
                                test.ShowDialog();


                                ThisAddIn.CurrentWebUrlLink = ofolder.WebViewURL;
                                ADXOlForm1Item.Enabled = true;
                            }
                            catch (Exception)
                            {


                            }
                            // ListWebClass.Log("Alternate switch:", true);

                            //  ofolder.WebViewURL = ListWebClass.WebViewUrl(ofolder.WebViewURL);
                        }


                    }
                }




            }
            catch (Exception ex)
            {
                ListWebClass.Log("javascript error caused:" + ex.Message, true);

            }

            updatefolderlocationin();



        }



        /// <summary>
        /// <c>OutlookWindow_Close</c>  Event Handler
        /// Outlook close window event.
        /// This event is fired when an outlook window is closed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void OutlookWindow_Close(object sender, EventArgs e)
        {
            /////////////////////////////////Modified on 16.08.2012///////////////////////////////////////////
           ///code written by Joy
           ///disposes the timer when outlook is being closed 
            timer.Dispose();
            /////////////////////////////////Modified on 16.08.2012///////////////////////////////////////////
            SetOriginalUrls();
            try
            {
                OutlookWindow = (OutlookExplorerWrapper)sender;
                OutlookWindow.Close -= new EventHandler(OutlookWindow_Close);
            }
            catch { }



        }


        /// <summary>
        /// <c>SetOriginalUrls</c> member function 
        /// This member function gets  the sharepoint document  library url  from configuration file and apply to Outlook folder.
        /// </summary>
        private void SetOriginalUrls()
        {
            try
            {
                foreach (MAPIFolderWrapper item in myFolders)
                {
                    try
                    {
                        string returl = UserLogManagerUtility.GetSPSiteURL(item.AttachedFolder.Name);
                        item.AttachedFolder.WebViewURL = returl;
                    }
                    catch (Exception)
                    { }
                }

                try
                {
                    string path = UserLogManagerUtility.RootDirectory;
                    string[] files = Directory.GetFiles(UserLogManagerUtility.RootDirectory, "*.msg");
                    foreach (string p in files)
                    {
                        File.Delete(p);
                    }


                }
                catch (Exception)
                {
                }


            }
            catch (Exception ex)
            { }

        }

        /// <summary>
        /// <c>addinExplorer_FolderSwitch</c> Event Handler
        /// Fires when folder  selection is changed . it retrieves the selected folder and its unique id
        /// </summary>
        void addinExplorer_FolderSwitch()
        {
            try
            {

                oCurrentSelectedFolder = addinExplorer.CurrentFolder;

                FromFolderGuid = addinExplorer.CurrentFolder.EntryID;


                //Check selected folder is out Itopia Folder or not

            }
            catch (Exception ex)
            { }
        }

        /// <summary>
        /// <c>oMailRootFolders_FolderRemove</c> Event Handler
        /// Fires when  a folder is removed 
        /// </summary>
        void oMailRootFolders_FolderRemove()
        {
            if (oCurrentSelectedFolder != null)
            {
                //Update the status
                //   UserLogManagerUtility.UpdateUserFolderRemovedStatus(oCurrentSelectedFolder.Name);
                //  UserLogManagerUtility.UpdateFolderConfigNodeDetails(oCurrentSelectedFolder.Name, "OutlookLocation",oCurrentSelectedFolder.FolderPath);
            }
        }


        #endregion

        #region ToolMenu Events

        /// <summary>
        /// <c>btnConnectionProperties_Click</c> Event handler
        /// it opens Configuration properties Window form to provide Connection setting for new Connection
        /// Currently it is not used in addin for outlook 2010(ribbon is being used )
        /// </summary>
        /// <param name="Ctrl"></param>
        /// <param name="CancelDefault"></param>
        void btnConnectionProperties_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                frmConnectionProperties objfrmConnectionProperties = new frmConnectionProperties();
                objfrmConnectionProperties.ShowDialog();
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// <c>btnOptions_Click</c> Event Handler
        /// this event display the options window form("frmOptions") to choose the option to Automatically delete the mailitem after uploading or not.
        /// and then apply changes to the configuration file.
        /// Currently it is not used in addin for outlook 2010(ribbon is being used ) 
        /// </summary>
        /// <param name="Ctrl"></param>
        /// <param name="CancelDefault"></param>
        void btnOptions_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
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
        /// <c>btnNewConnection_Click</c> Event Handler
        /// It opens the New Connection window form("frmSPSiteConfiguration") to create new connection. 
        /// this is used in outlook addin for outlook 2007 not for outlook 2010. in outlook 2010 ribbon is used
        /// </summary>
        /// <param name="Ctrl"></param>
        /// <param name="CancelDefault"></param>
        void btnNewConnection_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                frmSPSiteConfigurationObject = new frmSPSiteConfiguration();
                frmSPSiteConfigurationObject.ShowDialog();
            }
            catch (Exception ex)
            {

            }
        }


        # endregion

        #region VSTO generated code

        /// <summary>
        /// <c>InternalStartup</c> member function
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);




        }

        #endregion

        # region Methods
        /// <summary>
        /// <c>RemoveItopiaMenuBarIfExists</c> member function
        /// If the menu already exists, remove it.
        /// </summary>
        /// <param name="strParentMenuTag">Menu name as string</param>
        private void RemoveItopiaMenuBarIfExists(string strParentMenuTag)
        {
            try
            {
                menuBarItopia = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;

                Office.CommandBarPopup foundMenu = (Office.CommandBarPopup)menuBarItopia.FindControl(
                    Office.MsoControlType.msoControlPopup, System.Type.Missing, strParentMenuName, true, true);


                if (foundMenu != null)
                {
                    foundMenu.Delete(true);
                }
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// <c>CreateParentMenu</c> member function
        /// Method to create Parent Menu ie At the Top 
        /// </summary>
        /// <param name="strParentMenuTag"></param>
        /// <param name="strCaption"></param>
        /// <returns></returns>
        private bool CreateParentMenu(string strParentMenuTag, string strCaption)
        {
            try
            {
                //Define the existent Menu Bar


                menuBarItopia = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;

                //Add new ITopia menu bar to active menubar
                menuBarSharePoint = (Office.CommandBarButton)menuBarItopia.Controls.Add(
                Office.MsoControlType.msoControlButton, missing, missing, missing, false);

                //If I dont find the newMenuBar, I add it
                if (menuBarSharePoint != null)
                {
                    //Add caption and tag
                    menuBarSharePoint.Caption = strCaption;
                    menuBarSharePoint.Tag = strParentMenuTag;


                    //Create SharePoint menu under ItopiaToola Menu
                    //menuBarConnectionProperties = (Office.CommandBarPopup)menuBarSharePoint.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, 1, true);
                    //menuBarConnectionProperties.Caption = "Sharepoint";
                    //menuBarConnectionProperties.Tag = "SharePoint";

                    //Add button control to SharePoint popup menu
                    btnOptions = (Office.CommandBarButton)menuBarItopia.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                    btnOptions.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    btnOptions.Caption = "Options";
                    btnOptions.FaceId = 630;
                    btnOptions.Tag = "Options";
                    btnOptions.Visible = true;
                    btnOptions.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnOptions_Click);

                    //Add button control to SharePoint popup menu
                    btnNewConnection = (Office.CommandBarButton)menuBarItopia.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                    btnNewConnection.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    btnNewConnection.Caption = "New Connection";
                    btnNewConnection.FaceId = 611;

                    btnNewConnection.Tag = "New Connection";
                    btnNewConnection.Visible = true;
                    btnNewConnection.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnNewConnection_Click);

                    //Add button control to SharePoint popup menu
                    btnConnectionProperties = (Office.CommandBarButton)menuBarItopia.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                    btnConnectionProperties.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    btnConnectionProperties.Caption = "Connection Properties";
                    btnConnectionProperties.FaceId = 610;
                    btnConnectionProperties.Tag = "Connection Properties";
                    btnConnectionProperties.Visible = true;
                    btnConnectionProperties.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnConnectionProperties_Click);

                    return true;

                }
            }
            catch (Exception ex)
            {

                throw ex;
            }

            return false;

        }
       
        /// <summary>
        /// <c>CreateFolderInOutLookSideMenu</c> member Function
        /// Method to create folder  under destination folder.
        /// This member function creates folder in outlook side bar
        /// </summary>
        /// <param name="strDisplayName"></param>
        /// <param name="strGetURL"></param>
        /// <returns></returns>
        /// 


        public bool CreateFolderInOutLookSideMenu(string strDisplayName, string strGetURL, out Outlook.MAPIFolder newFolder, Microsoft.Office.Interop.Outlook.MAPIFolder Target)
        {
            newFolder = null;
            try
            {


                Outlook.MAPIFolder olInboxFolder = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                Outlook.MAPIFolder olMailRootFolder = (Outlook.MAPIFolder)Target; //olInboxFolder.Parent
                newFolder = olMailRootFolder.Folders.Add(strDisplayName, missing);
                try
                {
                    Target.Folders.FolderChange -= new Outlook.FoldersEvents_FolderChangeEventHandler(oMailRootFolders_FolderChange);
                }
                catch (Exception)
                {


                }
                Target.Folders.FolderChange += new Outlook.FoldersEvents_FolderChangeEventHandler(oMailRootFolders_FolderChange);
                try
                {
                    newFolder.Folders.FolderChange -= new Outlook.FoldersEvents_FolderChangeEventHandler(oMailRootFolders_FolderChange);
                }
                catch (Exception)
                {


                }

                newFolder.Folders.FolderChange += new Outlook.FoldersEvents_FolderChangeEventHandler(oMailRootFolders_FolderChange);

                if (newFolder.Name == strDisplayName)
                {
                    newFolder.WebViewAllowNavigation = true;




                    newFolder.WebViewURL = strGetURL;


                    newFolder.WebViewOn = true;


                    Outlook.Application oapp = new Outlook.Application();
                    Outlook.MailItem mitem = (Outlook.MailItem)oapp.CreateItem(Outlook.OlItemType.olMailItem);
                    mitem.Subject = "Sharepoint Site URL";
                    mitem.HTMLBody = "<a href=\"" + strGetURL + "\" >" + strGetURL + "</a>";
                    mitem.UnRead = false;
                    mitem.Move(newFolder);

                }




                return true;
            }
            catch (Exception ex)
            {
                throw ex;

            }
            return false;
        }


        /// <summary>
        /// <c>CreateAddEventOnFolders</c>
        /// Method to create item add event on all the added folders.
        /// this member function add all sharepoint mapped folders in List Collection
        /// </summary>
        private void CreateAddEventOnFolders()
        {


            Outlook.MAPIFolder oInBox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.MAPIFolder parentFolder = (Outlook.MAPIFolder)myTargetFolder; //oInBox.Parent;

            //Get Default folder information
            UserLogManagerUtility.AddDefaultMissingConnections();

            //Get all the folder information
            XmlNodeList xFolders = UserLogManagerUtility.GetAllFoldersDetails(UserStatus.Active);

            MAPIFolderWrapper omapi = null;
            if (xFolders != null)
            {
                string folderName = string.Empty, DocLibName = string.Empty;
                if (addinExplorer == null)
                {
                    MessageBox.Show("In Main Forms");
                }

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
                            oChildFolder.WebViewOn = true;
                            //  MAPIFolderWrapper omapi = null;
                            if (string.IsNullOrEmpty(DocLibName) == true)
                            {
                                //Doc name is empty means Folder is not mapped with Doc Lib
                                //omapi = new MAPIFolderWrapper(oChildFolder, addinExplorer, false);
                                omapi = new MAPIFolderWrapper(ref  oChildFolder, this.Application.ActiveExplorer(), false);

                            }
                            else
                            {
                                omapi = new MAPIFolderWrapper(ref  oChildFolder, this.Application.ActiveExplorer(), true);
                                //omapi = new MAPIFolderWrapper(oChildFolder, addinExplorer, true);

                            }

                            //
                            string returl = UserLogManagerUtility.GetSPSiteURL(omapi.AttachedFolder.Name);
                            //string relativepath = UserLogManagerUtility.GetRelativePath(omapi.AttachedFolder.Name);
                            //string rootpath = "";
                            //if (returl.LastIndexOf(relativepath)!=-1)
                            //{
                            //    rootpath = returl.Remove(returl.LastIndexOf(relativepath));
                            //}
                            // string virtualpath = ListWebClass.WebViewUrl(returl);

                            //  omapi.AttachedFolder.WebViewURL = rootpath + "/_layouts/OutlookIntegration/DisplayImage.aspx?Action=OLIssue&ReturnUrl=" + returl;
                            omapi.AttachedFolder.WebViewURL = ListWebClass.WebViewUrl(returl);

                            myFolders.Add(omapi);


                        }
                        else
                        {
                            //create mapi folder

                            XMLLogProperties xLogProperties = new XMLLogProperties();
                            xLogProperties.UserName = EncodingAndDecoding.Base64Decode(xNode.ChildNodes[0].InnerText);
                            xLogProperties.Password = EncodingAndDecoding.Base64Decode(xNode.ChildNodes[1].InnerText);
                            xLogProperties.DisplayFolderName = folderName;
                            xLogProperties.SiteURL = xNode.ChildNodes[4].InnerText;

                            xLogProperties.UsersStatus = UserStatus.Active;
                            xLogProperties.DocumentLibraryName = xNode.ChildNodes[3].InnerText;
                            xLogProperties.DroppedURLType = "";

                            if (xNode.ChildNodes[5].InnerText == "Manually Specified")
                            {
                                xLogProperties.FolderAuthenticationType = AuthenticationType.Manual;
                            }
                            else
                            {
                                xLogProperties.FolderAuthenticationType = AuthenticationType.Domain;

                            }

                            xLogProperties.SPSiteVersion = SPVersionClass.GetSPVersionFromUrl(xLogProperties.SiteURL, xLogProperties.UserName, xLogProperties.Password, xLogProperties.FolderAuthenticationType);
                            Microsoft.Office.Interop.Outlook.Application outlookObj = Globals.ThisAddIn.Application;
                            //Gte MAPI Name space
                            Microsoft.Office.Interop.Outlook.NameSpace outlookNameSpac = outlookObj.GetNamespace("MAPI");

                            oInBox = outlookNameSpac.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                            parentFolder = (Microsoft.Office.Interop.Outlook.MAPIFolder)oInBox.Parent;
                            Microsoft.Office.Interop.Outlook.MAPIFolder f = MAPIFolderWrapper.GetFolder(parentFolder, folderName);
                            f.WebViewOn = true;
                            frmConnectionProperties frmconnection = new frmConnectionProperties();
                            if (parentFolder.Name.Trim() != folderName.Trim())
                            {
                                if (f.Name.Trim() != parentFolder.Name.Trim())
                                {
                                    if (f != null)
                                    {

                                        f.WebViewURL = frmSPSiteConfigurationObject.URL;
                                        if (f.FolderPath.Contains("\\Deleted Items\\"))
                                        {
                                            try
                                            {
                                                f.Delete();
                                            }
                                            catch (Exception)
                                            {


                                            }
                                            frmconnection.CreateFolder(folderName, xLogProperties);
                                            Outlook.MAPIFolder ChildFolder = MAPIFolderWrapper.GetFolder(parentFolder, folderName);
                                            if (ChildFolder != null && ChildFolder.Name == folderName)
                                            {
                                                //  MAPIFolderWrapper omapi = null;
                                                if (string.IsNullOrEmpty(DocLibName) == true)
                                                {

                                                    omapi = new MAPIFolderWrapper(ref  ChildFolder, this.Application.ActiveExplorer(), false);

                                                }
                                                else
                                                {
                                                    omapi = new MAPIFolderWrapper(ref  ChildFolder, this.Application.ActiveExplorer(), true);


                                                }


                                                string returl = UserLogManagerUtility.GetSPSiteURL(omapi.AttachedFolder.Name);

                                                omapi.AttachedFolder.WebViewURL = ListWebClass.WebViewUrl(returl);

                                                myFolders.Add(omapi);
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    frmconnection.CreateFolder(folderName, xLogProperties);
                                    Outlook.MAPIFolder ChildFolder = MAPIFolderWrapper.GetFolder(parentFolder, folderName);
                                    if (ChildFolder != null && ChildFolder.Name == folderName)
                                    {
                                        //  MAPIFolderWrapper omapi = null;
                                        if (string.IsNullOrEmpty(DocLibName) == true)
                                        {

                                            omapi = new MAPIFolderWrapper(ref  ChildFolder, this.Application.ActiveExplorer(), false);

                                        }
                                        else
                                        {
                                            omapi = new MAPIFolderWrapper(ref  ChildFolder, this.Application.ActiveExplorer(), true);


                                        }


                                        string returl = UserLogManagerUtility.GetSPSiteURL(omapi.AttachedFolder.Name);

                                        omapi.AttachedFolder.WebViewURL = ListWebClass.WebViewUrl(returl);

                                        myFolders.Add(omapi);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ListWebClass.Log(ex.Message, true);
                    }
                }

            }

        }



        private void AddFolderInsession()
        {
            try
            {

            }
            catch (Exception)
            {


            }
        }
        private void CreateDefaultAddEventOnFolders()
        {


            Outlook.MAPIFolder oInBox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.MAPIFolder parentFolder = (Outlook.MAPIFolder)myTargetFolder; //oInBox.Parent;
            //Get all the folder information
            XmlNodeList xFolders = UserLogManagerUtility.GetDefaultFoldersDetails(UserStatus.Active);
            if (xFolders != null)
            {
                string folderName = string.Empty, DocLibName = string.Empty;
                if (addinExplorer == null)
                {
                    MessageBox.Show("In Main Forms");
                }

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

                            MAPIFolderWrapper omapi = null;
                            if (string.IsNullOrEmpty(DocLibName) == true)
                            {
                                //Doc name is empty means Folder is not mapped with Doc Lib
                                //omapi = new MAPIFolderWrapper(oChildFolder, addinExplorer, false);
                                omapi = new MAPIFolderWrapper(ref  oChildFolder, this.Application.ActiveExplorer(), false);

                            }
                            else
                            {
                                omapi = new MAPIFolderWrapper(ref  oChildFolder, this.Application.ActiveExplorer(), true);
                                //omapi = new MAPIFolderWrapper(oChildFolder, addinExplorer, true);

                            }

                            //
                            string returl = UserLogManagerUtility.GetSPSiteURL(omapi.AttachedFolder.Name);
                            //string relativepath = UserLogManagerUtility.GetRelativePath(omapi.AttachedFolder.Name);
                            //string rootpath = "";
                            //if (returl.LastIndexOf(relativepath)!=-1)
                            //{
                            //    rootpath = returl.Remove(returl.LastIndexOf(relativepath));
                            //}
                            // string virtualpath = ListWebClass.WebViewUrl(returl);

                            //  omapi.AttachedFolder.WebViewURL = rootpath + "/_layouts/OutlookIntegration/DisplayImage.aspx?Action=OLIssue&ReturnUrl=" + returl;
                            omapi.AttachedFolder.WebViewURL = ListWebClass.WebViewUrl(returl);
                            //
                            myFolders.Add(omapi);


                        }
                        else
                        {
                            //create mapi folder

                            XMLLogProperties xLogProperties = new XMLLogProperties();
                            xLogProperties.UserName = EncodingAndDecoding.Base64Decode(xNode.ChildNodes[0].InnerText);
                            xLogProperties.Password = EncodingAndDecoding.Base64Decode(xNode.ChildNodes[1].InnerText);
                            xLogProperties.DisplayFolderName = folderName;
                            xLogProperties.SiteURL = xNode.ChildNodes[4].InnerText;

                            xLogProperties.UsersStatus = UserStatus.Active;
                            xLogProperties.DocumentLibraryName = xNode.ChildNodes[3].InnerText;
                            xLogProperties.DroppedURLType = "";

                            if (xNode.ChildNodes[5].InnerText == "Manually Specified")
                            {
                                xLogProperties.FolderAuthenticationType = AuthenticationType.Manual;
                            }
                            else
                            {
                                xLogProperties.FolderAuthenticationType = AuthenticationType.Domain;

                            }

                            xLogProperties.SPSiteVersion = SPVersionClass.GetSPVersionFromUrl(xLogProperties.SiteURL, xLogProperties.UserName, xLogProperties.Password, xLogProperties.FolderAuthenticationType);
                            Microsoft.Office.Interop.Outlook.Application outlookObj = Globals.ThisAddIn.Application;
                            //Gte MAPI Name space
                            Microsoft.Office.Interop.Outlook.NameSpace outlookNameSpac = outlookObj.GetNamespace("MAPI");

                            oInBox = outlookNameSpac.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                            parentFolder = (Microsoft.Office.Interop.Outlook.MAPIFolder)oInBox.Parent;
                            Microsoft.Office.Interop.Outlook.MAPIFolder f = MAPIFolderWrapper.GetFolder(parentFolder, folderName);
                            frmConnectionProperties frmconnection = new frmConnectionProperties();
                            if (parentFolder.Name.Trim() != folderName.Trim())
                            {
                                if (f.Name.Trim() != parentFolder.Name.Trim())
                                {
                                    if (f != null)
                                    {

                                        f.WebViewURL = frmSPSiteConfigurationObject.URL;
                                        if (f.FolderPath.Contains("\\Deleted Items\\"))
                                        {
                                            try
                                            {
                                                f.Delete();
                                            }
                                            catch (Exception)
                                            {


                                            }
                                            frmconnection.CreateFolder(folderName, xLogProperties);
                                        }
                                    }
                                }
                                else
                                {
                                    frmconnection.CreateFolder(folderName, xLogProperties);
                                }
                            }
                        }
                    }
                    catch (Exception ex) { }
                }

            }

        }
        #endregion

        #region Folder Context menu events and methods


        /// <summary>
        /// <c>Application_FolderContextMenuDisplay</c> Event Handler
        /// Folder context menu display event
        /// it is executed when user right clicks on outlook folder to display context menu 
        /// </summary>
        /// <param name="CommandBar"></param>
        /// <param name="Folder"></param>
        void Application_FolderContextMenuDisplay(Microsoft.Office.Core.CommandBar CommandBar, Microsoft.Office.Interop.Outlook.MAPIFolder Folder)
        {
            try
            {
                currentFolderSelected = Folder.Name;

                currentFolderSelectedGuid = Folder.EntryID;
                AddfFolderinSessionMapi();
                MAPIFolderWrapper folderWrapper = myFolders.Find(delegate(MAPIFolderWrapper p) { return p.FolderName == Folder.Name; });

                if (folderWrapper != null && !string.IsNullOrEmpty(Folder.WebViewURL))
                {

                    oContextMenuFolder = Folder;
                    try
                    {
                        CreateEventToRenameOnFolderContextMenu(CommandBar, Folder);
                        //Remove the click event on the button
                        oFolderMenuButtonEditConnectionProperties.Click -= new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(oFolderMenuButtonEditConnectionProperties_Click);

                    }
                    catch { }

                    //System.Diagnostics.Debugger.Launch();



                    //Add Edit Connection properties button
                    oFolderMenuButtonEditConnectionProperties = (Office.CommandBarButton)CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                    oFolderMenuButtonEditConnectionProperties.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    oFolderMenuButtonEditConnectionProperties.Caption = "Edit Connection (ITOPIA)";
                    oFolderMenuButtonEditConnectionProperties.FaceId = 222;
                    oFolderMenuButtonEditConnectionProperties.Tag = Folder.Name;
                    oFolderMenuButtonEditConnectionProperties.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(oFolderMenuButtonEditConnectionProperties_Click);

                    //Add SharePoint View Context Menu button
                    oFolderMenuButtonSharePointView = (Office.CommandBarButton)CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                    oFolderMenuButtonSharePointView.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    oFolderMenuButtonSharePointView.Caption = "SharePoint View (ITOPIA)";
                    oFolderMenuButtonSharePointView.FaceId = 223;
                    oFolderMenuButtonSharePointView.Tag = Folder.Name + "A";
                    oFolderMenuButtonSharePointView.Click += new Office._CommandBarButtonEvents_ClickEventHandler(oFolderMenuButtonSharePointView_Click);

                    //Add Outlook View Context Menu button
                    oFolderMenuButtonOutlookView = (Office.CommandBarButton)CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                    oFolderMenuButtonOutlookView.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    oFolderMenuButtonOutlookView.Caption = "Outlook View (ITOPIA)";
                    oFolderMenuButtonOutlookView.FaceId = 224;
                    oFolderMenuButtonOutlookView.Tag = Folder.Name + "B";
                    oFolderMenuButtonOutlookView.Click += new Office._CommandBarButtonEvents_ClickEventHandler(oFolderMenuButtonOutlookView_Click);


                    ////Add view Connection properties button
                    //oFolderMenuButtonViewProperties = (Office.CommandBarButton)CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                    //oFolderMenuButtonViewProperties.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    //oFolderMenuButtonViewProperties.Caption = "ITOPIA Edit Connection Properties";
                    //oFolderMenuButtonViewProperties.FaceId = 222;
                    //oFolderMenuButtonViewProperties.Tag = Folder.Name;
                    //oFolderMenuButtonViewProperties.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(oFolderMenuButtonEditConnectionProperties_Click);
                    ////

                    Outlook.MAPIFolder myfolder = (Outlook.MAPIFolder)Folder.Parent;
                    try
                    {
                        myfolder.Folders.FolderChange -= new Outlook.FoldersEvents_FolderChangeEventHandler(oMailRootFolders_FolderChange);
                    }
                    catch (Exception)
                    {


                    }

                    myfolder.Folders.FolderChange += new Outlook.FoldersEvents_FolderChangeEventHandler(oMailRootFolders_FolderChange);


                }
            }
            catch (Exception ex)
            {
                EncodingAndDecoding.ShowMessageBox("Application_FolderContextMenuDisplay", ex.Message, MessageBoxIcon.Error);
            }

        }


        /// <summary>
        /// <c>oFolderMenuButtonOutlookView_Click</c> Event Handler
        /// this event is executed  when user selects outlook view from context menu
        /// it opens the current foler in outlook view
        /// </summary>
        /// <param name="Ctrl"></param>
        /// <param name="CancelDefault"></param>
        void oFolderMenuButtonOutlookView_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                string folderName = ((Microsoft.Office.Core.CommandBarButtonClass)(Ctrl)).Tag;
                folderName = folderName.Substring(0, folderName.Length - 1);
                this.Application.ActiveExplorer().CurrentFolder.WebViewOn = false;
                this.Application.ActiveExplorer().CurrentFolder = this.Application.ActiveExplorer().CurrentFolder;
            }
            catch (Exception ex)
            { }
        }


        /// <summary>
        /// <c>oFolderMenuButtonSharePointView_Click</c> Event Handler
        /// this event is executed  when user selects Sharepoint view from context menu
        /// it opens the current foler in Sharepoint view
        /// </summary>
        /// <param name="Ctrl"></param>
        /// <param name="CancelDefault"></param>
        void oFolderMenuButtonSharePointView_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {

                string folderName = ((Microsoft.Office.Core.CommandBarButtonClass)(Ctrl)).Tag;
                folderName = folderName.Substring(0, folderName.Length - 1);
                this.Application.ActiveExplorer().CurrentFolder.WebViewOn = true;

                this.Application.ActiveExplorer().CurrentFolder = this.Application.ActiveExplorer().CurrentFolder;
            }
            catch (Exception ex)
            { }
        }

        /// <summary>
        /// <c>CreateEventToRenameOnFolderContextMenu</c> Event Handler
        /// Method to  create rename event of folder context menu
        /// </summary>
        /// <param name="CommandBar"></param>
        /// <param name="Folder"></param>
        private void CreateEventToRenameOnFolderContextMenu(Microsoft.Office.Core.CommandBar CommandBar, Microsoft.Office.Interop.Outlook.MAPIFolder Folder)
        {


            for (int i = 1; i <= CommandBar.Controls.Count; i++)
            {
                if (CommandBar.Controls[i].Caption.StartsWith("&Rename"))
                {
                    try
                    {
                        renameButton.Click -= new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(renameButton_Click);
                    }
                    catch { }
                    renameButton = null;
                    renameButton = (Office.CommandBarButton)CommandBar.Controls[i];
                    renameButton.Tag = Folder.Name;

                    renameButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(renameButton_Click);
                    break;
                }
            }


        }

        /// <summary>
        /// Custome menu item "Itopia Edit proeprties" click event
        /// </summary>
        /// <param name="Ctrl"></param>
        /// <param name="CancelDefault"></param>
        void oFolderMenuButtonEditConnectionProperties_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                try
                {
                    renameButton.Click -= new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(renameButton_Click);
                }
                catch { }

                //Create instance
                frmSPSiteConfigurationObject = new frmSPSiteConfiguration();
                //Get the drop url
                string folderName = ((Microsoft.Office.Core.CommandBarButtonClass)(Ctrl)).Tag;
                frmSPSiteConfigurationObject.ShowEditForm(folderName);


                if (frmSPSiteConfigurationObject.IsConfigureCompleted)
                {
                    UserLogManagerUtility.UpdateFolderConfigDetails("", frmSPSiteConfigurationObject.FolderConfigProperties);
                    foreach (MAPIFolderWrapper folder in myFolders)
                    {
                        if (folder.FolderName == folderName)
                        {
                            folder.AttachedFolder.WebViewURL = frmSPSiteConfigurationObject.URL;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                EncodingAndDecoding.ShowMessageBox("oFolderMenuButtonConnectionProperties_Click", ex.Message, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// <c>renameButton_Click</c> member function
        /// Rename click event.
        /// This method updates the renamed folder name in Configuration File
        /// </summary>
        /// <param name="Ctrl"></param>
        /// <param name="CancelDefault"></param>
        void renameButton_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                try
                {
                    folderNewName = string.Empty;
                    //Remove the click event on the button
                    oFolderMenuButtonEditConnectionProperties.Click -= new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(oFolderMenuButtonEditConnectionProperties_Click);
                }
                catch { }

                frmRenameObject = new frmRename();
                frmRenameObject.ShowDialog();
                folderNewName = frmRenameObject.txtNewFolderName.Text;



                if (!string.IsNullOrEmpty(folderNewName))
                {
                    string oldName = oContextMenuFolder.Name;
                    MAPIFolderWrapper folderWrapper = myFolders.Find(delegate(MAPIFolderWrapper p) { return p.FolderName == oldName; });
                    if (folderWrapper != null)
                    {
                        oContextMenuFolder.Name = folderNewName;

                        //MAPIFolderWrapper folderWrapper = myFolders.Find(delegate(MAPIFolderWrapper p) { return p.FolderName == renameButton.Tag; });

                        folderWrapper.AttachedFolder = (Outlook.MAPIFolder)oContextMenuFolder;
                        folderWrapper.FolderName = folderNewName;

                        UserLogManagerUtility.UpdateFolderConfigNodeDetails(oldName, "DisplayName", folderNewName);
                        EncodingAndDecoding.ShowMessageBox("renameButton_Click", "Folder Renamed sucessfully", MessageBoxIcon.Information);

                    }
                }
                else
                {
                    CancelDefault = true;
                }
                updatefolderlocationin();

            }
            catch (Exception ex)
            {
                EncodingAndDecoding.ShowMessageBox("renameButton_Click", ex.Message, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Context menu close event
        /// not  currently used
        /// </summary>
        /// <param name="ContextMenu"></param>
        void Application_ContextMenuClose(Microsoft.Office.Interop.Outlook.OlContextMenu ContextMenu)
        {
            //if (ContextMenu.ToString() == "olFolderContextMenu")
            //{
            //    if (!string.IsNullOrEmpty(folderNewName))
            //    {
            //        oContextMenuFolder.Name = folderNewName;

            //    }
            frmRenameObject = null;
            // }
        }

        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {

            return base.CreateRibbonExtensibilityObject();
            //  return new Microsoft.Office.Tools.Ribbon.RibbonManager(new Microsoft.Office.Tools.Ribbon.OfficeRibbon[] { new SharePointRibbon() });


        }


        /// <summary>
        /// <c>ConvertImage</c> class 
        /// implements the functionality to convert image to outlook compatible format to be displayed in  outlook ribbon
        /// </summary>
        sealed public class ConvertImage : System.Windows.Forms.AxHost
        {
            private ConvertImage()
                : base(null)
            {

            }
            public static stdole.IPictureDisp Convert
                (System.Drawing.Image image)
            {
                return (stdole.IPictureDisp)System.
                    Windows.Forms.AxHost
                    .GetIPictureDispFromPicture(image);
            }
        }

        /// <summary>
        /// <c>getImage</c> member method 
        /// This function gets image from resource file and calls <c>ConvertImage.Convert</c>
        /// method to convert it in outlook compatible format
        /// </summary>
        /// <returns></returns>
        private stdole.IPictureDisp getImage()
        {
            stdole.IPictureDisp tempImage = null;
            try
            {
                System.Drawing.Bitmap newIcon =
                Properties.Resources.exclamation_red;

                ImageList newImageList = new ImageList();
                newImageList.Images.Add(newIcon);
                tempImage = ConvertImage.Convert(newImageList.Images[0]);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return tempImage;
        }

        /// <summary>
        /// <c>NewConnection</c> member function 
        /// returns true/false on the basis of folder whether it is recently created or not
        /// </summary>
        /// <param name="xLogProperties"></param>
        /// <returns></returns>
        public bool NewConnection(XMLLogProperties xLogProperties)
        {
            bool result = false;
            try
            {
                Outlook.MAPIFolder newFolder = null;
                ////////////////////////updated by Joy on 25.07.2012/////////////////
                Outlook.MAPIFolder newBrokenUploadsFolder = null;
                ////////////////////////updated by Joy on 25.07.2012/////////////////
                //outlookObj = new Outlook.Application();
                OutlookObj = Globals.ThisAddIn.Application;

                //Gte MAPI Name space
                outlookNameSpace = OutlookObj.GetNamespace("MAPI");
                Outlook.MAPIFolder olInboxFolder = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Outlook.MAPIFolder Target = (Outlook.MAPIFolder)olInboxFolder.Parent;

                bool created = CreateFolderInOutLookSideMenu(xLogProperties.DisplayFolderName, xLogProperties.SiteURL, out newFolder, Target);
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////////////////////Modified by Joy:25.07.2012///////////////////////////////////////////////
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                
                if (created == true && newFolder != null)
                {


                    //Set new folder location
                    xLogProperties.OutlookFolderLocation = newFolder.FolderPath;
                    //Create node in xml file
                    UserLogManagerUtility.CreateXMLFileForStoringUserCredentials(xLogProperties);

                    MAPIFolderWrapper omapi = null;
                    if (string.IsNullOrEmpty(xLogProperties.DocumentLibraryName) == true)
                    {
                        //Doc name is empty means Folder is not mapped with Doc Lib
                        omapi = new MAPIFolderWrapper(ref  newFolder, addinExplorer, false);
                    }
                    else
                    {
                        omapi = new MAPIFolderWrapper(ref newFolder, addinExplorer, true);
                    }

                    newFolders.Add(omapi);

                }


            }
            catch (Exception ex)
            {


            }
            return result;
        }

        /// <summary>
        /// <c>ReConnection</c> member function
        /// this function recreates the deleted  folder in outlook 
        /// </summary>
        /// <param name="xLogProperties"></param>
        /// <param name="parentfolder"></param>
        /// <returns></returns>
        public bool ReConnection(XMLLogProperties xLogProperties, Outlook.MAPIFolder parentfolder)
        {
            bool result = false;
            try
            {



                Outlook.MAPIFolder newFolder = null;
                ////////////////////////updated by Joy on 25.07.2012/////////////////
                Outlook.MAPIFolder newBrokenUploadsFolder = null;
                ////////////////////////updated by Joy on 25.07.2012/////////////////
                //outlookObj = new Outlook.Application();
                OutlookObj = Globals.ThisAddIn.Application;

                //Gte MAPI Name space
                outlookNameSpace = OutlookObj.GetNamespace("MAPI");
                Outlook.MAPIFolder olInboxFolder = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                Outlook.MAPIFolder Target = parentfolder;

                bool created = CreateFolderInOutLookSideMenu(xLogProperties.DisplayFolderName, xLogProperties.SiteURL, out newFolder, Target);
                
                if (created == true && newFolder != null)
                {


                    //Set new folder location
                    xLogProperties.OutlookFolderLocation = newFolder.FolderPath;
                    //Create node in xml file
                    //  UserLogManagerUtility.CreateXMLFileForStoringUserCredentials(xLogProperties);

                    MAPIFolderWrapper omapi = null;
                    if (string.IsNullOrEmpty(xLogProperties.DocumentLibraryName) == true)
                    {
                        //Doc name is empty means Folder is not mapped with Doc Lib
                        omapi = new MAPIFolderWrapper(ref  newFolder, addinExplorer, false);
                    }
                    else
                    {
                        omapi = new MAPIFolderWrapper(ref newFolder, addinExplorer, true);
                    }

                    newFolders.Add(omapi);

                }


            }
            catch (Exception ex)
            {


            }
            return result;
        }


        /// <summary>
        /// <c>AddfFolderinSessionMapi</c> member function 
        /// this member function add the newly created folder in  session folder Collection
        /// </summary>
        public void AddfFolderinSessionMapi()
        {
            try
            {
                foreach (MAPIFolderWrapper r in newFolders)
                {
                    MAPIFolderWrapper folderWrapper = myFolders.Find(delegate(MAPIFolderWrapper p) { return p.FolderName == r.FolderName; });
                    if (folderWrapper != null)
                    {
                        newFolders.Remove(r);
                    }
                    else
                    {
                        r.AttachedFolder.WebViewURL = ListWebClass.WebViewUrl(r.AttachedFolder.WebViewURL);
                        myFolders.Add(r);
                        // newFolders.Remove(r);
                    }
                }

            }
            catch (Exception)
            { }

        }


    }
}

using System;
using System.Text;
using System.Windows.Forms;
using Interfaces;
using System.Runtime.InteropServices;
using System.Xml;
using SharePoint_Link.Utility;
using Utility;

namespace SharePoint_Link
{
    // group type enum
    /// <summary>
    /// <c>SECURITY_IMPERSONATION_LEVEL</c> enum
    /// holds possible value of SECURITY_IMPERSONATION_LEVEL
    /// </summary>
    public enum SECURITY_IMPERSONATION_LEVEL : int
    {
        SecurityAnonymous = 0,
        SecurityIdentification = 1,
        SecurityImpersonation = 2,
        SecurityDelegation = 3
    }

    /// <summary>
    /// <c>frmEditUploadedItem</c> class inherits <c>Form, IOleClientSite, Interfaces.IServiceProvider, IAuthenticate</c>
    /// Implements the functionalities related to display uploaded file metatags information
    /// </summary>
    public partial class frmEditUploadedItem : Form, IOleClientSite, Interfaces.IServiceProvider, IAuthenticate
    {

        /// <summary>
        /// <c>IID_IAuthenticate</c> member field of type <c>Guid</c>
        /// </summary>
        public static Guid IID_IAuthenticate = new Guid("79eac9d0-baf9-11ce-8c82-00aa004ba90b");

        /// <summary>
        /// <c>INET_E_DEFAULT_ACTION</c> member field of  type  int
        /// </summary>
        public const int INET_E_DEFAULT_ACTION = unchecked((int)0x800C0011);

        /// <summary>
        /// <c>S_OK</c> member field of type 
        /// </summary>
        public const int S_OK = unchecked((int)0x00000000);

        /// <summary>
        /// <c>m_strUser</c> member  field of type string
        /// holds the user name
        /// <c>m_strPwd</c> member field of type string
        /// holds the password of user
        /// </summary>
        public string m_strUser = string.Empty, m_strPwd = string.Empty;

        /// <summary>
        /// <c>file_Name</c> member field of type string 
        /// holds the name of file being uploaded
        /// </summary>
        public string file_Name = "";

        /// <summary>
        /// <c>file_url</c> member field of type string .
        /// holds the url of uploaded file
        /// </summary>
        public string file_url = "";

        /// <summary>
        /// <c>attempt</c> member field of type int
        /// </summary>
        public int attempt = 0;

        /// <summary>
        /// <c>frmEditUploadedItem</c> default constructor
        /// </summary>
        public frmEditUploadedItem()
        {
            InitializeComponent();
            string oURL = "about:blank";
            webBrowser1.Navigate(oURL);
        }


        /// <summary>
        /// <c>frmUploadItemsSettings_Load</c> event handler
        /// display message "please wait.."
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmUploadItemsSettings_Load(object sender, EventArgs e)
        {

            this.Text = "Please Wait...";
        }

        /// <summary>
        /// <c>ShowWithBrowser</c> member function
        /// Navigate to the url of document library
        /// </summary>
        /// <param name="folderName"></param>
        /// <param name="url"></param>
        public void ShowWithBrowser(string folderName, string url)
        {

            //Notify the WebBrowser object about the client site.   
            //The client site, informs an embedded object of its display location   
            object obj = webBrowser1.ActiveXInstance;
            IOleObject oc = obj as IOleObject;
            oc.SetClientSite(this as IOleClientSite);


            XmlNode folderNode = UserLogManagerUtility.GetSPSiteURLDetails("", folderName);
            m_strUser = EncodingAndDecoding.Base64Decode(folderNode.ChildNodes[0].InnerText);
            m_strPwd = EncodingAndDecoding.Base64Decode(folderNode.ChildNodes[1].InnerText);
            webBrowser1.Navigate(url);
            this.ShowDialog();
        }

        /// <summary>
        /// <c>ShowWithBrowser</c> overloaded member function
        /// Navigate to the url of document library
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="pwd"></param>
        /// <param name="url"></param>
        public void ShowWithBrowser(string userName, string pwd, string url)
        {

            this.Text = "Please Wait...";

            try
            {
                int index = url.ToString().LastIndexOf("&UpFName=");
                if (index != -1)
                {
                    file_Name = url.Substring(index + 9);
                    url = url.Remove(index);
                    file_url = url;
                }

            }
            catch (Exception)
            {


            }
            //Notify the WebBrowser object about the client site.   
            //The client site, informs an embedded object of its display location   

            object obj = webBrowser1.ActiveXInstance;
            IOleObject oc = obj as IOleObject;
            oc.SetClientSite(this as IOleClientSite);
            this.TopLevel = true;
            this.TopMost = true;
            m_strUser = userName;
            m_strPwd = pwd;

            webBrowser1.Navigate(url);

            this.ShowDialog();

        }

        #region IOleClientSite Members

        /// <summary>
        /// <c>SaveObject</c> interface method implementation
        /// </summary>
        void IOleClientSite.SaveObject()
        {

        }


        /// <summary>
        /// <c>GetMoniker</c>  <c>IOleClientSite</c> interface overloaded method
        /// </summary>
        /// <param name="dwAssign"></param>
        /// <param name="dwWhichMoniker"></param>
        /// <param name="ppmk"></param>
        void IOleClientSite.GetMoniker(uint dwAssign, uint dwWhichMoniker, ref object ppmk)
        {

        }


        /// <summary>
        /// <c>GetContainer</c>  IOleClientSite interface method implementation
        /// </summary>
        /// <param name="ppContainer"></param>
        void IOleClientSite.GetContainer(ref object ppContainer)
        {

        }


        /// <summary>
        /// <c>ShowObject</c> IOleClientSite interface method implementation
        /// </summary>
        void IOleClientSite.ShowObject()
        {

        }


        /// <summary>
        /// <c>OnShowWindow</c> IOleClientSite interface method implementation
        /// </summary>
        /// <param name="fShow"></param>
        void IOleClientSite.OnShowWindow(bool fShow)
        {

        }


        /// <summary>
        /// <c>RequestNewObjectLayout</c> IOleClientSite interface method implementation
        /// </summary>
        void IOleClientSite.RequestNewObjectLayout()
        {

        }

        #endregion
         
        #region IServiceProvider Members


        /// <summary>
        /// <c>QueryService</c>  IServiceProvider interface method implementation
        /// </summary>
        /// <param name="guidService"></param>
        /// <param name="riid"></param>
        /// <param name="ppvObject"></param>
        /// <returns></returns>
        int Interfaces.IServiceProvider.QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObject)
        {
            ppvObject = new IntPtr();
            try
            {
                int nRet = guidService.CompareTo(IID_IAuthenticate);
                // Zero //returned if the compared objects are equal 
                if (nRet == 0)
                {
                    nRet = riid.CompareTo(IID_IAuthenticate);
                    // Zero //returned if the compared objects are equal 
                    if (nRet == 0)
                    {
                        ppvObject = Marshal.GetComInterfaceForObject(this, typeof(IAuthenticate));
                        return S_OK;
                    }
                }
            }
            catch (Exception ex)
            { }
            return INET_E_DEFAULT_ACTION;
        }

        #endregion

        #region IAuthenticate Members

        /// <summary>
        /// <c>Authenticate</c>  IAuthenticate  interface method implementation
        /// Authenticate users
        /// </summary>
        /// <param name="phwnd"></param>
        /// <param name="pszUsername"></param>
        /// <param name="pszPassword"></param>
        /// <returns></returns>
        int IAuthenticate.Authenticate(ref IntPtr phwnd, ref IntPtr pszUsername, ref IntPtr pszPassword)
        {
            try
            {
                IntPtr sUser = Marshal.StringToCoTaskMemAuto(m_strUser);
                IntPtr sPassword = Marshal.StringToCoTaskMemAuto(m_strPwd);


                pszUsername = sUser;
                pszPassword = sPassword;
            }
            catch (Exception ex)
            { }
            return S_OK;
        }

        #endregion


        /// <summary>
        /// <c>webBrowser1_DocumentCompleted</c> Event handler
        /// display  metatags form of the file uploaded 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            HideProgress(false);

            this.Text = "Please Wait...";
            if (webBrowser1.Url.OriginalString != "about:blank")
            {
                if (webBrowser1.Document.Title.StartsWith("No item exists at") || webBrowser1.Document.Title == "Error")
                {
                    StringBuilder sb1 = new StringBuilder();
                    sb1.Append("Unable to open the item. ");
                    sb1.Append("Follwing are the reasons\n\r\n");
                    sb1.Append("1. Item may have been deleted or renamed by another user.\n\r\n");
                    sb1.Append("2. List contains meta tags as required fields along with verison enabled.\r\n");
                    sb1.Append("    You need to update meta tags inforamtion manually and check-in using sharepoint.");
                    EncodingAndDecoding.ShowMessageBox("", sb1.ToString(), MessageBoxIcon.Information);
                    this.Close();

                }
                else
                {
                    if (!string.IsNullOrEmpty(file_Name))
                    {


                        ////
                        if (e.Url.ToString().Contains(file_Name) || e.Url.ToString().Contains("AllItems.aspx"))
                        {
                            string ul = file_url;
                            int indxedit = ul.LastIndexOf("?ID=");
                            if (indxedit != -1)
                            {

                                ul = ul.Remove(indxedit);
                                ul = ul.Replace("EditForm.aspx", "AllItems.aspx");
                                this.Close();
                                //  webBrowser1.Navigate(ul);
                                HideProgress(true);
                                EncodingAndDecoding.ShowMessageBox("", "MetaTag Updated successfully", MessageBoxIcon.Information);
                            }

                        }
                        else
                        {

                            if (attempt > 0)
                            {
                                if (e.Url.ToString().Trim() == file_url.Trim())
                                {
                                    this.Close();
                                    HideProgress(true);
                                    EncodingAndDecoding.ShowMessageBox("", "MetaTag Updated successfully", MessageBoxIcon.Information);
                                }
                                else
                                {

                                    if (e.Url.ToString().Contains(file_Name.Trim()))
                                    {


                                        this.Close();
                                        HideProgress(true);
                                        EncodingAndDecoding.ShowMessageBox("", "MetaTag Updated successfully", MessageBoxIcon.Information);

                                    }


                                }


                            }


                            webBrowser1.Visible = true;
                        }

                    }
                    else
                    {
                        webBrowser1.Visible = true;
                    }

                    attempt += 1;
                }

                SetCancetScript();
                this.Text = "Edit Properties";
            }
            HideProgress(true);


        }


        /// <summary>
        /// <c>SetCancetScript</c> member function
        /// update javascript  functions in browsed  file
        /// </summary>
        private void SetCancetScript()
        {
            try
            {

                System.Windows.Forms.HtmlElementCollection mycollection = webBrowser1.Document.GetElementsByTagName("input");
                if (mycollection.Count > 0)
                {
                    for (int i = 0; i < mycollection.Count; i++)
                    {
                        string ty = mycollection[i].GetAttribute("type");
                        if (ty.Trim() == "button")
                        {
                            string vale = mycollection[i].GetAttribute("value").Trim();
                            if (vale.ToLower() == "cancel")
                            {
                                mycollection[i].Click += new HtmlElementEventHandler(Form1_Click);
                            }

                        }

                    }

                }
            }
            catch (Exception)
            {


            }
        }


        /// <summary>
        /// <c>Form1_Click</c> event handler 
        /// close the window form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Click(object sender, System.Windows.Forms.HtmlElementEventArgs e)
        {
            this.Close();
        }


        /// <summary>
        /// <c>webBrowser1_Navigated</c> event handler 
        /// web browser event
        /// calls <c>HideProgress</c> function to display progress bar
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            try
            {
                HideProgress(false);
                this.Text = "Please wait..";
            }
            catch (Exception)
            {


            }
        }


        /// <summary>
        /// <c>webBrowser1_Navigating</c> event handler
        /// web browser event 
        /// calls the <c>HideProgress</c> member function to display or hide 
        /// progress bar
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void webBrowser1_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            try
            {
                HideProgress(false);
                this.Text = "Please wait";
            }
            catch (Exception)
            {


            }
        }

        /// <summary>
        /// <c>HideProgress</c> member function
        /// hide or display progress bar status 
        /// </summary>
        /// <param name="hide"></param>
        private void HideProgress(bool hide)
        {
            try
            {
                if (hide == true)
                {
                    pictureBox1.Visible = false;
                    lblprogress.Visible = false;
                }
                else
                {
                    pictureBox1.Visible = true;
                    lblprogress.Visible = true;
                }
            }
            catch (Exception)
            {


            }
        }




    }
}
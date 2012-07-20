using System;
using System.ComponentModel;
using System.Windows.Forms;
using Interfaces;
using System.Runtime.InteropServices;

namespace SharePoint_Link
{
    /// <summary>
    /// <c>frmtest</c> class inherits <c> Form, IOleClientSite, Interfaces.IServiceProvider, IAuthenticate</c>
    /// handles the functionality to display  please wait message
    /// </summary>
    public partial class frmtest : Form, IOleClientSite, Interfaces.IServiceProvider, IAuthenticate
    {
        bool isFormClosed = false;

        /// <summary>
        /// <c>IID_IAuthenticate</c> member field of type guid
        /// </summary>
        public static Guid IID_IAuthenticate = new Guid("79eac9d0-baf9-11ce-8c82-00aa004ba90b");

        /// <summary>
        /// <c>INET_E_DEFAULT_ACTION</c> member field of type int
        /// </summary>
        public const int INET_E_DEFAULT_ACTION = unchecked((int)0x800C0011);

        /// <summary>
        /// <c>S_OK</c> member field of type int
        /// </summary>
        public const int S_OK = unchecked((int)0x00000000);

        /// <summary>
        /// <c>m_strUser</c> member field of type string 
        /// holds username
        /// <c>m_strPwd</c> member field of type string
        /// holds password
        /// <c>URL</c> member field of type string
        /// holds the url of mapped library
        /// </summary>
        public string m_strUser = string.Empty, m_strPwd = string.Empty, URL = string.Empty;



        #region IOleClientSite Members

        /// <summary>
        /// <c>SaveObject</c>  IOleClientSite interface method implementation
        /// </summary>
        void IOleClientSite.SaveObject()
        {

        }

        /// <summary>
        /// <c>GetMoniker</c> IOleClientSite interface  method implementation
        /// </summary>
        /// <param name="dwAssign"></param>
        /// <param name="dwWhichMoniker"></param>
        /// <param name="ppmk"></param>
        void IOleClientSite.GetMoniker(uint dwAssign, uint dwWhichMoniker, ref object ppmk)
        {

        }

        /// <summary>
        /// <c>GetContainer</c> IOleClientSite interface method implementation
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
        /// <c>OnShowWindow</c> IOleClientSite interface method implementation.
        /// </summary>
        /// <param name="fShow"></param>
        void IOleClientSite.OnShowWindow(bool fShow)
        {

        }

        /// <summary>
        /// <c>RequestNewObjectLayout</c> IOleClientSite inerface method implementation
        /// </summary>
        void IOleClientSite.RequestNewObjectLayout()
        {

        }

        #endregion

        #region IServiceProvider Members

        /// <summary>
        /// <c>QueryService</c>IServiceProvider interface method implementation
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
        /// <c>Authenticate</c> IAuthenticate interface method implementation
        /// authenticate user
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
            authenticationCompleted = true;
            this.Hide();
            return S_OK;
        }

        #endregion


        /// <summary>
        /// <c>frmtest</c> default constructor
        /// </summary>
        public frmtest()
        {
            InitializeComponent();
        }

        /// <summary>
        /// <c>authenticationCompleted</c> member field of type bool
        /// holds the information whether the authentication process is completed or not.
        /// </summary>
        private bool authenticationCompleted = false;

        /// <summary>
        /// <c>AuthenticationCompleted</c> member property
        /// encapsulates authenticationCompleted member field
        /// </summary>
        public bool AuthenticationCompleted
        {
            get { return authenticationCompleted; }
            set { authenticationCompleted = value; }
        }


        /// <summary>
        /// <c>frmtest</c> browse to the folder url mapped to sharepoint library.
        /// </summary>
        /// <param name="uName"></param>
        /// <param name="pwd"></param>
        /// <param name="url"></param>
        public frmtest(string uName, string pwd, string url)
        {
            InitializeComponent();

            m_strUser = uName;
            m_strPwd = pwd;
            URL = url;
            string oURL = "about:blank";

            webBrowser2.Navigate(oURL);

        }

        /// <summary>
        /// <c>bgworker</c> an object of <c>BackgroundWorker</c> class
        /// </summary>
        BackgroundWorker bgworker = new BackgroundWorker();


        /// <summary>
        /// <c>frmtest_Load</c> event handler
        /// register <c>DoWork</c> and  <c>RunWorkerCompleted</c>
        /// Event handlers of <c>bgworker</c> and navigate to the sharepoint mapped url
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmtest_Load(object sender, EventArgs e)
        {
            bgworker.DoWork += new DoWorkEventHandler(bgworker_DoWork);
            bgworker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgworker_RunWorkerCompleted);
            AuthenticationCompleted = false;
            //Notify the WebBrowser object about the client site.   
            //The client site, informs an embedded object of its display location   

            object obj = webBrowser2.ActiveXInstance;
            IOleObject oc = obj as IOleObject;
            oc.SetClientSite(this as IOleClientSite);

            webBrowser2.Navigate(URL);
            dtFormLoadTime = DateTime.Now;
            bgworker.RunWorkerAsync();


        }

        /// <summary>
        /// <c>dtFormLoadTime</c> member field of type <c>DateTime</c>
        /// </summary>
        DateTime dtFormLoadTime = DateTime.Now;

        /// <summary>
        /// <c>bgworker_RunWorkerCompleted</c> event handler
        /// close the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void bgworker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            CloseForm();
        }

        /// <summary>
        /// <c>DelegateClose</c> delegate
        /// </summary>
        delegate void DelegateClose();


        delegate void DelegateCloseHelper();
        /// <summary>
        /// <c>CloseForm</c> member function
        /// close the form
        /// </summary>
        void CloseForm()
        {

            if (this.InvokeRequired)
            {
                DelegateCloseHelper d = new DelegateCloseHelper(CloseForm);
                this.Invoke(d, new object[] {  });
            }
            else
            {
                if (!isFormClosed)
                    this.Close();
            }


        }

        /// <summary>
        /// <c>bgworker_DoWork</c> event handler
        /// close the form if  authentication process is  completed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void bgworker_DoWork(object sender, DoWorkEventArgs e)
        {
            dtFormLoadTime = dtFormLoadTime.AddSeconds(2);
            do
            {
                if (authenticationCompleted == true && this.IsDisposed == true)
                {
                    this.Invoke(new DelegateClose(CloseForm));
                }
            } while (DateTime.Now < dtFormLoadTime);


        }

        /// <summary>
        /// <c>frmtest_VisibleChanged</c> event handler
        /// checks the authentication process and calls <c>Close</c> method to close the window form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmtest_VisibleChanged(object sender, EventArgs e)
        {
            if (authenticationCompleted == true)
            {
                this.Close();
            }
        }


        /// <summary>
        /// <c>webBrowser2_DocumentCompleted</c> Event Handler
        /// currently not in use
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void webBrowser2_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            //if (Convert.ToString(webBrowser2.Url) != "about:blank")
            //{
            //    HtmlDocument doc = webBrowser2.Document;
            //    HtmlElementCollection elements = doc.Links;

            //    foreach (HtmlElement item in elements)
            //    {
            //        if (item.InnerText.ToLower() == "sign in")
            //        {
            //            item.Click += new HtmlElementEventHandler(item_Click);
            //            item_Click(webBrowser2, null);
            //            //if (URL.Contains("/Forms/"))
            //            //{
            //            //    string s = URL.Substring(0, URL.LastIndexOf("/Forms/"));
            //            //    webBrowser2.Navigate(s + "/Forms/EditFrom.aspx?ID=0");
            //            //}
            //            return;
            //        }
            //    }
            //    //if (!string.IsNullOrEmpty(txtURL.Text))
            //    //{

            //    //    webBrowser1.Navigate(txtURL.Text);
            //    //}
            //}
        }

        private void frmtest_FormClosing(object sender, FormClosingEventArgs e)
        {
            isFormClosed = true;
        }

    }
}
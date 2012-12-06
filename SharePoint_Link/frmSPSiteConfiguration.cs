using System;
using System.Windows.Forms;
using SharePoint_Link.Utility;
using System.Web;
using Utility;
using System.Xml;
using SharePoint_Link.UserModule;
using System.Text.RegularExpressions;

namespace SharePoint_Link
{
    /// <summary>
    /// <c>frmSPSiteConfiguration</c> class inherits
    /// implements the functionality to create new sharepoint  mapped connection or modify 
    /// existing connection
    /// </summary>
    public partial class frmSPSiteConfiguration : Form
    {
        [System.Runtime.InteropServices.DllImport("User32")]
        private static extern bool SetForegroundWindow(
        IntPtr hWnd
        );

        #region Global Declaration & Constructor

        // commented by me  AuthenticationService.Authentication authenticateService = new AuthenticationService.Authentication();

        /// <summary>
        /// <c>xLogProperties</c> class object of <c>XMLLogProperties</c>
        /// required to handle configuration settings of connection
        /// </summary>
        XMLLogProperties xLogProperties = new XMLLogProperties();

        /// <summary>
        /// <c>frmSPSiteConfiguration</c> default constructor
        /// </summary>
        public frmSPSiteConfiguration()
        {
            InitializeComponent();
        }

        #endregion

        # region Properties
        /// <summary>
        /// <c>myVar</c> member field of type int 
        /// not currently in use
        /// </summary>
        readonly private int myVar;

        /// <summary>
        /// <c>FolderConfigProperties</c>member  property 
        /// encapsulates  xLogProperties object
        /// </summary>
        public XMLLogProperties FolderConfigProperties
        {
            get { return xLogProperties; }
            set { xLogProperties = value; }
        }


        /// <summary>
        /// <c>_URL</c> member field of type string
        /// holds the url of mapped connection 
        /// </summary>
        private string _URL = string.Empty;


        /// <summary>
        /// <c>URL</c> member  property 
        /// encapsulates <c>_URL</c>
        /// PRoperty to get the sharepoint site URL
        /// </summary>
        public string URL { get { return _URL; } set { _URL = value; } }

        /// <summary>
        /// <c>_isConfigureCompleted</c>  member field of type bool
        /// holds the  configuration status whether it is completed or not
        /// </summary>
        private bool _isConfigureCompleted = false;


        /// <summary>
        /// <c>IsConfigureCompleted</c> member property
        /// encapsulates _isConfigureCompleted
        /// Property to get configartion completed status
        /// </summary>
        public bool IsConfigureCompleted { get { return _isConfigureCompleted; } set { _isConfigureCompleted = value; } }


        /// <summary>
        /// <c>sPSiteVersion</c> member field of type string
        /// holds the sharepont version information
        /// </summary>
        private string sPSiteVersion = string.Empty;


        /// <summary>
        /// <c>SPSiteVersion</c> member  property 
        /// encapsulates sPSiteVersion
        /// </summary>
        public string SPSiteVersion
        {
            get { return sPSiteVersion; }
            set { sPSiteVersion = value; }
        }

        #endregion

        #region Methods

        /// <summary>
        /// <c>UpdateFolderConfigrationDetails</c> member function
        /// Method to assign properties to new configration folder
        /// </summary>
        private void UpdateFolderConfigrationDetails(bool checkFolderExists)
        {
            try
            {
                //Check the credentials are have the prmissions or not
                bool isValidCredentials = IsPassedCredentialHasAccess();
                if (isValidCredentials == true)
                {
                    URL = HttpUtility.UrlDecode(txtURL.Text);
                    xLogProperties.SiteURL = URL;
                    //version

                    Boolean isFolderExisted = false;
                    if (checkFolderExists == true)
                    {
                        //Check given folder name already existed or not
                        if (ThisAddIn.IsUrlIsTyped == true)
                        {
                            if (URL.Contains("/Forms/") == true)
                            {
                                string sTemp = URL.Substring(0, URL.IndexOf("/Forms/"));
                                xLogProperties.DocumentLibraryName = sTemp.Substring(sTemp.LastIndexOf("/") + 1);
                                xLogProperties.DroppedURLType = "/" + xLogProperties.DocumentLibraryName + "/Forms/";
                            }
                            else
                            {
                                xLogProperties.DocumentLibraryName = string.Empty;
                                xLogProperties.DroppedURLType = string.Empty;

                            }
                        }
                        else
                        {
                            isFolderExisted = UserLogManagerUtility.IsFolderExisted(txtDisplayName.Text.TrimEnd());
                        }


                    }
                    else
                    {

                        if (URL.Contains("/Forms/") == true)
                        {
                            string sTemp = URL.Substring(0, URL.IndexOf("/Forms/"));
                            xLogProperties.DocumentLibraryName = sTemp.Substring(sTemp.LastIndexOf("/") + 1);
                            xLogProperties.DroppedURLType = "/" + xLogProperties.DocumentLibraryName + "/Forms/";
                        }
                        else
                        {
                            xLogProperties.DocumentLibraryName = string.Empty;
                            xLogProperties.DroppedURLType = string.Empty;

                        }

                    }
                    if (isFolderExisted == false)
                    {
                       
                       
                       
                        IsConfigureCompleted = true;
                        xLogProperties.DisplayFolderName = txtDisplayName.Text;
                        if (rbtnUseDomainCredentials.Checked)
                        {
                            //domain credentials
                            xLogProperties.FolderAuthenticationType = AuthenticationType.Domain;

                        }
                        else
                        {
                            //for manual credentials

                            xLogProperties.FolderAuthenticationType = AuthenticationType.Manual;

                        }
                        xLogProperties.UserName = txtUserName.Text.Trim();
                        xLogProperties.Password = txtPassword.Text.Trim();
                        xLogProperties.SPSiteVersion = SPVersionClass.GetSPVersionFromUrl(URL, xLogProperties.UserName, xLogProperties.Password, xLogProperties.FolderAuthenticationType);

                        if (ThisAddIn.IsUrlIsTyped == true)
                        {
                            ThisAddIn tad = new ThisAddIn();
                            tad.NewConnection(xLogProperties);
                            ThisAddIn.IsUrlIsTyped = false;
                        }



                        this.Hide();
                    }
                    else
                    {
                        MessageBox.Show("Display name is already existed. Please provide another name", "ITOPIA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                EncodingAndDecoding.ShowMessageBox("Button OK", ex.Message, MessageBoxIcon.Error);

            }
        }
       
        /// <summary>
        /// <c>ShowEditForm</c> member function
        /// Method to open the selected folder details as Edit mode
        /// </summary>
        /// <param name="folderName"></param>
        public void ShowEditForm(string folderName)
        {
            btnOK.Text = "Save";
            txtDisplayName.Enabled = false;
            txtURL.Enabled = true;
            XmlNode xFolderDetails = UserLogManagerUtility.GetSPSiteURLDetails("", folderName);
            if (xFolderDetails != null)
            {
                //Save the details in log proeprties object

                xLogProperties.UserName = EncodingAndDecoding.Base64Decode(xFolderDetails.ChildNodes[0].InnerText);
                xLogProperties.Password = EncodingAndDecoding.Base64Decode(xFolderDetails.ChildNodes[1].InnerText);
                xLogProperties.DisplayFolderName = folderName;
                xLogProperties.SiteURL = xFolderDetails.ChildNodes[4].InnerText;
                xLogProperties.UsersStatus = UserStatus.Active;
                xLogProperties.DocumentLibraryName = xFolderDetails.ChildNodes[3].InnerText; ;
                xLogProperties.DroppedURLType = "";
                URL = xLogProperties.SiteURL;
                if (xFolderDetails.ChildNodes[5].InnerText == "Manually Specified")
                {
                    rbtnManuallySpecified.Checked = true;
                    txtUserName.Text = xLogProperties.UserName;
                    txtPassword.Text = xLogProperties.Password;
                    xLogProperties.FolderAuthenticationType = AuthenticationType.Manual;
                }
                else
                {
                    xLogProperties.FolderAuthenticationType = AuthenticationType.Domain;
                    rbtnUseDomainCredentials.Checked = true;
                }

                txtDisplayName.Text = xLogProperties.DisplayFolderName;
                txtURL.Text = xLogProperties.SiteURL;

                txtURL.ReadOnly = false;
                this.Focus();
                this.ShowDialog();
            }
            else
            {
                EncodingAndDecoding.ShowMessageBox("Show Form Event", "Unable to find the folder details in configration file.", MessageBoxIcon.Error);

            }
        }


        /// <summary>
        /// <c>IsPassedCredentialHasAccess</c> member function
        /// validates the input values for new connection
        /// </summary>
        /// <returns></returns>
        public bool IsPassedCredentialHasAccess()
        {
            bool bReturnValue = false;
            if (string.IsNullOrEmpty(txtDisplayName.Text))
            {
                MessageBox.Show("Please provide display name", "ITOPIA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            else
            {
                try
                {
                    int testname = Int32.Parse(txtDisplayName.Text);
                    MessageBox.Show("Only Digits are not accepted in display name.please provide different name ", "ITOPIA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                catch
                {


                }
            }
            if (string.IsNullOrEmpty(txtURL.Text))
            {
                MessageBox.Show("Please provide sharepoint URL", "ITOPIA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (rbtnManuallySpecified.Checked == true)
            {
                if (string.IsNullOrEmpty(txtUserName.Text) || string.IsNullOrEmpty(txtPassword.Text))
                {
                    MessageBox.Show("Please provide credentials.", "ITOPIA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }
            //if (rbtnUseDomainCredentials.Checked == true)
            //{
            //    System.Net.ICredentials credentionals = System.Net.CredentialCache.DefaultCredentials;
            //    authenticateService.UseDefaultCredentials = true;

            //    return true;
            //}
            //else
            //{
            //    System.Net.NetworkCredential credentionals = new System.Net.NetworkCredential(txtUserName.Text, txtUserName.Text);
            //    authenticateService.Credentials = credentionals;

            //}
            //try
            //{
            //    AuthenticationService.LoginResult result = authenticateService.Login(txtUserName.Text, txtUserName.Text);
            //    if (result.CookieName == "")
            //    {

            //    }
            //}
            //catch (Exception ex)
            //{

            //}
            //Check the manually specified userdetails with SP site 
            return true;

        }

        #endregion

        #region Events

        /// <summary>
        /// <c>btnOK_Click</c> Event handler
        /// calls <c>UpdateFolderConfigrationDetails</c> member function to update 
        /// configuration properties for new connection or modfy existing connection properties
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOK_Click(object sender, EventArgs e)
        {

            string pattern = @"^(http|https|ftp)\://[a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,3}(:[a-zA-Z0-9]*)?/?([a-zA-Z0-9\-\._\?\,\'/\\\+&amp;%\$#\=~])*[^\.\,\)\(\s]$";

            Regex reg = new Regex(pattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);
            string strurl = txtURL.Text.Trim().ToLower();
            if (strurl.StartsWith("http://") || strurl.StartsWith("https://") || strurl.StartsWith("ftp://"))
            {
                if (btnOK.Text.ToUpper() == "OK")
                {
                    UpdateFolderConfigrationDetails(true);

                }
                else
                {
                    UpdateFolderConfigrationDetails(false);
                }

            }
            else
            {

                EncodingAndDecoding.ShowMessageBox("Invalid url", "Please Enter a valid url", MessageBoxIcon.Error);

            }
        }

        /// <summary>
        /// <c>btnCancel_Click</c> Event Handler
        /// cancel the request and close the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            IsConfigureCompleted = false;
            this.Hide();
        }

        /// <summary>
        /// <c>rbtnUseDomainCredentials_CheckedChanged</c> Event handler
        /// hide and display fields based on Authentication type selection
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbtnUseDomainCredentials_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnUseDomainCredentials.Checked == true)
            {
                txtUserName.Text = string.Empty;
                txtPassword.Text = string.Empty;
                txtUserName.Enabled = false;
                txtPassword.Enabled = false;
            }
            else
            {

                txtUserName.Enabled = true;
                txtPassword.Enabled = true;
            }
        }


        /// <summary>
        /// <c>frmSPSiteConfiguration_Load</c> event handler
        /// display the  url in url field if dragged or set focus on url field if the user wants to type url
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmSPSiteConfiguration_Load(object sender, EventArgs e)
        {
            try
            {

                if (ThisAddIn.IsUrlIsTyped == true)
                {
                    txtDisplayName.Focus();
                }
                if (URL != null && URL != "")
                {
                    URL = HttpUtility.UrlDecode(URL);
                    xLogProperties.SiteURL = URL;
                    txtURL.Text = URL;
                    if (URL.Contains("/Forms/") == true)
                    {
                        string sTemp = URL.Substring(0, URL.IndexOf("/Forms/"));
                        xLogProperties.DocumentLibraryName = sTemp.Substring(sTemp.LastIndexOf("/") + 1);
                        xLogProperties.DroppedURLType = "/" + xLogProperties.DocumentLibraryName + "/Forms/";
                    }
                    else
                    {
                        xLogProperties.DocumentLibraryName = string.Empty;
                        xLogProperties.DroppedURLType = string.Empty;
                    }
                }
                //string version = SPVersionClass.GetSharePointVersionFromUrl(URL);
                //if (version == SPVersionClass.SiteVersion.SP2007.ToString())
                //{
                //    lblversion.Text = "Detected url is Sharepoint 2007 compatible";
                //}
                //else
                //{
                //    if (version == SPVersionClass.SiteVersion.SP2010.ToString())
                //    {
                //        lblversion.Text = "Detected url is Sharepoint 2010 compatible";
                //    }
                //    else
                //    {
                //        lblversion.Text = "Detected url is not a Sharepoint site url";
                //    }
                //}
                lblversion.Text = "";

            }
            catch (Exception ex)
            {
                EncodingAndDecoding.ShowMessageBox("Form Load Event", ex.Message, MessageBoxIcon.Error);
            }

        }

        #endregion

        /// <summary>
        /// <c>frmSPSiteConfiguration_KeyDown</c> Event handler
        /// calls <c>UpdateFolderConfigrationDetails</c>  method to update changes on pressing enter
        /// or close the form when escape is pressed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmSPSiteConfiguration_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Enter:
                    {
                        UpdateFolderConfigrationDetails(true);
                    }
                    break;
                case Keys.Escape:
                    {

                        IsConfigureCompleted = false;
                        this.Hide();
                    }
                    break;
            }
        }





    }
}
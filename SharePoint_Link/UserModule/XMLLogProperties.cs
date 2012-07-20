using SharePoint_Link.Utility;

namespace SharePoint_Link.UserModule
{
    # region Enum

    /// <summary>
    /// <c>UserStatus</c> Enum
    /// Enum holds possible values of user status Active/Removed
    /// </summary>
    public enum UserStatus
    {
        Active,
        Removed

    }
    # endregion

    /// <summary>
    /// <c>XMLLogProperties</c> class
    /// Implements the properties related to Sharepoint mapped connections  to be saved in Configuration Field
    /// </summary>
    public class XMLLogProperties
    {
        # region Private Variables

        /// <summary>
        /// <c>_username</c> private member field of type <c>string</c>
        /// Holds username  of connection  required to upload files and display Sharepoint library inside outlook
        /// </summary>
        private string _username = string.Empty;

        /// <summary>
        /// <c>_password</c>  private member field of type <c>string</c>
        ///  Holds password  of connection  required to upload files and display Sharepoint library inside outlook
        /// </summary>
        private string _password = string.Empty;

        /// <summary>
        /// <c>_foldername</c> private member field of type <c>string</c>
        ///   Holds outlook folder name  mapped to Sharepoint library inside outlook
        /// </summary>
        private string _foldername = string.Empty;

        /// <summary>
        /// <c>_siteurl</c> member field of type  <c>string</c>
        /// _siteurl holds the url of  sharepoint site mapped to document Library
        /// </summary>
        private string _siteurl = string.Empty;

        /// <summary>
        /// <c>_authenticationtype</c> member field of type Enum <c>AuthenticationType</c>
        /// _authenticationtype holds the authentication type User Credentials or Manual Configuration
        /// </summary>
        private AuthenticationType _authenticationtype = AuthenticationType.Manual;

        /// <summary>
        /// <c>_documentLibraryName</c> member field of type <c>string</c>
        /// _documentLibraryName member field holds the name of  Document Library
        /// </summary>
        private string _documentLibraryName = string.Empty;

        /// <summary>
        /// <c>status</c> member field of type enum <c>UserStatus</c>
        /// status holds the status of user Active or Not
        /// </summary>
        private UserStatus status = UserStatus.Active;

        /// <summary>
        /// <c>outlookFolderLocation</c> member field of type string
        /// Holds the location of sharepoint mappped folder in outlook
        /// </summary>
        private string outlookFolderLocation = string.Empty;

        /// <summary>
        /// <c>droppedURLType</c> member field of type <c>string</c>
        /// droppedURLType holds the dropped url type. document library or not
        /// </summary>
        private string droppedURLType = string.Empty;

        /// <summary>
        /// <c>sPSiteVersion</c> member field of type <c>string</c>
        /// holds the Information about sharepoint version whether it is sharepoint 2007 or sharepoint 2010
        /// </summary>
        private string sPSiteVersion = string.Empty;

        #endregion

        # region Properties

        /// <summary>
        /// <c>UserName</c> member property
        /// Encapsulates  _username
        /// </summary>
        public string UserName
        {
            get
            {
                return _username;
            }
            set
            {
                _username = value;
            }
        }

        /// <summary>
        /// <c>Password</c> member property
        /// Encapsulates  _password
        /// </summary>
        public string Password
        {
            get
            {
                return _password;
            }
            set
            {
                _password = value;
            }
        }

        /// <summary>
        /// <c>DisplayFolderName</c> member property
        /// Encapsulates  _foldername
        /// </summary>
        public string DisplayFolderName
        {
            get
            {
                return _foldername;
            }
            set
            {

                _foldername = value;
            }
        }

        /// <summary>
        /// <c>SiteURL</c> member property
        /// encapsulates  _siteurl member field
        /// </summary>
        public string SiteURL
        {
            get
            {
                return _siteurl;
            }
            set
            {
                _siteurl = value;
            }
        }

        /// <summary>
        /// <c>FolderAuthenticationType</c> member property
        /// encapsulates  _authenticationtype member field
        /// </summary>
        public AuthenticationType FolderAuthenticationType
        {
            get
            {
                return _authenticationtype;
            }
            set
            {
                _authenticationtype = value;
            }
        }

        /// <summary>
        /// <c>DocumentLibraryName</c> member property
        /// encapsulates  _documentLibraryName member field
        /// </summary>
        public string DocumentLibraryName
        {
            get
            {
                return _documentLibraryName;
            }
            set
            {
                _documentLibraryName = value;
            }
        }

        /// <summary>
        /// <c>UsersStatus</c> member property
        /// encapsulates  status member field
        /// </summary>
        public UserStatus UsersStatus
        {
            get { return status; }
            set { status = value; }
        }

        /// <summary>
        /// <c>OutlookFolderLocation</c> member property
        /// encapsulates  outlookFolderLocation member field
        /// </summary>
        public string OutlookFolderLocation
        {
            get
            {
                return outlookFolderLocation;
            }
            set
            {
                outlookFolderLocation = value;
            }
        }

        /// <summary>
        /// <c>DroppedURLType</c> member property
        /// encapsulates  droppedURLType member field
        /// </summary>
        public string DroppedURLType
        {
            get
            {
                return droppedURLType;
            }
            set
            {
                droppedURLType = value;
            }
        }

        /// <summary>
        /// <c>SPSiteVersion</c> member property
        /// encapsulates  sPSiteVersion member field
        /// </summary>
        public string SPSiteVersion
        {
            get { return sPSiteVersion; }
            set { sPSiteVersion = value; }
        }

        #endregion

    }

}

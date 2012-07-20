using System.Xml;

namespace SharePoint_Link.Utility
{
    /// <summary>
    /// <c>CommonProperties</c> Implements the fields and properties 
    /// Related to New  Connection to sharepoint document library
    /// </summary>
    class CommonProperties
    {

        /// <summary>
        /// <c>m_UploadDocLibraryName</c> member field of type string
        /// Holds document library name to be uploaded
        /// </summary>
        private string m_UploadDocLibraryName = string.Empty;


        /// <summary>
        /// <c>UploadDocLibraryName</c> member property
        /// Encapsulates  m_UploadDocLibraryName
        /// </summary>
        public string UploadDocLibraryName
        {
            get { return m_UploadDocLibraryName; }
            set { m_UploadDocLibraryName = value; }
        }

        /// <summary>
        /// <c>m_sharepointLibraryURL</c> member field of type string
        /// Holds the Url of the sharepoint mapped document library 
        /// </summary>
        private string m_sharepointLibraryURL = string.Empty;

        /// <summary>
        /// <c>SharepointLibraryURL</c> property field
        /// Encapsulates  m_sharepointLibraryURL
        /// </summary>
        public string SharepointLibraryURL
        {
            get { return m_sharepointLibraryURL; }
            set { m_sharepointLibraryURL = value; }
        }

        /// <summary>
        /// <c>m_userName</c> member field of type string
        /// holds user name of sharepoint site
        /// </summary>
        private string m_userName = string.Empty;

        /// <summary>
        /// <c>UserName</c> member property 
        /// encapsulates m_userName member field
        /// </summary>
        public string UserName
        {
            get { return m_userName; }
            set { m_userName = value; }
        }

        /// <summary>
        /// <c>m_password</c> member field of type string
        /// holds the password of sharepoint site
        /// </summary>
        private string m_password = string.Empty;


        /// <summary>
        /// <c>Password</c> property field
        /// encapsulates  m_password member field
        /// </summary>
        public string Password
        {
            get { return m_password; }
            set { m_password = value; }
        }

        /// <summary>
        /// <c>m_uploadingFileName</c> member field
        /// holds the name of the file  being uploaded to sharepoint document library
        /// </summary>
        private string m_uploadingFileName = string.Empty;

        /// <summary>
        /// <c>UploadingFileName</c> property field
        /// encapsulates  m_uploadingFileName
        /// </summary>
        public string UploadingFileName
        {
            get { return m_uploadingFileName; }
            set { m_uploadingFileName = value; }
        }

        /// <summary>
        /// <c>m_AuthenticationType</c> member field of type string
        /// holds the authentication type for sharepoint  document library
        /// </summary>
        private string m_AuthenticationType = string.Empty;


        /// <summary>
        /// <c>AuthenticationType</c> property field
        /// encapsulates  m_AuthenticationType
        /// </summary>
        public string AuthenticationType
        {
            get { return m_AuthenticationType; }
            set { m_AuthenticationType = value; }
        }

        /// <summary>
        /// <c>m_LibSite</c> member field of type string
        /// holds the url of sharepoint site url
        /// </summary>
        private string m_LibSite = string.Empty;


        /// <summary>
        /// <c>LibSite</c>  member property
        /// encapsulates m_LibSite
        /// </summary>
        public string LibSite
        {
            get { return m_LibSite; }
            set { m_LibSite = value; }
        }

        /// <summary>
        /// <c>credentionals</c> interface member of  ICredentials
        /// required to provide credentials 
        /// </summary>
        private System.Net.ICredentials credentionals;


        /// <summary>
        /// <c>Credentionals</c> member property 
        /// encapsulates  credentionals member field
        /// </summary>
        public System.Net.ICredentials Credentionals
        {
            get { return credentionals; }
            set { credentionals = value; }
        }


        /// <summary>
        /// <c>copyServiceURL</c> member field of type string
        /// holds the url of  sharepoint service to copy file to sharepoint document library
        /// </summary>
        private string copyServiceURL;


        /// <summary>
        /// <c>CopyServiceURL</c> member property
        /// Encapsulates  copyServiceURL member field
        /// </summary>
        public string CopyServiceURL
        {
            get { return copyServiceURL; }
            set { copyServiceURL = value; }
        }


        /// <summary>
        /// <c>m_uploadFolderNode</c> member field of type <c>XmlNode</c>
        /// holds all the node attributes of document library from configuration file
        /// </summary>
        private XmlNode m_uploadFolderNode = null;


        /// <summary>
        /// <c>UploadFolderNode</c> property field
        /// encapsulates  m_uploadFolderNode
        /// </summary>
        public XmlNode UploadFolderNode
        {
            get { return m_uploadFolderNode; }
            set { m_uploadFolderNode = value; }
        }

        /// <summary>
        /// <c>isUplodingCompleted</c> member field of type bool
        /// holds the true/false value to determin whether uploading is completed or not
        /// </summary>
        private bool isUplodingCompleted = false;


        /// <summary>
        /// <c>IsUplodingCompleted</c> property field
        /// encapsulates isUplodingCompleted
        /// </summary>
        public bool IsUplodingCompleted
        {
            get { return isUplodingCompleted; }
            set { isUplodingCompleted = value; }
        }

        /// <summary>
        /// <c>fileBytes</c> member field of type byte array
        /// holds file data in binary format
        /// </summary>
        private byte[] fileBytes = new byte[0];


        /// <summary>
        /// <c>FileBytes</c> member property 
        /// encapsulates fileBytes member field
        /// </summary>
        public byte[] FileBytes
        {
            get { return fileBytes; }
            set { fileBytes = value; }
        }


        /// <summary>
        /// <c>completedoclibraryURL</c> member field of type string
        /// holds the document library  absolute url
        /// </summary>
        private string completedoclibraryURL;


        /// <summary>
        /// <c>CompletedoclibraryURL</c> member property 
        /// encapsulates  completedoclibraryURL
        /// </summary>
        public string CompletedoclibraryURL
        {
            get { return completedoclibraryURL; }
            set { completedoclibraryURL = value; }
        }

        /// <summary>
        /// <c>libraryActualNameINSP</c> member field of type string
        /// holds the library name of sharepoint document library  
        /// </summary>
        private string libraryActualNameINSP;


        /// <summary>
        /// <c>LibraryActualNameINSP</c> member property 
        /// encapsulates  libraryActualNameINSP
        /// </summary>
        public string LibraryActualNameINSP
        {
            get { return libraryActualNameINSP; }
            set { libraryActualNameINSP = value; }
        }
    }
}

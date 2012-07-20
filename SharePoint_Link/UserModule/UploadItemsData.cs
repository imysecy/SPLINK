using Outlook = Microsoft.Office.Interop.Outlook;
using System;

namespace SharePoint_Link.UserModule
{
    /// <summary>
    /// <c>UploadItemsData</c> class
    /// This class implements the properties related to mail item or attachment to be uploaded 
    /// </summary>
    public class UploadItemsData
    {
        /// <summary>
        ///<c>uploadFileName</c> member field of type <c>string</c>
        /// uploadFileName holds the name of file being uploaded
        /// </summary>
        private string uploadFileName = string.Empty;

        
        /// <summary>
        /// <c>displayName</c> member field
        /// displayName holds the Display name of Mapped Folder
        /// </summary>
        private string displayName = string.Empty;

        /// <summary>
        /// <c>uploadType</c> member field of type <c>TypeOfUploading</c>
        /// uploadType holds the type of file being uploaded mailitem or Attachment
        /// </summary>
        private TypeOfUploading uploadType = TypeOfUploading.Mail;

        /// <summary>
        /// <c>typeofmailitem</c> member field of type TypeOfMailItem
        /// typeofmailitem holds the type of mail item. appointment, mailItem etc  
        /// </summary>
        private TypeOfMailItem typeofmailitem = TypeOfMailItem.Mail;

        /// <summary>
        /// <c>attachmentData</c> member field (byte array) 
        /// attachmentData holds the uploaded data in binary format
        /// </summary>
        private byte[] attachmentData = new byte[0];

        /// <summary>
        /// <c>uploadingMailItem</c> member field of type  <c>MailItem</c>
        /// holds the Mailitem currently being uploaded
        /// </summary>
        private Outlook.MailItem uploadingMailItem = null;

        /// <summary>
        /// <c>uploadingReportItem</c> member field of type <c>ReportItem</c>
        /// uploadingReportItem holds ReportItem being uploaded(if the file is of ReportItem type)
        /// </summary>
        private Outlook.ReportItem uploadingReportItem = null;

        /// <summary>
        /// <c>mailSubject</c> member field of type <c>mailSubject</c>
        /// mailSubject holds the mail subject of mailitem being uploaded
        /// </summary>
        private string mailSubject = string.Empty;

        /// <summary>
        /// <c>MailSubject</c> member property 
        /// MailSubject encapsulate  "mailSubject" member field
        /// </summary>
        public string MailSubject
        {
            get { return mailSubject; }
            set { mailSubject = value; }
        }

        /// <summary>
        /// <c>modifiedDate</c> member field of type <c>string</c>
        /// holds the mailitem modified/send date
        /// </summary>
        private string modifiedDate = string.Empty;


        private DateTime _eLapsedTime = DateTime.Now;

        /// <summary>
        /// <c>ModifiedDate</c> membe property
        /// Encapsulat "modifiedDate" member field 
        /// </summary>
        public string ModifiedDate
        {
            get { return modifiedDate; }
            set { modifiedDate = value; }
        }

        /// <summary>
        /// <c>uploadItemExtension</c> member field of type <c>string</c>
        /// holds the extension of  item being uploaded
        /// </summary>
        private string uploadItemExtension = ".msg";

        /// <summary>
        /// <c>UploadFileName</c> member property
        /// Get/Set property for uploaditem name(Encapsulates "uploadFileName" member field )
        /// </summary>
        public string UploadFileName
        {
            get
            {
                return uploadFileName;
            }
            set
            {
                uploadFileName = value;
            }
        }

        /// <summary>
        /// <c>UploadFileExtension</c> member property
        /// Get/Set property for uploaditem file/message extesnsion.
        /// Encapsulates 
        /// </summary>
        public string UploadFileExtension
        {
            get
            {
                return uploadItemExtension;
            }
            set
            {
                uploadItemExtension = value;
            }
        }

        /// <summary>
        /// <c>DisplayFolderName</c> member property
        /// Encapsulates "displayName" member field
        /// Get/Set property folder name, where the items associated with the folder.
        /// </summary>
        public string DisplayFolderName
        {
            get
            {
                return displayName;
            }
            set
            {
                displayName = value;
            }
        }

        /// <summary>
        /// <c>UploadType</c> member field of type <c>TypeOfUploading</c>
        /// Get/set property for check the uplaod item type
        /// </summary>
        public TypeOfUploading UploadType
        {
            get
            {
                return uploadType;
            }
            set
            {

                uploadType = value;
            }
        }

        /// <summary>
        /// <c>AttachmentData</c>
        /// Get/Set property for attachments data as byte array.
        /// Encapsulates  "attachmentData" member field
        /// </summary>
        public byte[] AttachmentData
        {
            get
            { return attachmentData; }
            set
            {
                attachmentData = value;
            }
        }

        /// <summary>
        /// <c>UploadingMailItem</c>
        /// get/set property for assign the uplaod mail item
        /// Encapsulates "uploadingMailItem" member field
        /// </summary>
        public Outlook.MailItem UploadingMailItem
        {
            get
            { return uploadingMailItem; }
            set
            {
                uploadingMailItem = value;
            }
        }

        /// <summary>
        /// <c>UploadingReportItem</c>
        /// member property and Encapsulates "uploadingReportItem"
        /// </summary>
        public Outlook.ReportItem UploadingReportItem
        {
            get
            { return uploadingReportItem; }
            set
            {
                uploadingReportItem = value;
            }
        }

        /// <summary>
        /// <c>TypeOfMailItem</c> member property
        /// Encapsulates "typeofmailitem" member field
        /// </summary>
        public TypeOfMailItem TypeOfMailItem
        {
            get
            {
                return typeofmailitem;
            }
            set
            {

                typeofmailitem = value;
            }
        }

        
        public DateTime ElapsedTime
        {
            get { return _eLapsedTime; }
            set { _eLapsedTime=value;}
        }
    }

    /// <summary>
    /// <c>TypeOfUploading</c> Enum
    /// Holds Possible values of uploading type MailItem or Attachment
    /// </summary>
    public enum TypeOfUploading
    {
        Mail,
        Attachment,

    }



    /// <summary>
    /// <c>TypeOfMailItem</c> Enum
    /// Holds Implemented MailItem Types e.g mail or ReportingItem
    /// </summary>
    public enum TypeOfMailItem
    {
        Mail,
        ReportItem,
    }
}

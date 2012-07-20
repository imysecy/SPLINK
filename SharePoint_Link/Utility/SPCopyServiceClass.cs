using System;
using SharePoint_Link.UserModule;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using SharePoint_Link.SPCopyService;
namespace SharePoint_Link.Utility
{
    /// <summary>
    /// <c>SPCopyServiceClass</c> class
    /// implements the functionalities to upload files to sharepoint document library
    /// </summary>
    class SPCopyServiceClass
    {

        /// <summary>
        /// <c>copyws</c> an object of class SPCopyService.Copy
        /// </summary>
        SPCopyService.Copy copyws;

        /// <summary>
        /// <c>SPCopyServiceClass</c> default constructor
        /// </summary>
        public SPCopyServiceClass()
        {
            copyws = new SPCopyService.Copy();
            copyws.CopyIntoItemsCompleted += new CopyIntoItemsCompletedEventHandler(copyws_CopyIntoItemsCompleted);
        }

        /// <summary>
        /// <c>copyws_CopyIntoItemsCompleted</c>  Event Handler 
        /// not currently used
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void copyws_CopyIntoItemsCompleted(object sender, SharePoint_Link.SPCopyService.CopyIntoItemsCompletedEventArgs e)
        {
          
        }

        /// <summary>
        /// <c>UploadItemUsingCopyService</c> member function
        /// this uploads the file to sharepoint document library
        /// </summary>
        /// <param name="uploadData"></param>
        /// <param name="property"></param>
        /// <returns></returns>
        public static bool UploadItemUsingCopyService(UploadItemsData uploadData, CommonProperties property)
        {
            bool result = false;
            try
            {

                SPCopyService.Copy copyws = new Copy();


                byte[] fileBytes = null;
                string[] destinationUrls = null;

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
                    Outlook.MailItem omail = uploadData.UploadingMailItem;
                    omail.SaveAs(tempFilePath, Outlook.OlSaveAsType.olMSG);

                    //// load the file into a file stream

                    //Read data to byte
                    fileBytes = File.ReadAllBytes(tempFilePath);
                    uploadData.UploadFileName = uploadData.UploadFileName.Trim() + uploadData.UploadFileExtension;
                }
                else
                {
                    //Set fullname to filename
                    fileBytes = uploadData.AttachmentData;
                    uploadData.UploadFileName = uploadData.UploadFileName.Trim() + uploadData.UploadFileExtension;
                }


                // format the destination URL
                destinationUrls = new string[] { property.LibSite + property.UploadDocLibraryName + "/" + uploadData.UploadFileName };

                // to specify the content type
                FieldInformation ctInformation = new FieldInformation();
                ctInformation.DisplayName = "Content Type";
                ctInformation.InternalName = "ContentType";
                ctInformation.Type = FieldType.Choice;
                ctInformation.Value = "Your content type";



                //FieldInformation[] metadata = { titleInformation };
                FieldInformation[] metadata = { };

                // execute the CopyIntoItems method
                copyws.CopyIntoItemsAsync("OutLook", destinationUrls, metadata, fileBytes, uploadData);




            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }


    }
}

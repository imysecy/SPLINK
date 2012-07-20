using System;
using SharePoint_Link.UserModule;
namespace SharePoint_Link.Utility
{
    /// <summary>
    /// <c>SPCopyClass</c> class
    /// 
    /// </summary>
    class SPCopyClass
    {
        /// <summary>
        /// <c>copyws</c> an object of <c> SPCopyService.Copy</c>
        /// this object is used to upload files to sharepoint document library in sharepoint 2007
        /// </summary>
        SPCopyService.Copy copyws;

        /// <summary>
        /// <c>SPCopyClass</c> Default constructor
        /// </summary>
        public SPCopyClass()
        {
            copyws = new SPCopyService.Copy();
            copyws.CopyIntoItemsCompleted += new SPCopyService.CopyIntoItemsCompletedEventHandler(copyws_CopyIntoItemsCompleted);
        }




        /// <summary>
        /// <c>UploadItemUsingCopyService</c> member function
        /// it calls <c>SPCopyServiceClass.UploadItemUsingCopyService</c> function to uploads the file to mapped sharepoint 2007 document library
        /// </summary>
        /// <param name="uploadData"></param>
        /// <param name="property"></param>
        /// <returns></returns>
        public static bool UploadItemUsingCopyService(UploadItemsData uploadData, CommonProperties property)
        {
            bool result = false;
            try
            {

                // for 2007
                result = SPCopyServiceClass.UploadItemUsingCopyService(uploadData, property);

                //for 2010


            }
            catch (Exception ex)
            {
                result = false;
            }
            return result;
        }


        /// <summary>
        /// <c>copyws_CopyIntoItemsCompleted</c> Event Handler
        /// it is executed after uploading file to document library 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void copyws_CopyIntoItemsCompleted(object sender, SPCopyService.CopyIntoItemsCompletedEventArgs e)
        {
            try
            {
                UploadItemsData uploaditemdata = (UploadItemsData)e.UserState;
                SPCopyService.CopyResult[] result = e.Results;
                int resultlength = result.Length;
                string outMessage = null;
                if (result.Length > 0)
                {

                    if (result[0].ErrorMessage != null)
                    {

                        if (Convert.ToString(result[0].ErrorMessage) == "Object reference not set to an instance of an object.")
                        {
                            outMessage = "Cause :: No prermission to access the List.(OR) List or Library is deleted or moved";
                        }
                        else if (Convert.ToString(result[0].ErrorCode) == "InvalidUrl")
                        {
                            outMessage = "Upload Falied.Filename contains some special characters." + result[0].DestinationUrl;
                        }
                        else
                        {
                            outMessage = result[0].ErrorMessage;
                        }

                    }

                }
                SharePoint_Link.frmUploadItemsList fmuploaditemlist = new frmUploadItemsList();
                fmuploaditemlist.CopyIntoItemCompleted(uploaditemdata, resultlength, outMessage);

            }
            catch (Exception)
            {


            }
        }
    }
}

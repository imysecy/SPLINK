namespace SharePoint_Link.UserModule
{
    /// <summary>
    /// <c>XMLLogOptions</c> class
    /// Implements the properties for Email Auto Delete option after uploading file to 
    /// Sharepoint Document Library mapped to Outlook folder
    /// </summary>
    public class XMLLogOptions
    {
        # region Private Variables

        /// <summary>
        /// <c>_autoDeleteEmails</c> member field of type <c>bool</c>
        ///_autoDeleteEmails  holds  true\false value to check whether the Auto Delete Email Option is selected or not
        /// </summary>
        private bool _autoDeleteEmails = false;
        
        #endregion

        # region Properties

        /// <summary>
        /// <c>AutoDeleteEmails</c> member Property
        /// Encapsulates "_autoDeleteEmails" member field
        /// </summary>
        public bool AutoDeleteEmails
        {
            get
            {
                return _autoDeleteEmails;
            }
            set
            {
                _autoDeleteEmails = value;
            }
        }
        
        #endregion

    }

}

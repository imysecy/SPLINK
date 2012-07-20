using System;
using System.Xml;
using System.IO;
using SharePoint_Link.UserModule;
using Utility;

namespace SharePoint_Link.Utility
{
    /// <summary>
    /// <c>UserLogManagerUtility</c> class
    /// implements the functionalities to manage configuration file
    /// </summary>
    static class UserLogManagerUtility
    {
        #region Global Declarations

        /// <summary>
        /// <c>userXMLFileName</c> member field
        /// holds the name of configuration  file 
        /// </summary>
        private static string userXMLFileName;

        /// <summary>
        /// <c>XMLFilePath</c> member field of type string
        /// holds the path of the <c>UserCredentialsLog</c> file
        /// </summary>
        public static string XMLFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Itopia\UserCredentialsLog.xml";
        public static string DefaultXMLFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Itopia\DefaultConnections.xml";

        /// <summary>
        /// <c>XMLOptionsFilePath</c> member field of type string
        /// XMLOptionsFilePath holds path of <c>UserOptions</c> file 
        /// </summary>
        public static string XMLOptionsFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Itopia\UserOptions.xml";

        /// <summary>
        /// <c>RootDirectory</c>  member field of type string
        /// holds the  directory path  where configuration files are stored
        /// </summary>
        public static string RootDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Itopia";

        # endregion

        #region Methods

        /// <summary>
        /// <c>CreateXMLFileForStoringUserCredentials</c>
        /// Method to save user credentionals
        /// </summary>
        /// <param name="xLogproperties"></param>
        /// <returns></returns>
        public static bool CreateXMLFileForStoringUserCredentials(XMLLogProperties xLogproperties)
        {
            try
            {
                //, Folder name, site url, authentication mode and credentials.
                xLogproperties.UserName = EncodingAndDecoding.Base64Encode(xLogproperties.UserName);
                xLogproperties.Password = EncodingAndDecoding.Base64Encode(xLogproperties.Password);
                XmlDocument xmlDoc = new XmlDocument();
                XmlElement elemRoot = null, elem = null;
                XmlNode root = null;
                bool isNewXMlFile = true;

                UserLogManagerUtility.CheckItopiaDirectoryExits();

                //Check file is existed or not
                if (System.IO.File.Exists(UserLogManagerUtility.XMLFilePath) == true)
                {
                    //Save the details in Xml file
                    xmlDoc.Load(UserLogManagerUtility.XMLFilePath);
                    //Get the root Elemet
                    root = xmlDoc.DocumentElement;
                    elemRoot = xmlDoc.CreateElement("Folder");
                    isNewXMlFile = false;
                }
                else
                {
                    XmlDeclaration xmlDec = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", String.Empty);
                    xmlDoc.PrependChild(xmlDec);
                    XmlElement docRoot = xmlDoc.CreateElement("UserCredentialsLog");
                    xmlDoc.AppendChild(docRoot);

                    //Create root node
                    elemRoot = xmlDoc.CreateElement("Folder");
                    docRoot.AppendChild(elemRoot);
                }
                elem = xmlDoc.CreateElement("UserName");
                elem.InnerText = xLogproperties.UserName;
                elemRoot.AppendChild(elem);

                elem = xmlDoc.CreateElement("Password");
                elem.InnerText = xLogproperties.Password;
                elemRoot.AppendChild(elem);
                //string strFolderName,string strSiteURL,string strAuthenticationType)
                elem = xmlDoc.CreateElement("DisplayName");
                elem.InnerText = xLogproperties.DisplayFolderName;
                elemRoot.AppendChild(elem);

                elem = xmlDoc.CreateElement("DocLibName");
                elem.InnerText = xLogproperties.DocumentLibraryName;
                elemRoot.AppendChild(elem);

                elem = xmlDoc.CreateElement("URL");
                elem.InnerText = xLogproperties.SiteURL;
                elemRoot.AppendChild(elem);
                elem = xmlDoc.CreateElement("AuthenticationType");
                if (xLogproperties.FolderAuthenticationType == AuthenticationType.Domain)
                {
                    elem.InnerText = "Domain Credentials";
                }
                else
                {
                    elem.InnerText = "Manually Specified";
                }
                elemRoot.AppendChild(elem);

                elem = xmlDoc.CreateElement("Status");
                if (xLogproperties.UsersStatus == UserStatus.Active)
                {
                    elem.InnerText = "Active";
                }
                else
                {
                    elem.InnerText = "Removed";
                }
                elemRoot.AppendChild(elem);

                //this will used to check the folder already existed or not
                elem = xmlDoc.CreateElement("FolderNameToCompare");
                elem.InnerText = xLogproperties.DisplayFolderName.ToUpper();
                elemRoot.AppendChild(elem);

                elem = xmlDoc.CreateElement("OutlookLocation");
                xLogproperties.OutlookFolderLocation = xLogproperties.OutlookFolderLocation.Replace("\\\\", "");
                elem.InnerText = xLogproperties.OutlookFolderLocation;
                elemRoot.AppendChild(elem);

                elem = xmlDoc.CreateElement("DateAdded");
                elem.InnerText = System.DateTime.Now.ToString();
                elemRoot.AppendChild(elem);

                elem = xmlDoc.CreateElement("LastUpload");
                elem.InnerText = System.DateTime.Now.ToString();
                elemRoot.AppendChild(elem);

                elem = xmlDoc.CreateElement("URLType");
                elem.InnerText = xLogproperties.DroppedURLType;
                elemRoot.AppendChild(elem);

                elem = xmlDoc.CreateElement("SPSiteVersion");
                elem.InnerText = xLogproperties.SPSiteVersion;
                elemRoot.AppendChild(elem);

                if (isNewXMlFile == false)
                {
                    //XML file already existed add the node to xml file 
                    root.InsertBefore(elemRoot, root.FirstChild);
                }
                //Save xml file 
                xmlDoc.Save(UserLogManagerUtility.XMLFilePath);

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return false;
        }

        /// <summary>
        /// <c>CreateXMLFileForStoringUserOptions</c>
        /// creates xml file to store UserOptions either to select AutoDelete Emails after  uploading or not
        /// </summary>
        /// <param name="xLogOptions"></param>
        /// <returns></returns>
        public static bool CreateXMLFileForStoringUserOptions(XMLLogOptions xLogOptions)
        {
            try
            {
                //, Folder name, site url, authentication mode and credentials.
                XmlDocument xmlDoc = new XmlDocument();
                XmlElement elemRoot = null, elem = null;
                XmlNode root = null;
                bool isNewXMlFile = true;

                UserLogManagerUtility.CheckItopiaDirectoryExits();

                //Check file is existed or not
                if (System.IO.File.Exists(UserLogManagerUtility.XMLOptionsFilePath) == true)
                {
                    //Save the details in Xml file
                    xmlDoc.Load(UserLogManagerUtility.XMLOptionsFilePath);
                    xmlDoc.RemoveAll();
                }

                XmlDeclaration xmlDec = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", String.Empty);
                xmlDoc.PrependChild(xmlDec);
                XmlElement docRoot = xmlDoc.CreateElement("UserOptionsLog");
                xmlDoc.AppendChild(docRoot);

                //Create root node
                elemRoot = xmlDoc.CreateElement("Options");
                docRoot.AppendChild(elemRoot);

                elem = xmlDoc.CreateElement("AutoDeleteEmails");
                elem.InnerText = Convert.ToString(xLogOptions.AutoDeleteEmails);
                elemRoot.AppendChild(elem);

                if (isNewXMlFile == false)
                {
                    //XML file already existed add the node to xml file 
                    root.InsertBefore(elemRoot, root.FirstChild);
                }
                //Save xml file 
                xmlDoc.Save(UserLogManagerUtility.XMLOptionsFilePath);

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return false;
        }



        /// <summary>
        /// <c>UpdateUserFolderRemovedStatus</c> member function
        /// Method to update folder status in XML file
        /// </summary>
        /// <param name="folderName">Removing folder name as string.</param>
        /// <returns></returns>
        public static bool UpdateUserFolderRemovedStatus(string folderName)
        {
            try
            {
                XmlNode xNode = null;
                UserLogManagerUtility.CheckItopiaDirectoryExits();
                XmlDocument xDoc = new XmlDocument();
                if (File.Exists(UserLogManagerUtility.XMLFilePath))
                {
                    XmlNode root;
                    xDoc.Load(UserLogManagerUtility.XMLFilePath);
                    //Get the root Elemet
                    root = xDoc.DocumentElement;
                    //Get the Nodes contains with filename
                    XmlNodeList xlst = root.SelectNodes("descendant::Folder[FolderNameToCompare='" + folderName.ToUpper() + "']");
                    if (xlst.Count > 0)
                    {
                        xNode = xlst[0];
                        //Update status node
                        xNode.ChildNodes[6].InnerText = "Removed";
                        //Update time
                        xNode.ChildNodes[10].InnerText = DateTime.Now.ToString();
                        xDoc.Save(UserLogManagerUtility.XMLFilePath);
                        return true;
                    }

                }
            }
            catch (Exception ex)
            { }
            return false;
        }

        /// <summary>
        /// <c>GetSPSiteURLDetails</c> member function
        /// Method to get the folder details from XML file
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="folderName">folder name to get the details </param>
        /// <returns>Folder details as XmlNode </returns>
        public static XmlNode GetSPSiteURLDetails(string userName, string folderName)
        {
            XmlNode xNode = null;
            try
            {
                UserLogManagerUtility.CheckItopiaDirectoryExits();
                XmlDocument xDoc = new XmlDocument();
                if (File.Exists(UserLogManagerUtility.XMLFilePath))
                {
                    XmlNode root;
                    xDoc.Load(UserLogManagerUtility.XMLFilePath);
                    //Get the root Elemet
                    root = xDoc.DocumentElement;

                    //XPathNavigator navigator = xDoc.CreateNavigator();
                    //// Select all books authored by Melville.
                    //XPathNodeIterator nodes = navigator.Select("descendant::File[name='" + fileName + "']");

                    //Get the Nodes contains with filename
                    XmlNodeList xlst = root.SelectNodes("descendant::Folder[FolderNameToCompare='" + folderName.ToUpper() + "']");
                    if (xlst.Count > 0)
                    {
                        xNode = xlst[0];
                    }

                }
            }
            catch (Exception ex)
            { }
            return xNode;
        }

        /// <summary>
        /// <c>IsFolderExisted</c> member function
        /// Method check the given folder name is already existed or not
        /// </summary>
        /// <param name="folderName">Folder name to check</param>
        /// <returns>True/False as Boolean</returns>
        public static Boolean IsFolderExisted(string folderName)
        {

            try
            {
                UserLogManagerUtility.CheckItopiaDirectoryExits();
                XmlDocument xDoc = new XmlDocument();
                if (File.Exists(UserLogManagerUtility.XMLFilePath))
                {
                    XmlNode root;
                    xDoc.Load(UserLogManagerUtility.XMLFilePath);
                    //Get the root Elemet
                    root = xDoc.DocumentElement;

                    //Get the Nodes contains with filename
                    XmlNodeList xlst = root.SelectNodes("descendant::Folder[FolderNameToCompare='" + folderName.ToUpper() + "']");
                    if (xlst.Count > 0)
                    {
                        string s1 = xlst[0].ChildNodes[11].InnerText;
                        s1 = s1.Substring(s1.IndexOf("<![CDATA[") + 9, s1.Length - 9);
                        s1 = s1.Replace("<![CDATA[", "");
                        s1 = s1.Replace("]]>", "");

                        XmlNode xNode = xlst[0];
                        //If folder status is removed then user can create new folder
                        if (xNode.ChildNodes[6].InnerText == "Active")
                        {
                            return true;
                        }

                    }

                }
            }
            catch (Exception ex)
            { }
            return false;
        }


        /// <summary>
        /// <c>GetAllFoldersDetails</c> member function
        /// Method to get all the folders information
        /// </summary>
        /// <returns>XmlNodeList</returns>
        public static XmlNodeList GetAllFoldersDetails()
        {

            if (System.IO.File.Exists(XMLFilePath) == true)
            {

                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(XMLFilePath);
                //Get the root Elemet
                XmlNode root = xDoc.DocumentElement;
                return root.ChildNodes;

            }
            return null;
        }

        /// <summary>
        /// <c>GetUserConfigurationOptions</c> member function
        /// get Email Auto Delete option from configuration option 
        /// </summary>
        /// <returns></returns>
        public static XMLLogOptions GetUserConfigurationOptions()
        {
            XMLLogOptions myOption = new XMLLogOptions();
            try
            {
                if (System.IO.File.Exists(XMLOptionsFilePath) == true)
                {

                    XmlDocument xDoc = new XmlDocument();
                    xDoc.Load(XMLOptionsFilePath);

                    XmlNode root = xDoc.DocumentElement;
                    root = xDoc.DocumentElement;
                    myOption.AutoDeleteEmails = Convert.ToBoolean(root.ChildNodes[0].ChildNodes[0].ChildNodes[0].Value);
                }
                else
                {
                    myOption.AutoDeleteEmails = true;

                    CreateXMLFileForStoringUserOptions(myOption);
                }
            }
            catch (Exception ex) { }

            return myOption;
        }


        /// <summary>
        /// <c>GetAllFoldersDetails</c>
        /// Method to get the user fodlers details based on status
        /// </summary>
        /// <param name="userCurrentStatus">Status as UserStatus</param>
        /// <returns>XmlNodeList</returns>
        public static XmlNodeList GetAllFoldersDetails(UserStatus userCurrentStatus)
        {

            if (System.IO.File.Exists(XMLFilePath) == true)
            {
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(XMLFilePath);
                //Get the root Elemet
                XmlNode root = xDoc.DocumentElement;
                if (userCurrentStatus == UserStatus.Active)
                {
                    //Get the Nodes status is Active
                    return root.SelectNodes("descendant::Folder[Status='Active']");
                }
                else
                {
                    //Get the Nodes status is removed
                    return root.SelectNodes("descendant::Folder[Status='Removed']");
                }

            }
            return null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="userCurrentStatus"></param>
        /// <returns></returns>
        public static XmlNodeList GetDefaultFoldersDetails(UserStatus userCurrentStatus)
        {

            if (System.IO.File.Exists(DefaultXMLFilePath) == true)
            {
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(DefaultXMLFilePath);
                //Get the root Elemet
                XmlNode root = xDoc.DocumentElement;
                if (userCurrentStatus == UserStatus.Active)
                {
                    //Get the Nodes status is Active
                    return root.SelectNodes("descendant::Folder[Status='Active']");
                }
                else
                {
                    //Get the Nodes status is removed
                    return root.SelectNodes("descendant::Folder[Status='Removed']");
                }

            }
            return null;
        }

        /// <summary>
        /// <c>UpdateFolderConfigNodeDetails</c> member function
        /// Method to update xml node data in xml file
        /// </summary>
        /// <param name="updateNode">Update node as XmlNode</param>
        /// <param name="updateNodeName">Updating child node name as string. </param>
        /// <param name="updateNodeValue">Updating child node value as string.</param>
        /// <returns></returns>
        public static bool UpdateFolderConfigNodeDetails(string folderName, string updateNodeName, string updateNodeValue)
        {
            try
            {
                XmlNode xNode = null;
                UserLogManagerUtility.CheckItopiaDirectoryExits();
                XmlDocument xDoc = new XmlDocument();
                if (File.Exists(XMLFilePath))
                {
                    XmlNode root;
                    xDoc.Load(XMLFilePath);
                    //Get the root Elemet
                    root = xDoc.DocumentElement;
                    //Get the upload node index
                    int uploadNodeIndex = 0;
                    //Get the Nodes contains with filename
                    XmlNodeList xlst = root.SelectNodes("descendant::Folder[FolderNameToCompare='" + folderName.ToUpper() + "']");
                    if (xlst.Count > 0)
                    {
                        xNode = xlst[0];
                        switch (updateNodeName)
                        {
                            case "Status":
                                xNode.ChildNodes[6].InnerText = updateNodeValue;
                                //uploadNodeIndex = 6;
                                break;
                            case "DisplayName":
                                uploadNodeIndex = 2;
                                xNode.ChildNodes[2].InnerText = updateNodeValue;
                                xNode.ChildNodes[7].InnerText = updateNodeValue.ToUpper();
                                //FolderNameToCompare
                                break;
                            case "LastUpload":
                                xNode.ChildNodes[10].InnerText = updateNodeValue;
                                break;
                            case "OutlookLocation":
                                xNode.ChildNodes[8].InnerText = updateNodeValue;
                                break;
                            default:
                                break;
                        }
                        xDoc.Save(UserLogManagerUtility.XMLFilePath);
                        return true;
                    }






                }
            }
            catch (Exception ex)
            { }
            return false;
        }

        /// <summary>
        /// <c>UpdateFolderConfigDetails</c> member function
        /// Method to update xml node data in xml file
        /// </summary>
        /// <param name="updateNode">Update node as XmlNode</param>
        /// <param name="updateNodeName">Updating child node name as string. </param>
        /// <param name="updateNodeValue">Updating child node value as string.</param>
        /// <returns></returns>
        public static bool UpdateFolderConfigDetails(string oldFolderName, XMLLogProperties xLogProperties)
        {
            try
            {
                XmlNode xNode = null;
                UserLogManagerUtility.CheckItopiaDirectoryExits();
                XmlDocument xDoc = new XmlDocument();
                if (File.Exists(XMLFilePath))
                {
                    XmlNode root;
                    xDoc.Load(XMLFilePath);
                    //Get the root Elemet
                    root = xDoc.DocumentElement;
                    if (string.IsNullOrEmpty(oldFolderName))
                    {
                        oldFolderName = xLogProperties.DisplayFolderName;
                    }
                    //Get the Nodes contains with filename
                    XmlNodeList xlst = root.SelectNodes("descendant::Folder[FolderNameToCompare='" + oldFolderName.ToUpper() + "']");
                    if (xlst.Count > 0)
                    {
                        xNode = xlst[0];

                        if (xLogProperties.FolderAuthenticationType == AuthenticationType.Manual)
                        {
                            //, Folder name, site url, authentication mode and credentials.
                            xLogProperties.UserName = EncodingAndDecoding.Base64Encode(xLogProperties.UserName);
                            xLogProperties.Password = EncodingAndDecoding.Base64Encode(xLogProperties.Password);

                            xNode.ChildNodes[5].InnerText = "Manually Specified";

                        }
                        else
                        {
                            xNode.ChildNodes[5].InnerText = "Domain Credentials";

                            xLogProperties.UserName = string.Empty;
                            xLogProperties.Password = string.Empty;
                        }
                        xNode.ChildNodes[0].InnerText = xLogProperties.UserName;
                        xNode.ChildNodes[1].InnerText = xLogProperties.Password;
                        xNode.ChildNodes[2].InnerText = xLogProperties.DisplayFolderName;
                        xNode.ChildNodes[3].InnerText = xLogProperties.DocumentLibraryName;
                        xNode.ChildNodes[4].InnerText = xLogProperties.SiteURL;

                        //xNode.ChildNodes[5].InnerText = xLogProperties.FolderAuthenticationType;
                        //xNode.ChildNodes[6].InnerText = "Active"; //Active
                        //xNode.ChildNodes[7].InnerText = xLogProperties.DisplayFolderName.ToUpper();
                        //xNode.ChildNodes[8].InnerText = xLogProperties.OutlookFolderLocation;
                        //xNode.ChildNodes[9].InnerText = DateTime.Now;//Date Added
                        xNode.ChildNodes[10].InnerText = DateTime.Now.ToString();
                        xNode.ChildNodes[11].InnerText = xLogProperties.DroppedURLType;


                        xDoc.Save(UserLogManagerUtility.XMLFilePath);
                        return true;
                    }

                }
            }
            catch (Exception ex)
            { }
            return false;
        }



        /// <summary>
        /// <c>CheckItopiaDirectoryExits</c> member function
        /// Method to create root directroy
        /// </summary>
        public static void CheckItopiaDirectoryExits()
        {
            try
            {
                if (System.IO.Directory.Exists(RootDirectory) == false)
                {
                    System.IO.Directory.CreateDirectory(RootDirectory);
                }
            }
            catch { }
        }

        # endregion

        #region Properties

        /// <summary>
        /// <c>UserXMLFileName</c>
        /// Property to get the user xml file name except  _UserCredentialsLog.xml
        /// </summary>
        public static string UserXMLFileName
        {
            get { return userXMLFileName; }
            set
            {
                userXMLFileName = value;
                XMLFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Itopia\" + userXMLFileName + "_UserCredentialsLog.xml";
            }
        }

        #endregion


        /// <summary>
        /// <c>IsDocumentLibrary</c> member function
        /// check to verify either the url  belongs to a sharepoint document library or not based on folder
        /// </summary>
        /// <param name="FolderName"></param>
        /// <returns></returns>
        public static bool IsDocumentLibrary(string FolderName)
        {
            bool result = true;

            try
            {
                XmlNodeList xFolders = UserLogManagerUtility.GetAllFoldersDetails(UserStatus.Active);
                if (xFolders != null)
                {
                    string folderName = string.Empty, DocLibName = string.Empty;

                    foreach (XmlNode xNode in xFolders)
                    {
                        try
                        {

                            folderName = xNode.ChildNodes[2].InnerText;
                            //Get Doc Lib Name

                            DocLibName = xNode.ChildNodes[3].InnerText;

                            if (folderName.Trim() == folderName.Trim())
                            {
                                if (string.IsNullOrEmpty(DocLibName) == true)
                                {
                                    result = false;
                                }
                            }


                        }
                        catch { }
                    }

                }
            }
            catch (Exception)
            {


            }
            return result;
        }


        /// <summary>
        /// <c>GetSPSiteURL</c> member function
        /// get the Sharepoint site url from configuration file based on folder name
        /// </summary>
        /// <param name="FolderName"></param>
        /// <returns></returns>
        public static string GetSPSiteURL(string FolderName)
        {
            string result = "";
            try
            {
                XmlNodeList xFolders = UserLogManagerUtility.GetAllFoldersDetails(UserStatus.Active);
                if (xFolders != null)
                {
                    string folderName = string.Empty, siteurl = string.Empty;

                    foreach (XmlNode xNode in xFolders)
                    {
                        try
                        {

                            folderName = xNode.ChildNodes[2].InnerText;
                            //Get Doc Lib Name

                            siteurl = xNode.ChildNodes[4].InnerText;

                            if (folderName.Trim() == FolderName.Trim())
                            {
                                result = siteurl;

                            }


                        }
                        catch { }
                    }

                }
            }
            catch (Exception)
            {


            }
            return result;
        }

        /// <summary>
        /// <c>GetFolderOutLookLocation</c>  member function
        /// returns the location of outlook folder 
        /// </summary>
        /// <param name="FolderName"></param>
        /// <returns></returns>
        public static string GetFolderOutLookLocation(string FolderName)
        {
            string result = "";
            try
            {
                XmlNodeList xFolders = UserLogManagerUtility.GetAllFoldersDetails(UserStatus.Active);
                if (xFolders != null)
                {
                    string folderName = string.Empty, siteurl = string.Empty;

                    foreach (XmlNode xNode in xFolders)
                    {
                        try
                        {

                            folderName = xNode.ChildNodes[2].InnerText;
                            //Get Doc Lib Name

                            siteurl = xNode.ChildNodes[8].InnerText;

                            if (folderName.Trim() == FolderName.Trim())
                            {
                                result = siteurl;

                            }


                        }
                        catch { }
                    }

                }
            }
            catch (Exception)
            {


            }
            return result;
        }

        /// <summary>
        /// <c>GetRelativePath</c> member function
        /// returns the relative url of document library 
        /// </summary>
        /// <param name="FolderName"></param>
        /// <returns></returns>
        public static string GetRelativePath(string FolderName)
        {
            string result = "";
            try
            {
                XmlNodeList xFolders = UserLogManagerUtility.GetAllFoldersDetails(UserStatus.Active);
                if (xFolders != null)
                {
                    string folderName = string.Empty, siteurl = string.Empty;

                    foreach (XmlNode xNode in xFolders)
                    {
                        try
                        {

                            folderName = xNode.ChildNodes[2].InnerText;
                            //Get Doc Lib Name

                            siteurl = xNode.ChildNodes[11].InnerText;

                            if (folderName.Trim() == FolderName.Trim())
                            {
                                result = siteurl;

                            }


                        }
                        catch { }
                    }

                }
            }
            catch (Exception)
            {


            }
            return result;
        }


        internal static void AddDefaultMissingConnections()
        {
            try
            {
                XmlNodeList user_list = GetAllFoldersDetails(UserStatus.Active);
                XmlNodeList defa_list = GetDefaultFoldersDetails(UserStatus.Active);

                string userListfolderName = string.Empty;
                string userListurl = string.Empty;
                string defListfolderName = string.Empty;
                string defListurl = string.Empty;

                foreach (XmlNode xNode in defa_list)
                {
                    bool isFound = false;

                    defListfolderName = xNode.ChildNodes[2].InnerText;
                    defListurl = xNode.ChildNodes[4].InnerText;

                    foreach (XmlNode xNode1 in user_list)
                    {
                        userListfolderName = xNode1.ChildNodes[2].InnerText;
                        userListurl = xNode1.ChildNodes[4].InnerText;

                        if (defListfolderName.Equals(userListfolderName, StringComparison.OrdinalIgnoreCase)
                          && userListurl.Equals(defListurl, StringComparison.OrdinalIgnoreCase))
                        {
                            isFound = true;
                            break;
                        }

                    }

                    if (!isFound)
                    {
                        XMLLogProperties xLogProperties = new XMLLogProperties();

                        xLogProperties.DisplayFolderName = xNode.ChildNodes[2].InnerText;

                        xLogProperties.DocumentLibraryName = xNode.ChildNodes[3].InnerText;
                        xLogProperties.DroppedURLType = xNode.ChildNodes[11].InnerText;

                        if (xNode.ChildNodes[5].InnerText.Equals("Manually Specified", StringComparison.OrdinalIgnoreCase))
                            xLogProperties.FolderAuthenticationType = AuthenticationType.Manual;
                        else
                            xLogProperties.FolderAuthenticationType = AuthenticationType.Domain;

                        xLogProperties.OutlookFolderLocation = xNode.ChildNodes[8].InnerText;
                        xLogProperties.Password = xNode.ChildNodes[1].InnerText;
                        xLogProperties.SiteURL = xNode.ChildNodes[4].InnerText;
                        xLogProperties.SPSiteVersion = xNode.ChildNodes[12].InnerText;
                        xLogProperties.UserName = xNode.ChildNodes[0].InnerText;
                        xLogProperties.UsersStatus = UserStatus.Active;

                        CreateXMLFileForStoringUserCredentials(xLogProperties);
                    }
                }
            }
            catch (Exception ex)
            {
                ListWebClass.Log(ex.Message, true);
            }
            
        }
    }

    /// <summary>
    /// <c>AuthenticationType</c> member field of type enum
    /// holds authentication type Manual or Domain
    /// </summary>
    public enum AuthenticationType
    {
        Manual = 0,
        Domain = 1
    }



}

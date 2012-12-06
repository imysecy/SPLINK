using System;
using System.Text;
using System.Xml;
using System.IO;
using SharePoint_Link.UserModule;
using Microsoft.SharePoint.Client;

namespace SharePoint_Link.Utility
{
    /// <summary>
    /// <c>ListWebClass</c> class
    /// Implements functionalities  which are common across the addin application
    /// </summary>
    class ListWebClass
    {


        //  WebClient m_WC;

        /// <summary>
        ///<c>listService</c>  class object of class <c>Lists</c> 
        /// </summary>
        ListWebService.Lists listService;
        
      

        /// <summary>
        /// <c>fileID</c> member field of type string
        /// holds the uploaded file ID
        /// </summary>
        private static string fileID = string.Empty;

        /// <summary>
        /// <c>FileID</c> member property 
        /// Encapsulates  fileID member field
        /// </summary>
        public static string FileID
        {
            get { return ListWebClass.fileID; }
            set { ListWebClass.fileID = value; }
        }

        /// <summary>
        /// <c>ListWebClass</c> Default Constructor
        /// </summary>
        public ListWebClass()
        {
            if (SPVersionClass.SPSiteVersion == SPVersionClass.SiteVersion.SP2007.ToString())
            {
                listService = new ListWebService.Lists();
            }


        }



        /// <summary>
        /// <c>GetListItemsAsync</c> Event Handler
        /// assigns properties to <c>ListWebService.Lists</c> object.
        /// it register <c>GetListItemsCompleted</c>  Event of  <c>ListWebService.Lists</c> object
        /// and calls <c>GetListItemsAsync</c> function to upload files to sharepoint document library in sharepoint 2007
        /// </summary>
        /// <param name="m_uploadDocLibraryName"></param>
        /// <param name="viewname"></param>
        /// <param name="ndQuery"></param>
        /// <param name="ndViewFields"></param>
        /// <param name="rowlimit"></param>
        /// <param name="ndQueryOptions"></param>
        /// <param name="webid"></param>
        /// <param name="uploadData"></param>
        /// <param name="property"></param>
        public void GetListItemsAsync(string m_uploadDocLibraryName, string viewname, XmlNode ndQuery, XmlNode ndViewFields, string rowlimit, XmlNode ndQueryOptions, string webid, object uploadData, CommonProperties property)
        {
            try
            {

                if (SPVersionClass.SPSiteVersion == SPVersionClass.SiteVersion.SP2007.ToString())
                {


                    // for 2007//
                    ListWebService.Lists listService = new ListWebService.Lists();
                    listService.Credentials = property.Credentionals;
                    listService.Url = property.CopyServiceURL;
                    listService.GetListItemsCompleted += new ListWebService.GetListItemsCompletedEventHandler(listService_GetListItemsCompleted);
                    listService.GetListItemsAsync(m_uploadDocLibraryName, null, ndQuery, ndViewFields, "2", ndQueryOptions, null, uploadData);

                }
            }
            catch (Exception)
            {


            }

        }



        /// <summary>
        /// <c>listService_GetListItemsCompleted</c> Event Handler
        /// it gets the uploaded data node from sharepoint 2007 document library and calls 
        /// <c>listService_GetListItemsCompleted</c>  method to update status 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listService_GetListItemsCompleted(object sender, SharePoint_Link.ListWebService.GetListItemsCompletedEventArgs e)
        {
            try
            {
                if (e.Result != null)
                {
                    XmlNode ndVolunteerListItems = e.Result;
                    UserModule.UploadItemsData uploaditemdata = (UserModule.UploadItemsData)e.UserState;

                    SharePoint_Link.frmUploadItemsList fmuploaditemlist = new frmUploadItemsList();
                    fmuploaditemlist.ListItemCompleted(ndVolunteerListItems, uploaditemdata);


                }

            }
            catch (Exception ex)
            {


                string outMessage1 = ex.Message;
                if (ex.InnerException != null)
                {
                    outMessage1 = ex.InnerException.Message;
                }



                UserModule.UploadItemsData uitemdata = (UserModule.UploadItemsData)e.UserState;

                SharePoint_Link.frmUploadItemsList fmuploaditemlist = new frmUploadItemsList();
                fmuploaditemlist.UpdataGridRows(false, "exception in getlistitemscompleted event." + outMessage1, " ", uitemdata);

                fmuploaditemlist.Show();
                fmuploaditemlist.HideProgressPanel(true);


            }
        }

        /// <summary>
        /// <c>GetSPListName</c> member function
        /// returns the name of sharepoint List by calling  <c>GetSPListNameUsingService</c> member function for sharepoint 2007
        /// or <c>GetSPListNameUsingClientOM</c> member function for sharepoit 2010
        /// </summary>
        /// <param name="property"></param>
        /// <returns></returns>
        public Boolean GetSPListName(CommonProperties property)
        {
            bool result = false;
            try
            {
                if (SPVersionClass.SPSiteVersion == SPVersionClass.SiteVersion.SP2007.ToString())
                {
                    //using service for 2007
                    result = GetSPListNameUsingService(property);
                }
                else
                {
                    if (SPVersionClass.SPSiteVersion == SPVersionClass.SiteVersion.SP2010.ToString())
                    {
                        //using clien for 2010
                        result = GetSPListNameUsingClientOM(property);

                    }
                    else
                    {

                    }
                }




            }
            catch (Exception)
            {

                result = false;
            }
            return result;
        }

        /// <summary>
        /// <c>GetSPListNameUsingService</c> member function
        /// Retrieve the Document Library Nmae from sharepoint 2007
        /// </summary>
        /// <param name="property"></param>
        /// <returns></returns>
        public Boolean GetSPListNameUsingService(CommonProperties property)
        {
            XmlNode listResponse = null;
            try
            {
                //Get the list details by its name
                listResponse = listService.GetList(property.UploadDocLibraryName);

            }
            catch (Exception ex)
            {
                property.UploadDocLibraryName = string.Empty;
                if (ex.Message.Contains("HTTP status 401: Unauthorized."))
                {
                    return false;
                }
                else
                {
                    try
                    {

                        XmlNode allLists = listService.GetListCollection();
                        XmlDocument allListsDoc = new XmlDocument();
                        allListsDoc.LoadXml(allLists.OuterXml);

                        XmlNamespaceManager ns = new XmlNamespaceManager(allListsDoc.NameTable);
                        ns.AddNamespace("d", allLists.NamespaceURI);


                        XmlNodeList xlist = allListsDoc.SelectNodes("/d:Lists/d:List[starts-with(@DefaultViewUrl,'" + property.UploadFolderNode.ChildNodes[11].InnerText + "')]", ns);
                        if (xlist.Count > 0)
                        {
                            string viewURL = xlist[0].Attributes["DefaultViewUrl"].Value;
                            if (viewURL.StartsWith(property.UploadFolderNode.ChildNodes[11].InnerText))
                            {
                                property.UploadDocLibraryName = xlist[0].Attributes["Title"].Value;
                                //m_uploadLibNameFromAllLists = xlist[0].Attributes["Title"].Value;
                            }

                        }

                    }
                    catch (Exception ex1)
                    {

                    }
                }
            }
            if (string.IsNullOrEmpty(property.UploadDocLibraryName))
            {

                return false;
            }
            return true;
        }




        /// <summary>
        /// <c>GetSPListNameUsingClientOM</c> member function
        /// gets the Sharepoint document library name for sharepoint 2010
        /// </summary>
        /// <param name="property"></param>
        /// <returns></returns>
        public static Boolean GetSPListNameUsingClientOM(CommonProperties property)
        {
            Boolean result = false;
            try
            {

                //Get the list details by its name

                ClientContext clientcontext = new ClientContext(property.LibSite);

                List mylist = clientcontext.Web.Lists.GetByTitle(property.UploadDocLibraryName);
                try
                {
                    clientcontext.Credentials = property.Credentionals;
                    clientcontext.Load(clientcontext.Web);
                    clientcontext.Load(mylist);
                    clientcontext.ExecuteQuery();
                    result = true;
                    clientcontext.Dispose();
                }
                catch (Exception)
                {

                    try
                    {
                        clientcontext.Credentials = new System.Net.NetworkCredential(property.UserName, property.Password); // property.Credentionals;
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.Load(mylist);
                        clientcontext.ExecuteQuery();
                        result = true;
                        clientcontext.Dispose();
                    }
                    catch (Exception)
                    {
                        string strdomainanduser = property.UserName;
                        int indx = strdomainanduser.LastIndexOf("\\");
                        if (indx != -1)
                        {
                            string domain = strdomainanduser.Substring(0, indx);
                            string uname = strdomainanduser.Substring(indx + 1);

                            clientcontext.Credentials = new System.Net.NetworkCredential(uname, property.Password, domain);
                            clientcontext.Load(clientcontext.Web);
                            clientcontext.Load(mylist);
                            clientcontext.ExecuteQuery();
                            result = true;
                            clientcontext.Dispose();

                        }

                    }
                }


            }
            catch (Exception ex)
            {

                result = false;
            }

            return result;
        }


        #region uploadusingclient


        #endregion

        /// <summary>
        /// <c>uploadFilestoLibraryUsingClientOM</c> member function
        /// upload file to sharepoint Document Library using Client object model
        /// </summary>
        /// <param name="udataitem"></param>
        /// <param name="cmproperty"></param>
        /// <returns></returns>
        public static bool uploadFilestoLibraryUsingClientOM(SharePoint_Link.UserModule.UploadItemsData udataitem, CommonProperties cmproperty)
        {
            bool result = false;
            try
            {

              


                CommonProperties cmproperties = new CommonProperties();
                string tempFilePath = UserLogManagerUtility.RootDirectory + "\\" + udataitem.UploadFileName;// +udataitem.UploadFileExtension;


                if (System.IO.File.Exists(tempFilePath))
                {
                    System.IO.File.Delete(tempFilePath);
                }


                FileStream fs = new FileStream(tempFilePath, FileMode.Create, FileAccess.ReadWrite);
              
                //////// out look storage//////////
                
                try
                {
                    if (udataitem.UploadType == TypeOfUploading.Mail)
                    {
                       
                        Stream msgstream = new MemoryStream(cmproperty.FileBytes);
                        
                        Utility.OutlookStorage.Message message = new OutlookStorage.Message(msgstream);
                        if (!string.IsNullOrEmpty(message.BodyText))
                        {
                            //  HashingClass.Hashedemailbody = HashingClass.ComputeHashWithoutSalt(message.BodyText, "MD5");

                        }
                        if (!string.IsNullOrEmpty(message.Subject))
                        {
                            HashingClass.Mailsubject = message.Subject;

                        }
                        if (message.SentDate != "")
                        {
                            HashingClass.Modifieddate = message.SentDate;

                        }




                    }
                }
                catch (Exception)
                {


                }


                //////////End of outlook storage////


                if (cmproperty.FileBytes != null)
                {
                    
                    BinaryWriter bw = new BinaryWriter(fs);
                    bw.Write(cmproperty.FileBytes);
                    bw.Close();
                    
                }


                ///updated new 6jan//////
                string strsitepath = cmproperty.LibSite;
               
                string strleftpart = "/";
                int firstindex = -1;
                int secondindex = -1;
                firstindex = strsitepath.IndexOf(@"//");
                try
                {
                   

                    if (firstindex != -1)
                    {
                        if (strsitepath.Length >= (firstindex + 3))
                        {
                            secondindex = strsitepath.IndexOf(@"/", firstindex + 3);
                            if (secondindex != -1)
                            {
                                if (strsitepath.Length >= (secondindex + 1))
                                {
                                    strleftpart = strleftpart + strsitepath.Substring(secondindex + 1);
                                    if (strsitepath.Length > (secondindex + 1))
                                    {
                                        strsitepath = strsitepath.Remove(secondindex + 1);
                                    }

                                    cmproperty.LibSite = strsitepath;
                                }

                            }

                        }
                        

                    }
                }
                catch (Exception)
                {
                }
               
                string pth = strleftpart + cmproperty.UploadDocLibraryName + "/" + udataitem.UploadFileName;

                // end of new///
                
                Microsoft.SharePoint.Client.ClientContext ct = new ClientContext(cmproperty.LibSite);
               
                ct.Credentials = ListWebClass.GetCredential(cmproperty);
                
                using (FileStream filestream = new FileStream(tempFilePath, FileMode.Open))
                {
                   
                 
                    
                        Microsoft.SharePoint.Client.File.SaveBinaryDirect(ct, pth, filestream, true);
                       
                        
                };

                result = true;


                /////
                // Get ID of Uploaded File
                FileID = GetFileID(udataitem, cmproperty, strleftpart);
                HashingClass.Mailsubject = "";
                HashingClass.Modifieddate = "";
                HashingClass.Hashedemailbody = "";

                ////

                ////////////


                /////////////

                if (System.IO.File.Exists(tempFilePath))
                {
                    System.IO.File.Delete(tempFilePath);
                }




            }
            catch (Exception ex)
            {
                string errmessage = ex.Message.ToString();

                SharePoint_Link.frmUploadItemsList.UpLoadErrorMessage = errmessage;// ex.InnerException.Message.ToString();
                result = false;

            }

            return result;
        }

       
        /// <summary>
        /// <c>GetFileID</c> member function
        /// this member function is used to get ID of the file recently uploaded to sharepoint document library
        /// </summary>
        /// <param name="udataitem"></param>
        /// <param name="cmproperty"></param>
        /// <param name="leftpart"></param>
        /// <returns></returns>
        private static string GetFileID(SharePoint_Link.UserModule.UploadItemsData udataitem, CommonProperties cmproperty, string leftpart)
        {
            string test = cmproperty.SharepointLibraryURL;
            string result = string.Empty;
            try
            {
                Microsoft.SharePoint.Client.ClientContext ct = new ClientContext(cmproperty.LibSite);
                ct.Credentials = ListWebClass.GetCredential(cmproperty);

                result = UpdateAttributes(udataitem, result, ct, cmproperty, cmproperty.UploadDocLibraryName);
                ct.Dispose();
            }
            catch (Exception ex)
            {
                try
                {

                    result = getFileIDByIteration(udataitem, cmproperty, cmproperty.LibSite, cmproperty.UploadDocLibraryName);

                    if (string.IsNullOrEmpty(result))
                    {
                        ////////////////
                        if (leftpart.StartsWith("/"))
                        {
                            leftpart = leftpart.Remove(0, 1);
                        }
                        string temurl = cmproperty.LibSite + leftpart;
                        result = getFileIDAlternate(udataitem, cmproperty, temurl);

                        ///////////////////
                        if (string.IsNullOrEmpty(result))
                        {
                            result = getFileID_SecondAlternate(udataitem, cmproperty, cmproperty.LibSite, leftpart + cmproperty.UploadDocLibraryName);
                        }
                        /////////////////

                        if (string.IsNullOrEmpty(result))
                        {
                            string p = leftpart;

                            string[] split = p.Split('/');
                            int len = split.Length;

                            for (int i = len - 1; i > 0; i--)
                            {

                                int ind = p.LastIndexOf(split[i].ToString());
                                if (ind != -1)
                                {


                                    p = p.Remove(ind);
                                    temurl = cmproperty.LibSite + p;

                                    result = getFileIDAlternate(udataitem, cmproperty, temurl);


                                    if (!string.IsNullOrEmpty(result))
                                    {
                                        break;
                                    }
                                }
                            }

                            if (string.IsNullOrEmpty(result))
                            {
                                for (int i = 1; i < len - 1; i++)
                                {

                                    int ind = p.LastIndexOf(split[i].ToString());
                                    if (ind != -1)
                                    {


                                        p = p.Remove(ind);
                                        temurl = cmproperty.LibSite + p;

                                        result = getFileID_SecondAlternate(udataitem, cmproperty, cmproperty.LibSite, p + cmproperty.UploadDocLibraryName); getFileIDAlternate(udataitem, cmproperty, temurl);


                                        if (!string.IsNullOrEmpty(result))
                                        {
                                            break;
                                        }
                                    }

                                }
                            }

                        }
                    }
                    //////////////////
                }
                catch (Exception)
                {


                }



            }

            return result;
        }

        /// <summary> 
        /// <c>UpdateAttributes</c>   member function
        /// this member function updates metatags in sharepoint document library
        /// </summary>
        /// <param name="udataitem"></param>
        /// <param name="result"></param>
        /// <param name="ct"></param>
        /// <param name="cmproperty"></param>
        /// <param name="UploadDocLibraryName"></param>
        /// <returns></returns>
        private static string UpdateAttributes(SharePoint_Link.UserModule.UploadItemsData udataitem, string result, Microsoft.SharePoint.Client.ClientContext ct, CommonProperties cmproperty, string UploadDocLibraryName)
        {

            List list;
            string doclibtitle = UploadDocLibraryName;
            try
            {
                list = ct.Web.Lists.GetByTitle(UploadDocLibraryName);
                ct.Load(list); ct.ExecuteQuery();
            }
            catch (Exception ex)
            {
                ListCollection lc = ct.Web.Lists;
                ct.Load(lc); ct.ExecuteQuery();
                foreach (List item in lc)
                {
                    string siteurlroot = cmproperty.LibSite.Remove(cmproperty.LibSite.Length - 1);
                    string url = siteurlroot + item.DefaultViewUrl;
                    if (url == cmproperty.CompletedoclibraryURL)
                    {
                        doclibtitle = item.Title;
                        break;
                    }
                }

            }

            list = ct.Web.Lists.GetByTitle(doclibtitle);
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml =
            @"<View>
                <Query>
                  <Where>
                    <Eq>
                      <FieldRef Name='FileLeafRef' />
                      <Value Type='Text' >" + udataitem.UploadFileName + @" </Value>
                    </Eq>
                  </Where>
                </Query>
                <RowLimit>100</RowLimit>
              </View>";
            ListItemCollection listItems = list.GetItems(camlQuery);

            ct.Load(listItems);
            ct.ExecuteQuery();


            foreach (ListItem l in listItems)
            {
                result = l["ID"].ToString();

                try
                {

                    l["Title"] = HashingClass.Mailsubject;
                    l.Update();
                    ct.ExecuteQuery();

                    l["ModifiedDate"] = HashingClass.Modifieddate;

                    l.Update();
                    ct.ExecuteQuery();


                }
                catch (Exception ex)
                {

                }

            }
            return result;
        }

        /// <summary>
        /// <c>getFileIDAlternate</c> member function
        /// gets ID of the file uploaded to sharepoint document library when <c>GetFileID</c> member function  throw exception
        /// </summary>
        /// <param name="udataitem"></param>
        /// <param name="cmproperty"></param>
        /// <param name="tempsiteurl"></param>
        /// <returns></returns>
        private static string getFileIDAlternate(SharePoint_Link.UserModule.UploadItemsData udataitem, CommonProperties cmproperty, string tempsiteurl)
        {
            string result = string.Empty;

            try
            {

                Microsoft.SharePoint.Client.ClientContext ct = new ClientContext(tempsiteurl);
                ct.Credentials = ListWebClass.GetCredential(cmproperty);
                result = UpdateAttributes(udataitem, result, ct, cmproperty, cmproperty.UploadDocLibraryName);
                List list = ct.Web.Lists.GetByTitle(cmproperty.UploadDocLibraryName);

                ct.Dispose();
            }
            catch (Exception)
            {

                result = getFileIDByIteration(udataitem, cmproperty, tempsiteurl, cmproperty.UploadDocLibraryName);

            }
            return result;
        }

        /// <summary>
        /// <c>getFileID_SecondAlternate</c> member function
        /// gets ID of the file uploaded to sharepoint document library when <c>getFileIDAlternate</c> member function  throw exception
        /// </summary>
        /// <param name="udataitem"></param>
        /// <param name="cmproperty"></param>
        /// <param name="siteurl"></param>
        /// <param name="docLibrary"></param>
        /// <returns></returns>
        private static string getFileID_SecondAlternate(SharePoint_Link.UserModule.UploadItemsData udataitem, CommonProperties cmproperty, string siteurl, string docLibrary)
        {
            string result = string.Empty;

            try
            {

                Microsoft.SharePoint.Client.ClientContext ct = new ClientContext(siteurl);
                ct.Credentials = ListWebClass.GetCredential(cmproperty);
                result = UpdateAttributes(udataitem, result, ct, cmproperty, docLibrary);

                ct.Dispose();
            }
            catch (Exception)
            {

                result = getFileIDByIteration(udataitem, cmproperty, siteurl, docLibrary);

            }
            return result;
        }

        /// <summary>
        /// <c>getFileIDByIteration</c> member function
        /// gets ID of the file uploaded to sharepoint document library when <c>getFileID_SecondAlternate</c> member function  throw exception
        /// </summary>
        /// <param name="udataitem"></param>
        /// <param name="cmproperty"></param>
        /// <param name="siteurl"></param>
        /// <param name="docLibrary"></param>
        /// <returns></returns>
        private static string getFileIDByIteration(SharePoint_Link.UserModule.UploadItemsData udataitem, CommonProperties cmproperty, string siteurl, string docLibrary)
        {
            string result = string.Empty;
            Microsoft.SharePoint.Client.ClientContext ct = new ClientContext(siteurl);
            try
            {

                ct.Credentials = ListWebClass.GetCredential(cmproperty);
                Web oWeb = ct.Web;
                ct.Load(oWeb);
                ct.ExecuteQuery();
                ListCollection currentListCollection = oWeb.Lists;
                ct.Load(currentListCollection);
                ct.ExecuteQuery();
                int found = 0;

                foreach (List oList in currentListCollection)
                {


                    string one = docLibrary.Trim();
                    string two = oList.Title.Trim();
                    int indexofspace = -1;
                    indexofspace = one.IndexOf(' ');
                    while (indexofspace != -1)
                    {
                        one = one.Remove(indexofspace, 1);
                        indexofspace = one.IndexOf(' ');
                    }
                    indexofspace = two.IndexOf(' ');
                    while (indexofspace != -1)
                    {
                        two = two.Remove(indexofspace, 1);
                        indexofspace = two.IndexOf(' ');
                    }

                    if (one == two)
                    {
                        List list2 = ct.Web.Lists.GetByTitle(oList.Title);
                        result = UpdateAttributes(udataitem, result, ct, cmproperty, docLibrary);
                        //old removed
                        found += 1;
                    }

                    if (found > 0)
                    {
                        break;
                    }

                }
            }
            catch (Exception)
            {

                try
                {
                    result = UpdateAttributes(udataitem, result, ct, cmproperty, docLibrary);
                }
                catch
                { }
            }
            ct.Dispose();
            return result;
        }


        /// <summary>
        /// <c>GetCredential</c> member function
        /// gets the user credentials based on configuration settings and returns  ICredentials
        /// </summary>
        /// <param name="cProp"></param>
        /// <returns></returns>
        public static System.Net.ICredentials GetCredential(CommonProperties cProp)
        {
            System.Net.ICredentials result = null;
            try
            {

                //Get the list details by its name

                ClientContext clientcontext = new ClientContext(cProp.LibSite);

                List mylist = clientcontext.Web.Lists.GetByTitle(cProp.UploadDocLibraryName);
                try
                {
                    result = cProp.Credentionals;
                    clientcontext.Credentials = result;
                    clientcontext.Load(clientcontext.Web);
                    clientcontext.Load(mylist);
                    clientcontext.ExecuteQuery();
                    result = cProp.Credentionals;
                    clientcontext.Dispose();
                }
                catch (Exception)
                {

                    try
                    {
                        clientcontext.Credentials = new System.Net.NetworkCredential(cProp.UserName, cProp.Password); // property.Credentionals;
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.Load(mylist);
                        clientcontext.ExecuteQuery();
                        result = new System.Net.NetworkCredential(cProp.UserName, cProp.Password);
                        clientcontext.Dispose();
                    }
                    catch (Exception)
                    {
                        string strdomainanduser = cProp.UserName;
                        int indx = strdomainanduser.LastIndexOf("\\");
                        if (indx != -1)
                        {
                            string domain = strdomainanduser.Substring(0, indx);
                            string uname = strdomainanduser.Substring(indx + 1);
                            try
                            {
                                clientcontext.Credentials = new System.Net.NetworkCredential(uname, cProp.Password, domain);
                                clientcontext.Load(clientcontext.Web);
                                clientcontext.Load(mylist);
                                clientcontext.ExecuteQuery();
                                result = new System.Net.NetworkCredential(uname, cProp.Password, domain);
                                clientcontext.Dispose();
                            }
                            catch (Exception)
                            {

                                result = cProp.Credentionals;
                            }


                        }

                    }
                }


            }
            catch (Exception ex)
            {


            }
            return result;
        }

        /// <summary>
        /// <c>Log</c> member function
        /// logs the exceptions in LogFile
        /// </summary>
        /// <param name="message"></param>
        /// <param name="newlog"></param>
        public static void Log(string message, bool newlog)
        {
            try
            {

                string path = System.Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                path = path + "\\ITOPIALOGS";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                path = path + "\\MetaTags.txt";
                if (System.IO.File.Exists(path))
                {
                    System.IO.File.Delete(path);
                }
                FileStream fs = null;
                if (!System.IO.File.Exists(path))
                {
                    using (fs = System.IO.File.Create(path))
                    {

                    }
                }
                if (System.IO.File.Exists(path))
                {
                    using (StreamWriter sw = System.IO.File.AppendText(path))
                    {
                        if (newlog == true)
                        {
                            sw.WriteLine("");
                            sw.WriteLine("----------" + System.DateTime.Now + "-------------------");
                            sw.WriteLine(message);
                        }
                        else
                        {

                            sw.WriteLine(message);
                        }



                    }
                }



            }
            catch (Exception)
            {


            }
        }

        /// <summary>
        /// <c>ComputeHashSP07</c>
        /// calls <c>ComputeHashWithoutSalt</c> function to calculate hash value of string
        /// </summary>
        /// <param name="FileBytes"></param>
        /// <param name="msgbody"></param>
        public static void ComputeHashSP07(byte[] FileBytes, string msgbody)
        {

            try
            {


                Stream msgstream = new MemoryStream(FileBytes);


                Utility.OutlookStorage.Message message = new OutlookStorage.Message(msgstream);

                string tstval = message.BCCTest;

                string to = "";
                to = message.ReceivedByEmail;
                string ccs = "";

                foreach (OutlookStorage.Recipient r in message.Recipients)
                {
                    if (r.Type == OutlookStorage.RecipientType.CC)
                    {
                        ccs += r.DisplayName;
                    }
                    else
                    {
                        if (r.Type == OutlookStorage.RecipientType.To)
                        {
                            to = r.DisplayName;
                        }
                    }
                }


                string subject = message.Subject;
                HtmlFromRtf htmlfromrtf = new HtmlFromRtf();
                string strmessagebody = htmlfromrtf.GetRefinedHtmlFromRtf(message.BodyRTF); // msgbody;
                //try
                //{
                //    System.Windows.Forms.RichTextBox rt = new System.Windows.Forms.RichTextBox();
                //    rt.Rtf = message.BodyRTF;
                //    strmessagebody = rt.Text;
                //}
                //catch (Exception ex)
                //{
                //    strmessagebody = message.BodyText;
                //}
                //try
                //{
                //    while (strmessagebody.IndexOf("\r\n") != -1)
                //    {
                //        strmessagebody = strmessagebody.Remove(strmessagebody.IndexOf("\r\n"), 2);
                //    }

                //}
                //catch (Exception)
                //{


                //}

                string strtime = "";
                string msgdate = message.SentDate;
                if (string.IsNullOrEmpty(msgdate))
                {
                    msgdate = msgbody;
                }
                try
                {
                    msgdate = Convert.ToDateTime(msgdate).ToString("M/d/yyyy hh:mm tt zz");
                    strtime = Convert.ToDateTime(msgdate).ToString("hh:mm tt");
                }
                catch (Exception)
                {

                }
                Log("Hashing Log 10:", true);

                string strfinalvalue = to + ccs + subject + strmessagebody;// +strtime;

                Log("tobehash" + strfinalvalue, false);

                if (!string.IsNullOrEmpty(strfinalvalue))
                {
                    string v = HashingClass.ComputeHashWithoutSalt(strfinalvalue, "MD5");
                    HashingClass.Hashedemailbody = v;

                }



                if (!string.IsNullOrEmpty(message.Subject))
                {
                    HashingClass.Mailsubject = message.Subject;
                }
                else
                {
                    try
                    {
                        HashingClass.Mailsubject = message.Subject;
                    }
                    catch (Exception)
                    {


                    }
                }

                if (message.SentDate != "")
                {
                    HashingClass.Modifieddate = msgdate;// message.SentDate;

                }
                else
                {
                    if (message.SentDate != null)
                    {
                        HashingClass.Modifieddate = msgdate;// message.SentDate;
                    }
                }


            }
            catch (Exception)
            {
            }

        }


        /// <summary>
        /// <c>WebViewUrl</c> member function
        /// returns the  url for outlook  folder to view the mapped sharepoint document library
        /// inside outlook
        /// </summary>
        /// <param name="strurl"></param>
        /// <returns></returns>
        public static string WebViewUrl(string strurl)
        {

            string result = strurl;

            try
            {

                string path = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + "MaPiFolderTemp.htm";
                path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Itopia\MaPiFolderTemp.htm";

                if (System.IO.File.Exists(path))
                {
                    System.IO.File.Delete(path);
                }

                FileStream fstream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite);
                TextWriter twrite = new StreamWriter(fstream);
                StringBuilder strbuilder = new StringBuilder();

                strbuilder.AppendLine("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">");
                strbuilder.AppendLine("<html xmlns=\"http://www.w3.org/1999/xhtml\"><head>");
                strbuilder.AppendLine("<meta http-equiv=\"content-type\" content=\"text/html; charset=UTF-8\">");
                strbuilder.AppendLine("<div style=\"text-align:center; background-color:Gray; color:White; font-weight:bold;\"> Please Wait...</div> <title></title>");
                strbuilder.Append("<script type=\"text/javascript\">");

                strbuilder.Append(" function qs(search_for) { ");
                strbuilder.Append(" var query = window.location.search.substring(1); ");
                strbuilder.Append(" var parms = query.split('&'); ");
                strbuilder.Append("  for (var i = 0; i < parms.length; i++) { ");
                strbuilder.Append(" var pos = parms[i].indexOf('='); ");
                strbuilder.Append(" if (pos > 0 && search_for == parms[i].substring(0, pos)) { ");
                strbuilder.Append(" return parms[i].substring(pos + 1);  ");
                strbuilder.Append("  } ");
                strbuilder.Append("  } ");
                strbuilder.Append("  return \"\"; ");
                strbuilder.Append("  } ");
                strbuilder.Append("var returnurl = qs(\"ReturnUrl\");");

                if (strurl.Trim().Contains(path.Trim()))
                {
                    strbuilder.Append(" var action ='no'; ");
                }
                else
                {
                    strbuilder.Append(" var action ='view'; ");
                }

                strbuilder.Append(" if (action == 'view') { ");
                strbuilder.Append(" window.location =' ");
                // strbuilder.Append("returnurl");
                strbuilder.Append(strurl);
                strbuilder.Append("';   }  </script>");
                strbuilder.AppendLine("</head><body></body></html>");
                twrite.WriteLine(strbuilder.ToString());
                twrite.Close();



                result = path;
                fstream.Close();
                fstream.Dispose();


            }
            catch (Exception)
            {


            }



            return result;
        }

        /// <summary>
        /// <c>DocumentLibraryName</c>
        /// not currently used
        /// </summary>
        /// <param name="strurl"></param>
        /// <returns></returns>
        public static string DocumentLibraryName(string strurl)
        {
            string result = string.Empty;
            try
            {

            }
            catch (Exception)
            {


            }

            return result;
        }

    
    }
}

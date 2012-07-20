using System;
using System.Text;
using System.Net;

namespace SharePoint_Link.Utility
{

    /// <summary>
    /// <c>SPVersionClass</c> class
    /// implements the functionalities to find whether   url is from sharepoint 2007 or sharepoint 2010
    /// </summary>
    class SPVersionClass
    {
        /// <summary>
        /// <c>sPSiteVersion</c> member field of type string
        /// it holds the  value SP2010 if it is sharepoint 2010 or SP2007 in case of sharepoint 2007
        /// otherwise NotSP
        /// </summary>
        private static string sPSiteVersion = string.Empty;

        /// <summary>
        /// <c>SPSiteVersion</c> member property
        /// encapsulates  sPSiteVersion
        /// </summary>
        public static string SPSiteVersion
        {
            get { return SPVersionClass.sPSiteVersion; }
            set { SPVersionClass.sPSiteVersion = value; }
        }

        /// <summary>
        /// <c>SiteVersion</c> member field of type enum
        /// holds  values to find either the site is sharepoint 2010, 2007 or not a sharepoint site
        /// </summary>
        public enum SiteVersion
        {
            SP2007,
            SP2010,
            NotSP

        }




        /// <summary>
        /// <c>GetSPVersionFromUrl</c> member function
        /// it calls <c>FindVersion</c> member function to determines sharepoint site version(2007,2010) based on url.
        /// </summary>
        /// <param name="url"></param>
        /// <param name="UName"></param>
        /// <param name="PWord"></param>
        /// <param name="FolderAuthenticationType"></param>
        /// <returns></returns>
        public static string GetSPVersionFromUrl(string url, string UName, string PWord, AuthenticationType FolderAuthenticationType)
        {

            string spVersion = string.Empty;
            try
            {

                spVersion = SPVersionClass.FindVersion(url);
            }
            catch (Exception)
            {
                if (FolderAuthenticationType == AuthenticationType.Manual)
                {


                    try
                    {
                        spVersion = SPVersionClass.FindVersionUserNamePassword(url, UName, PWord);
                    }
                    catch (Exception)
                    {

                        spVersion = SPVersionClass.FindVersionUserNamePasswordDomain(url, UName, PWord);
                    }
                }
                else
                {
                    if (spVersion == "" || spVersion == null)
                    {
                        spVersion = SiteVersion.SP2010.ToString();
                    }
                }
            }

            return spVersion;


        }

        /// <summary>
        /// <c>FindVersion</c> member function
        /// finds sharepoint version  wether it is sharepoint 2007 or sharepoint 2010
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static string FindVersion(string url)
        {
            string result = string.Empty;
            try
            {
                Uri uri = new Uri(url);

                WebRequest request = WebRequest.Create((uri));

                ///////
                IWebProxy proxy = WebRequest.GetSystemWebProxy();
                proxy.Credentials = CredentialCache.DefaultCredentials;

                request.Proxy = proxy;
                request.Method = "GET";

                request.Timeout = -1;

                /////////

                WebResponse response = request.GetResponse();

                if (response != null && response.Headers["MicrosoftSharePointTeamServices"] != null)
                {
                    result = response.Headers["MicrosoftSharePointTeamServices"].ToString();
                }
                //
                if (result.Trim().StartsWith("12"))
                {
                    result = SiteVersion.SP2007.ToString();
                }
                else
                {
                    if (result.Trim().StartsWith("14"))
                    {
                        result = SiteVersion.SP2010.ToString();
                    }
                    result = SiteVersion.SP2010.ToString();

                }
            }
            catch (Exception)
            {



                string line;
                StringBuilder sb = new StringBuilder();
                HttpWebResponse resp;
                HttpWebRequest rqs = (HttpWebRequest)WebRequest.Create(url);
                rqs.MaximumAutomaticRedirections = 4;
                rqs.MaximumResponseHeadersLength = 4;
                // either user default credentials or the credentials specified on the

                rqs.Credentials = CredentialCache.DefaultCredentials;



                // send the request to the remote site, get the response.
                resp = (HttpWebResponse)rqs.GetResponse();
                if (resp != null && resp.Headers["MicrosoftSharePointTeamServices"] != null)
                {
                    result = resp.Headers["MicrosoftSharePointTeamServices"].ToString();
                }
                //
                if (result.Trim().StartsWith("12"))
                {
                    result = SiteVersion.SP2007.ToString();
                }
                else
                {
                    //if (result.Trim().StartsWith("14"))
                    //{
                    //    result = SiteVersion.SP2010.ToString();
                    //}
                    //else
                    //{
                    //    result = SiteVersion.NotSP.ToString();
                    //}
                    result = SiteVersion.SP2010.ToString();
                }
                //

            }
            return result;
        }


        /// <summary>
        /// <c>FindVersionUserNamePassword</c> member function 
        /// finds sharepoint version using <c>WebRequest</c> object
        /// </summary>
        /// <param name="url"></param>
        /// <param name="UName"></param>
        /// <param name="PWord"></param>
        /// <returns></returns>
        public static string FindVersionUserNamePassword(string url, string UName, string PWord)
        {
            string result = string.Empty;
            try
            {
                Uri uri = new Uri(url);

                WebRequest request = WebRequest.Create((uri));



                ///////
                IWebProxy proxy = WebRequest.GetSystemWebProxy();

                proxy.Credentials = new System.Net.NetworkCredential(UName, PWord);

                request.Proxy = proxy;
                request.Method = "GET";

                request.Timeout = -1;

                /////////

                WebResponse response = request.GetResponse();

                if (response != null && response.Headers["MicrosoftSharePointTeamServices"] != null)
                {
                    result = response.Headers["MicrosoftSharePointTeamServices"].ToString();
                }
                //
                if (result.Trim().StartsWith("12"))
                {
                    result = SiteVersion.SP2007.ToString();
                }
                else
                {
                    if (result.Trim().StartsWith("14"))
                    {
                        result = SiteVersion.SP2010.ToString();
                    }
                    else
                    {
                        result = SiteVersion.NotSP.ToString();
                    }
                    result = SiteVersion.SP2010.ToString();
                }
            }
            catch (Exception)
            {



                string line;
                StringBuilder sb = new StringBuilder();
                HttpWebResponse resp;
                HttpWebRequest rqs = (HttpWebRequest)WebRequest.Create(url);
                rqs.MaximumAutomaticRedirections = 4;
                rqs.MaximumResponseHeadersLength = 4;
                // either user default credentials or the credentials specified on the

                System.Net.ICredentials cred = new System.Net.NetworkCredential(UName, PWord);
                rqs.Credentials = cred;



                // send the request to the remote site, get the response.
                resp = (HttpWebResponse)rqs.GetResponse();
                if (resp != null && resp.Headers["MicrosoftSharePointTeamServices"] != null)
                {
                    result = resp.Headers["MicrosoftSharePointTeamServices"].ToString();
                }
                //
                if (result.Trim().StartsWith("12"))
                {
                    result = SiteVersion.SP2007.ToString();
                }
                else
                {
                    if (result.Trim().StartsWith("14"))
                    {
                        result = SiteVersion.SP2010.ToString();
                    }
                    else
                    {
                        result = SiteVersion.NotSP.ToString();
                    }
                    result = SiteVersion.SP2010.ToString();
                }
                //

            }
            return result;
        }


        /// <summary>
        /// <c>FindVersionUserNamePasswordDomain</c> member function
        /// find sharepoint version if <c>FindVersionUserNamePassword</c> cannot find version
        /// </summary>
        /// <param name="url"></param>
        /// <param name="UName"></param>
        /// <param name="PWord"></param>
        /// <returns></returns>
        public static string FindVersionUserNamePasswordDomain(string url, string UName, string PWord)
        {
            string result = string.Empty;
            try
            {
                Uri uri = new Uri(url);

                WebRequest request = WebRequest.Create((uri));



                ///////
                IWebProxy proxy = WebRequest.GetSystemWebProxy();
                int indx = UName.LastIndexOf("\\");
                if (indx != -1)
                {
                    string domain = UName.Substring(0, indx);
                    string uname = UName.Substring(indx + 1);

                    System.Net.ICredentials cred = new System.Net.NetworkCredential(uname, PWord, domain);

                    proxy.Credentials = cred;

                }


                request.Proxy = proxy;
                request.Method = "GET";

                request.Timeout = -1;

                /////////

                WebResponse response = request.GetResponse();

                if (response != null && response.Headers["MicrosoftSharePointTeamServices"] != null)
                {
                    result = response.Headers["MicrosoftSharePointTeamServices"].ToString();
                }
                //
                if (result.Trim().StartsWith("12"))
                {
                    result = SiteVersion.SP2007.ToString();
                }
                else
                {
                    if (result.Trim().StartsWith("14"))
                    {
                        result = SiteVersion.SP2010.ToString();
                    }
                    else
                    {
                        result = SiteVersion.NotSP.ToString();
                    }
                    result = SiteVersion.SP2010.ToString();
                }
            }
            catch (Exception)
            {



                string line;
                StringBuilder sb = new StringBuilder();
                HttpWebResponse resp;
                HttpWebRequest rqs = (HttpWebRequest)WebRequest.Create(url);
                rqs.MaximumAutomaticRedirections = 4;
                rqs.MaximumResponseHeadersLength = 4;
                // either user default credentials or the credentials specified on the
                int indx = UName.LastIndexOf("\\");
                if (indx != -1)
                {
                    string domain = UName.Substring(0, indx);
                    string uname = UName.Substring(indx + 1);

                    System.Net.ICredentials cred = new System.Net.NetworkCredential(uname, PWord, domain);

                    rqs.Credentials = cred;

                }

                


                // send the request to the remote site, get the response.
                resp = (HttpWebResponse)rqs.GetResponse();
                if (resp != null && resp.Headers["MicrosoftSharePointTeamServices"] != null)
                {
                    result = resp.Headers["MicrosoftSharePointTeamServices"].ToString();
                }
                //
                if (result.Trim().StartsWith("12"))
                {
                    result = SiteVersion.SP2007.ToString();
                }
                else
                {
                    if (result.Trim().StartsWith("14"))
                    {
                        result = SiteVersion.SP2010.ToString();
                    }
                    else
                    {
                        result = SiteVersion.NotSP.ToString();
                    }
                    result = SiteVersion.SP2010.ToString();
                }
                //

            }
            return result;
        }
    }
}

using System;
using System.DirectoryServices;
using System.Collections;

namespace SharePoint_Link.UserModule
{
    class ActiveDirectoryChecking
    {
        /// <summary>
        /// Methodt o authenticate the user datails with ActiveDirectory
        /// </summary>
        /// <param name="userName">UserName as String</param>
        /// <param name="password">Password as String</param>
        /// <param name="domain">ActiveDirctory domain name as String.</param>
        /// <returns>True/False as Boolean.</returns>
        public bool Authenticate(string userName, string password, string domain)
        {
            bool authentic = false;
            try
            {
                
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + domain, userName, password);
                object nativeObject = entry.NativeObject;
                authentic = true;
            }
            catch (DirectoryServicesCOMException) { }
            return authentic;
        }

        public ArrayList GetDomains()
        {
            ArrayList arrDomains = new ArrayList();
            DirectoryEntry ParentEntry = new DirectoryEntry();
            try
            {
                ParentEntry.Path = "WinNT:";
                foreach (DirectoryEntry childEntry in ParentEntry.Children)
                {
                    switch (childEntry.SchemaClassName)
                    {
                        case "Domain":
                            {
                                arrDomains.Add(childEntry.Name);
                                break;
                            }
                        default:
                            {
                                break;
                            }
                    }
                }
            }
            catch (Exception e)
            {
            }
            finally
            {
                ParentEntry = null;
            }
            return arrDomains;
        }
    }
}

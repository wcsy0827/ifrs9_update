using System;
using System.Configuration;
using System.DirectoryServices;
using System.Security.Principal;

namespace Transfer.Utility
{
    public static class LdapAuthentication
    {
        public static bool isAuthenticatrd(string userId, string pwd)
        {
            try
            {
                string strAdDns = Properties.Settings.Default["ADDomain"]?.ToString();
          
                using (DirectoryEntry entry =
                    new DirectoryEntry(strAdDns, userId, pwd))
                {
                    Object obj = entry.NativeObject;
                    DirectorySearcher search = new DirectorySearcher(entry);
                    search.Filter = "(SAMAccountName=" + userId + ")";
                    search.PropertiesToLoad.Add("cn");
                    SearchResult result = search.FindOne();

                    return (result != null);

                    //string objectSid = (new SecurityIdentifier((byte[])entry.Properties["objectSid"].Value, 0).Value);
                    //return (objectSid != null);
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
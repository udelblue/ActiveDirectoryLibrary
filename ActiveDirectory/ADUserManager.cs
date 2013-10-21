using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.Configuration;

namespace ActiveDirectoryHelper
{
    public class ADUserManager
    {
        private DirectoryEntry _directoryEntry = null;

        private DirectoryEntry SearchRoot
        {
            get
            {
                if (_directoryEntry == null)
                {
                    _directoryEntry = new DirectoryEntry(ADConfig.LDAPPath, ADConfig.LDAPUser, ADConfig.LDAPPassword, AuthenticationTypes.Secure);
                }
                return _directoryEntry;
            }
        }

 
        /// <summary>
        /// Get user by fullname
        /// </summary>
        /// <param name="userName"></param>
        /// <returns></returns>
        public ADUser GetUserByFullName(String userName)
        {
            try
            {
                _directoryEntry = null;
                DirectorySearcher directorySearch = new DirectorySearcher(SearchRoot);
                directorySearch.Filter = "(&(objectClass=user)(cn=" + userName + "))";
                SearchResult results = directorySearch.FindOne();

                if (results != null)
                {
                    DirectoryEntry user = new DirectoryEntry(results.Path, ADConfig.LDAPUser, ADConfig.LDAPPassword);
                    return ADUser.GetUser(user);
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }



   

        /// <summary>
        /// Get ADUser by login name
        /// </summary>
        /// <param name="userName"></param>
        /// <returns></returns>
        public ADUser GetUserByLoginName(String userName)
        {
            try
            {
                _directoryEntry = null;
                DirectorySearcher directorySearch = new DirectorySearcher(SearchRoot);
                directorySearch.Filter = "(&(objectClass=user)(SAMAccountName=" + userName + "))";
                SearchResult results = directorySearch.FindOne();

                if (results != null)
                {
                    DirectoryEntry user = new DirectoryEntry(results.Path, ADConfig.LDAPUser, ADConfig.LDAPPassword);
                    return ADUser.GetUser(user);
                }
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }


        /// <summary>
        /// This function will take a DL or Group name and return list of users
        /// </summary>
        /// <param name="groupName"></param>
        /// <returns></returns>
        public List<ADUser> GetUserFromGroup(String groupName)
        {
            List<ADUser> userlist = new List<ADUser>();
            try
            {
                _directoryEntry = null;
                DirectorySearcher directorySearch = new DirectorySearcher(SearchRoot);
                directorySearch.Filter = "(&(objectClass=group)(SAMAccountName=" + groupName + "))";
                SearchResult results = directorySearch.FindOne();
                if (results != null)
                {

                    DirectoryEntry deGroup = new DirectoryEntry(results.Path, ADConfig.LDAPUser, ADConfig.LDAPPassword);
                    System.DirectoryServices.PropertyCollection pColl = deGroup.Properties;
                    int count = pColl["member"].Count;


                    for (int i = 0; i < count; i++)
                    {
                        string respath = results.Path;
                        string[] pathnavigate = respath.Split("CN".ToCharArray());
                        respath = pathnavigate[0];
                        string objpath = pColl["member"][i].ToString();
                        string path = respath + objpath;


                        DirectoryEntry user = new DirectoryEntry(path, ADConfig.LDAPUser, ADConfig.LDAPPassword);
                        ADUser userobj = ADUser.GetUser(user);
                        userlist.Add(userobj);
                        user.Close();
                    }
                }
                return userlist;
            }
            catch (Exception ex)
            {
                return userlist;
            }

        }

    

        /// <summary>
        /// Returns a list of ADUsers
        /// </summary>
        /// <param name="fName"></param>
        /// <returns></returns>
        public List<ADUser> GetUsersByFirstName(string fName)
        {

            //UserProfile user;
            List<ADUser> userlist = new List<ADUser>();
            string filter = "";

            _directoryEntry = null;
            DirectorySearcher directorySearch = new DirectorySearcher(SearchRoot);
            directorySearch.Asynchronous = true;
            directorySearch.CacheResults = true;
            filter = string.Format("(givenName={0}*", fName);
            //            filter = "(&(objectClass=user)(objectCategory=person)(givenName="+fName+ "*))";


            directorySearch.Filter = filter;

            SearchResultCollection userCollection = directorySearch.FindAll();
            foreach (SearchResult users in userCollection)
            {
                DirectoryEntry userEntry = new DirectoryEntry(users.Path, ADConfig.LDAPUser, ADConfig.LDAPPassword);
                ADUser userInfo = ADUser.GetUser(userEntry);

                userlist.Add(userInfo);

            }

            directorySearch.Filter = "(&(objectClass=group)(SAMAccountName=" + fName + "*))";
            SearchResultCollection results = directorySearch.FindAll();
            if (results != null)
            {

                foreach (SearchResult r in results)
                {
                    DirectoryEntry deGroup = new DirectoryEntry(r.Path, ADConfig.LDAPUser, ADConfig.LDAPPassword);

                    ADUser agroup = ADUser.GetUser(deGroup);
                    userlist.Add(agroup);
                }

            }
            return userlist;
        }

    


     /// <summary>
     /// Add ADUser to Group
     /// </summary>
     /// <param name="userlogin"></param>
     /// <param name="groupName"></param>
     /// <returns></returns>
        public bool AddUserToGroup(string userlogin, string groupName)
        {
            try
            {
                _directoryEntry = null;
                ADManager admanager = new ADManager(ADConfig.LDAPDomain, ADConfig.LDAPUser, ADConfig.LDAPPassword);
                admanager.AddUserToGroup(userlogin, groupName);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
     

     /// <summary>
     /// Remove ADUser from group
     /// </summary>
     /// <param name="userlogin"></param>
     /// <param name="groupName"></param>
     /// <returns></returns>
        public bool RemoveUserToGroup(string userlogin, string groupName)
        {
            try
            {
                _directoryEntry = null;
                ADManager admanager = new ADManager("xxx", ADConfig.LDAPUser, ADConfig.LDAPPassword);
                admanager.RemoveUserFromGroup(userlogin, groupName);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
      
    }
}
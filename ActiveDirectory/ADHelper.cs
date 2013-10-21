using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.Configuration;
using System.Collections;

namespace ActiveDirectoryHelper
{
    public class ADHelper
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
        /// Authenticate a User Against the Directory domain in app.config
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public bool Authenticate(string userName, string password)
        {
            bool authentic = false;
            try
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + ADConfig.LDAPDomain,
                    userName, password);
                object nativeObject = entry.NativeObject;
                authentic = true;
            }
            catch (DirectoryServicesCOMException) { }
            return authentic;
        }


        /// <summary>
        /// Authenticate a User Against the Directory
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <param name="domain"></param>
        /// <returns></returns>
        public bool Authenticate(string userName, string password, string domain)
        {
            bool authentic = false;
            try
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + domain,
                    userName, password);
                object nativeObject = entry.NativeObject;
                authentic = true;
            }
            catch (DirectoryServicesCOMException) { }
            return authentic;
        }

        /// <summary>
        /// Add User to Group
        /// </summary>
        /// <param name="userDn"></param>
        /// <param name="groupDn"></param>
        public void AddToGroup(string userDn, string groupDn)
        {
            try
            {
                DirectoryEntry dirEntry = new DirectoryEntry("LDAP://" + groupDn);
                dirEntry.Properties["member"].Add(userDn);
                dirEntry.CommitChanges();
                dirEntry.Close();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //doSomething with E.Message.ToString();

            }
        }

        /// <summary>
        /// Remove User from Group
        /// </summary>
        /// <param name="userDn"></param>
        /// <param name="groupDn"></param>
        public void RemoveUserFromGroup(string userDn, string groupDn)
        {
            try
            {
                DirectoryEntry dirEntry = new DirectoryEntry("LDAP://" + groupDn);
                dirEntry.Properties["member"].Remove(userDn);
                dirEntry.CommitChanges();
                dirEntry.Close();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //doSomething with E.Message.ToString();

            }
        }



        /// <summary>
        /// Create User Account
        /// </summary>
        /// <param name="ldapPath"></param>
        /// <param name="userName"></param>
        /// <param name="userPassword"></param>
        /// <returns></returns>
        public string CreateUserAccount(string ldapPath, string userName,
    string userPassword)
        {
            string oGUID = string.Empty;

            try
            {
                
                string connectionPrefix = "LDAP://" + ldapPath;
                DirectoryEntry dirEntry = new DirectoryEntry(connectionPrefix);
                DirectoryEntry newUser = dirEntry.Children.Add
                    ("CN=" + userName, "user");
                newUser.Properties["samAccountName"].Value = userName;
                newUser.CommitChanges();
                oGUID = newUser.Guid.ToString();

                newUser.Invoke("SetPassword", new object[] { userPassword });
                newUser.CommitChanges();
                dirEntry.Close();
                newUser.Close();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //DoSomethingwith --> E.Message.ToString();

            }
            return oGUID;
        }

        /// <summary>
        /// Enable User Account
        /// </summary>
        /// <param name="userDn"></param>
        public void Enable(string userDn)
        {
            try
            {
                DirectoryEntry user = new DirectoryEntry(userDn);
                int val = (int)user.Properties["userAccountControl"].Value;
                user.Properties["userAccountControl"].Value = val & ~0x2;
                //ADS_UF_NORMAL_ACCOUNT;

                user.CommitChanges();
                user.Close();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //DoSomethingWith --> E.Message.ToString();

            }
        }

        /// <summary>
        /// Disable User Account
        /// </summary>
        /// <param name="userDn"></param>
        public void Disable(string userDn)
        {
            try
            {
                DirectoryEntry user = new DirectoryEntry(userDn);
                int val = (int)user.Properties["userAccountControl"].Value;
                user.Properties["userAccountControl"].Value = val | 0x2;
                //ADS_UF_ACCOUNTDISABLE;

                user.CommitChanges();
                user.Close();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //DoSomethingWith --> E.Message.ToString();

            }
        }

        /// <summary>
        /// Unlock User Account
        /// </summary>
        /// <param name="userDn"></param>
        public void Unlock(string userDn)
        {
            try
            {
                DirectoryEntry uEntry = new DirectoryEntry(userDn);
                uEntry.Properties["LockOutTime"].Value = 0; //unlock account

                uEntry.CommitChanges(); //may not be needed but adding it anyways

                uEntry.Close();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //DoSomethingWith --> E.Message.ToString();

            }
        }

        /// <summary>
        /// Reset Password to Account
        /// </summary>
        /// <param name="userDn"></param>
        /// <param name="password"></param>
        public void ResetPassword(string userDn, string password)
        {
            DirectoryEntry uEntry = new DirectoryEntry(userDn);
            uEntry.Invoke("SetPassword", new object[] { password });
            uEntry.Properties["LockOutTime"].Value = 0; //unlock account

            uEntry.Close();
        }





    }
}

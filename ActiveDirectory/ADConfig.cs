using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.Configuration;

namespace ActiveDirectoryHelper
{
    internal static class ADConfig
    {



        public static String LDAPPath
        {
            get
            {
                return ConfigurationManager.AppSettings["LDAPPath"];
            }
        }

        public static String LDAPUser
        {
            get
            {
                return ConfigurationManager.AppSettings["LDAPUser"];
            }
        }

        public static String LDAPPassword
        {
            get
            {
                return ConfigurationManager.AppSettings["LDAPPassword"];
            }
        }

        public static String LDAPDomain
        {
            get
            {
                return ConfigurationManager.AppSettings["LDAPDomain"];
            }
        }


    }
}

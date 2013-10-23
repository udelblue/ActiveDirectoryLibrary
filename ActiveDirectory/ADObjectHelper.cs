using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.Collections;


namespace ActiveDirectoryHelper
{
    public class ADObjectHelper
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
        /// Check for the Existence of an Object
        /// </summary>
        /// <param name="objectPath"></param>
        /// <returns></returns>
        public static bool Exists(string objectPath)
        {
            bool found = false;
            if (DirectoryEntry.Exists("LDAP://" + objectPath))
            {
                found = true;
            }
            return found;
        }

        /// <summary>
        /// Move an Object from one Location to Another
        /// </summary>
        /// <param name="objectLocation"></param>
        /// <param name="newLocation"></param>
        public void Move(string objectLocation, string newLocation)
        {

            DirectoryEntry eLocation = new DirectoryEntry("LDAP://" + objectLocation);
            DirectoryEntry nLocation = new DirectoryEntry("LDAP://" + newLocation);
            string newName = eLocation.Name;
            eLocation.MoveTo(nLocation, newName);
            nLocation.Close();
            eLocation.Close();
        }




        /// <summary>
        /// Enumerate Multi-String Attribute Values of an Object
        /// </summary>
        /// <param name="attributeName"></param>
        /// <param name="objectDn"></param>
        /// <param name="valuesCollection"></param>
        /// <param name="recursive"></param>
        /// <returns></returns>
        public ArrayList AttributeValuesMultiString(string attributeName,
     string objectDn, ArrayList valuesCollection, bool recursive)
        {
            DirectoryEntry ent = new DirectoryEntry(objectDn);
            PropertyValueCollection ValueCollection = ent.Properties[attributeName];
            IEnumerator en = ValueCollection.GetEnumerator();

            while (en.MoveNext())
            {
                if (en.Current != null)
                {
                    if (!valuesCollection.Contains(en.Current.ToString()))
                    {
                        valuesCollection.Add(en.Current.ToString());
                        if (recursive)
                        {
                            AttributeValuesMultiString(attributeName, "LDAP://" +
                            en.Current.ToString(), valuesCollection, true);
                        }
                    }
                }
            }
            ent.Close();
            ent.Dispose();
            return valuesCollection;
        }



        /// <summary>
        /// Enumerate Single String Attribute Values of an Object
        /// </summary>
        /// <param name="attributeName"></param>
        /// <param name="objectDn"></param>
        /// <returns></returns>
        public string AttributeValuesSingleString
    (string attributeName, string objectDn)
        {
            string strValue;
            DirectoryEntry ent = new DirectoryEntry(objectDn);
            strValue = ent.Properties[attributeName].Value.ToString();
            ent.Close();
            ent.Dispose();
            return strValue;
        }

        /// <summary>
        /// Enumerate an Object's Properties: The Ones with Values
        /// </summary>
        /// <param name="objectDn"></param>
        /// <returns></returns>
        public static ArrayList GetUsedAttributes(string objectDn)
        {
            DirectoryEntry objRootDSE = new DirectoryEntry("LDAP://" + objectDn);
            ArrayList props = new ArrayList();

            foreach (string strAttrName in objRootDSE.Properties.PropertyNames)
            {
                props.Add(strAttrName);
            }
            return props;
        }




        /// <summary>
        /// Create a New Security Group
        /// </summary>
        /// <param name="ouPath"></param>
        /// <param name="name"></param>
        public void Create(string ouPath, string name)
        {
            if (!DirectoryEntry.Exists("LDAP://CN=" + name + "," + ouPath))
            {
                try
                {
                    DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                    DirectoryEntry group = entry.Children.Add("CN=" + name, "group");
                    group.Properties["sAmAccountName"].Value = name;
                    group.CommitChanges();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message.ToString());
                }
            }
            else { Console.WriteLine(ouPath + " already exists"); }
        }


        /// <summary>
        /// Delete a group
        /// </summary>
        /// <param name="ouPath"></param>
        /// <param name="groupPath"></param>
        public void Delete(string ouPath, string groupPath)
        {
            if (DirectoryEntry.Exists("LDAP://" + groupPath))
            {
                try
                {
                    DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                    DirectoryEntry group = new DirectoryEntry("LDAP://" + groupPath);
                    entry.Children.Remove(group);
                    group.CommitChanges();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message.ToString());
                }
            }
            else
            {
                Console.WriteLine(path + " doesn't exist");
            }
        }


    }
}

using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace GSuiteGAL
{

    public class Address
    {
        public string name { get; set; }
        public string address { get; set; }

        public Address(string address, string name)
        {
            this.name = name;
            this.address = address;
        }
    }

    public class Changes
    {
        public List<Address> tochange;
        public List<Address> toadd;
        public List<Address> toremove;
    }

    public class ListMerger
    {
        Dictionary<string, Address> outlook;
        Dictionary<string, Address> google;

        public ListMerger()
        {
            outlook = new Dictionary<string, Address>();
            google = new Dictionary<string, Address>();
        }

        public void addOutlookAddress(string email, string name)
        {
            outlook.Add(email, new Address(email, name));
        }

        public void addOutlookAddress(Address addr)
        {
            outlook.Add(addr.address, addr);
        }

        public void addGoogleAddress(string email, string name)
        {
            google.Add(email, new Address(email, name));
        }
        public void addGoogleAddress(Address addr)
        {
            google.Add(addr.address, addr);
        }

        public Changes processLists()
        {
            List<Address> addToOutlook = new List<Address>();
            List<Address> modifyOutlook = new List<Address>();
            List<Address> delFromOutlook = new List<Address>();

            foreach (string email in google.Keys)
            {
                Address googleAddress, outlookAddress;
                google.TryGetValue(email, out googleAddress);

                // found in google but not found in outlook
                if (!outlook.TryGetValue(email, out outlookAddress))
                {
                    addToOutlook.Add(googleAddress);
                }
                else if (!outlookAddress.name.Equals(googleAddress.name))
                {
                    modifyOutlook.Add(googleAddress);
                }
            }

            foreach (string email in outlook.Keys)
            {
                Address outlookAddress;
                outlook.TryGetValue(email, out outlookAddress);

                if (!google.ContainsKey(email))
                    delFromOutlook.Add(outlookAddress);
            }


            Changes c = new Changes();
            c.toadd = addToOutlook;
            c.tochange = modifyOutlook;
            c.toremove = delFromOutlook;

            MessageBox.Show(String.Format("Syncing changes: +{0} <>{1} -{2}", addToOutlook.Count, modifyOutlook.Count, delFromOutlook.Count));
            return c;
        }
    }
}

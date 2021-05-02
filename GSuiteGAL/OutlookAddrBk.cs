using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GSuiteGAL
{
    public class OutlookAddrBk
    {
        Outlook.Folder addressBook;
        public List<Address> entries { get; } = new List<Address>();

        public OutlookAddrBk(Folder addressBook)
        {
            this.addressBook = addressBook;
        }

        public void retreiveEntries()
        {
            foreach (Outlook.ContactItem item in addressBook.Items)
            {
                Address a = new Address(item.Email1Address, item.FullName);
                entries.Add(a);
            }
        }

        public void RemoveEntries(List<Address> entries)
        {
            HashSet<string> emailsToRemove = new HashSet<string>();
            foreach (Address a in entries)
            {
                emailsToRemove.Add(a.address.ToLower());
            }

            List<Outlook.ContactItem> itemsToRemove = new List<Outlook.ContactItem>();
            foreach (Outlook.ContactItem item in addressBook.Items)
            {
                if (emailsToRemove.Contains(item.Email1Address))
                    item.Delete();

            }

        } // end remove entries

        public void AddEntries(List<Address> entries)
        {
            foreach (Address entry in entries)
            {
                Outlook.ContactItem contact = addressBook.Items.Add();
                contact.FullName = entry.name;
                contact.Email1Address = entry.address;
                contact.Save();

            }
        }

        public void ChangeEntries(List<Address> entries)
        {
            Dictionary<string, string> changes = new Dictionary<string, string>();

            foreach (Address a in entries)
            {
                changes.Add(a.address.ToLower(), a.name);
            }

            List<Outlook.ContactItem> itemsToRemove = new List<Outlook.ContactItem>();
            foreach (Outlook.ContactItem item in addressBook.Items)
            {
                if (changes.ContainsKey(item.Email1Address.ToLower()))
                {
                    String newName = "";
                    changes.TryGetValue(item.Email1Address.ToLower(), out newName);
                    item.FullName = newName;
                    item.Save();
                }
            } // end foreach
        } // end ChangeEntries

    } // end class
} // end namespace

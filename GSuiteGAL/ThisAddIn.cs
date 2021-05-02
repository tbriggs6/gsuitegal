using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GSuiteGAL
{
    public partial class ThisAddIn
    {
        // get the inspector object
        Outlook.Application application;
        Outlook.AddressLists addrlists;
        Outlook.Folder addressBook;

        /**
         * This does the heavy-lifting for the program.
         */
        private void Run()
        {
            if (!isCustomDefined())
            {
                this.addressBook = addAddressBook();
            }
            else
            {
                this.addressBook = getAddressBook();
            }

            // enumerate GSuite Addresses
            GSuiteDirectory gdir = new GSuiteDirectory();
            gdir.retrieveAddresses();

            // enumerate Outlook Addresses
            OutlookAddrBk odir = new OutlookAddrBk(addressBook);
            odir.retreiveEntries();

            // prepare to merge
            ListMerger merger = new ListMerger();
            foreach (Address ga in gdir.entries)
            {
                merger.addGoogleAddress(ga);
            }
            foreach (Address oa in odir.entries)
            {
                merger.addOutlookAddress(oa);
            }

            // merge
            Changes chchanges = merger.processLists();

            // make changes
            odir.RemoveEntries(chchanges.toremove);
            odir.ChangeEntries(chchanges.tochange);
            odir.AddEntries(chchanges.toadd);

            Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
            Config cfg = new Config();
            cfg.SetLastSync(unixTimestamp);
        }

        private Outlook.Folder addAddressBook()
        {
            Outlook.Folder contacts = this.application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts) as Outlook.Folder;
            Outlook.Folder addressBook = contacts.Folders.Add("GSuite GAL", Outlook.OlDefaultFolders.olFolderContacts) as Outlook.Folder;
            addressBook.ShowAsOutlookAB = true; // force display in Outlook Address Book

            return addressBook;
        }

        Boolean isCustomDefined()
        {
            addrlists = application.Session.AddressLists;
            foreach (Outlook.AddressList addrList in addrlists)
            {
                if (addrList.Name.Equals("GSuite GAL"))
                    return true;
            }
            return false;

        }

        Outlook.Folder getAddressBook()
        {
            Outlook.Folder contacts = this.application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts) as Outlook.Folder;
            foreach (Outlook.Folder folder in contacts.Folders)
            {
                if (folder.Name.Equals("GSuite GAL"))
                    return folder;
            }
            throw new SystemException("Not found.");
        }

        /* ------------------------------------------------------------------------------ */

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // get the application object
            application = this.Application;
            Config cfg = new Config();
            Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
            cfg.SetLastSync(unixTimestamp);

            if ((unixTimestamp < cfg.lastSync) || ((unixTimestamp - cfg.lastSync) > cfg.syncPeriod))
            {
                try
                {
                    Run();
                }
                catch (AggregateException ex)
                {
                    Console.WriteLine(ex);
                }
            }
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

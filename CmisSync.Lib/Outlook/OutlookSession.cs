using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CmisSync.Lib.Outlook
{
    public class OutlookSession
    {

        private Application application;
        private NameSpace nameSpace;
        private MAPIFolder defaultFolder;

        public Application Application 
        {
            get
            {
                return application;
            }
        }

        public NameSpace NameSpace
        {
            get
            {
                return nameSpace;
            }
        }

        public MAPIFolder DefaultFolder
        {
            get
            {
                return defaultFolder;
            }
        }
         
        public OutlookSession()
        {
            application = OutlookService.Instance.getApplication();
            nameSpace = OutlookService.Instance.getNameSpace(application);
            defaultFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
        }

        public MAPIFolder getFolderFromID(string entryID)
        {
            return nameSpace.GetFolderFromID(entryID);
        }

        public string getDefaultStoreID()
        {
            return nameSpace.DefaultStore.StoreID;
        }
    }
}

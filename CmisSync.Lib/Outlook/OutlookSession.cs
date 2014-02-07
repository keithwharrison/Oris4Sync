using log4net;
using Microsoft.Office.Interop.Outlook;
using stdole;
using System.Collections.Generic;

namespace CmisSync.Lib.Outlook
{
    public class OutlookSession
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(OutlookSession));

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

        public List<OutlookFolder> getFolderTree()
        {
            List<OutlookFolder> root = new List<OutlookFolder>();
            Folders folders = nameSpace.Folders;
            fillFolderTree(root, folders);
            return root;
        }

        private void fillFolderTree(List<OutlookFolder> folderList, Folders folders)
        {
            foreach (Folder folder in folders)
            {
                OutlookFolder outlookFolder = new OutlookFolder()
                {
                    entryId = folder.EntryID,
                    name = folder.Name,
                };

                Folders children = folder.Folders;
                fillFolderTree(outlookFolder.children, children);

                folderList.Add(outlookFolder);
            }
        }

    }
}

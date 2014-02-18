using log4net;
using Microsoft.Office.Interop.Outlook;
using stdole;
using System;
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
            application = OutlookService.getApplication();
            nameSpace = OutlookService.getNameSpace(application);
            defaultFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
        }

        public void close()
        {
            application = null;
            nameSpace = null;
            defaultFolder = null;
        }

        public void sendAndRecieve()
        {
            OutlookService.sendAndRecieve(nameSpace);
        }

        public MAPIFolder getFolderFromID(string entryID)
        {
            return nameSpace.GetFolderFromID(entryID);
        }

        public string getDefaultStoreID()
        {
            return Utils.Sha256Data(nameSpace.DefaultStore.StoreID);
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
                    folderPath = folder.FolderPath,
                };

                Folders children = folder.Folders;
                fillFolderTree(outlookFolder.children, children);

                folderList.Add(outlookFolder);
            }
        }

        public MAPIFolder getFolderByPath(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                return null;
            }

            string[] pathElements = folderPath.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            int currentElement = 0;
            Folder currentFolder = null;
            while (currentElement < pathElements.Length)
            {
                string pathElement = pathElements[currentElement];
                Folders folders = currentFolder != null ? currentFolder.Folders : nameSpace.Folders;
                Folder foundFolder = null;
                foreach (Folder folder in folders)
                {
                    if (folder.Name.Equals(pathElement))
                    {
                        foundFolder = folder;
                        break;
                    }
                }

                if (foundFolder != null)
                {
                    currentFolder = foundFolder;
                    currentElement++;
                }
                else 
                {
                    break;
                }
            }

            return (currentFolder != null && folderPath.Equals(currentFolder.FolderPath)) ?
                currentFolder : null;
        }

        public Email getEmail(MAPIFolder folder, MailItem mailItem)
        {
            return OutlookService.getEmail(folder, mailItem);
        }
        
        public List<EmailAttachment> getEmailAttachments(MailItem mailItem, Email email)
        {
            return OutlookService.getEmailAttachments(mailItem, email);
        }

        public EmailAttachment getEmailAttachment(Attachment attachment, Email email)
        {
            return OutlookService.getEmailAttachment(attachment, email);
        }
    }
}

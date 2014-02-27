using AddinExpress.Outlook;
using log4net;
using Microsoft.Office.Interop.Outlook;
using stdole;
using System;
using System.Collections.Generic;

namespace CmisSync.Lib.Outlook
{
    public class OutlookSession : IDisposable
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(OutlookSession));

        private readonly Application application;
        private readonly NameSpace nameSpace;
        private readonly MAPIFolder defaultFolder;
        private SecurityManager securityManager;
        private bool disposed = false;

        private SecurityManager SecurityManager
        {
            get
            {
                if (securityManager == null)
                {
                    securityManager = new SecurityManager();
                    securityManager.ConnectTo(application);
                    OutlookService.checkSecurityManager(securityManager);
                }
                return securityManager;
            }
        }

        public OutlookSession()
        {
            application = OutlookService.getApplication();
            nameSpace = OutlookService.getNameSpace(application);
            defaultFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
        }

        /// <summary>
        /// Destructor.
        /// </summary>
        ~OutlookSession()
        {
            Dispose(false);
        }


        /// <summary>
        /// Implement IDisposable interface. 
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Implement IDisposable interface. 
        /// </summary>
        public void Close()
        {
            Dispose();
        }


        /// <summary>
        /// Dispose pattern implementation.
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (securityManager != null)
                    {
                        if (securityManager.DisableOOMWarnings)
                        {
                            //Ensure OOM warnings are enabled at end of session.
                            Logger.Warn("Security Manager OOM Warnings Left Disabled at end of session");
                            securityManager.DisableOOMWarnings = false;
                        }
                        securityManager.Disconnect(application);
                        securityManager.Dispose();
                    }
                }
                this.disposed = true;
            }
        }

        public void sendAndRecieve()
        {
            if (disposed)
            {
                throw new ObjectDisposedException(typeof(OutlookSession).Name);
            }
            OutlookService.sendAndRecieve(nameSpace);
        }

        public MAPIFolder getFolderFromID(string entryID)
        {
            if (disposed)
            {
                throw new ObjectDisposedException(typeof(OutlookSession).Name);
            }
            if (disposed) throw new ObjectDisposedException(typeof(OutlookSession).Name);
            return nameSpace.GetFolderFromID(entryID);
        }

        public List<OutlookFolder> getFolderTree()
        {
            if (disposed)
            {
                throw new ObjectDisposedException(typeof(OutlookSession).Name);
            }
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
            if (disposed)
            {
                throw new ObjectDisposedException(typeof(OutlookSession).Name);
            }
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
            if (disposed)
            {
                throw new ObjectDisposedException(typeof(OutlookSession).Name);
            }
            return OutlookService.getEmail(SecurityManager, folder, mailItem);
        }
        
        public List<EmailAttachment> getEmailAttachments(MailItem mailItem, Email email)
        {
            if (disposed)
            {
                throw new ObjectDisposedException(typeof(OutlookSession).Name);
            }
            return OutlookService.getEmailAttachments(SecurityManager, mailItem, email);
        }

        public EmailAttachment getEmailAttachmentWithTempFile(EmailAttachment emailAttachment)
        {
            if (disposed)
            {
                throw new ObjectDisposedException(typeof(OutlookSession).Name);
            }
            return OutlookService.getEmailAttachmentWithTempFile(SecurityManager, emailAttachment);
        }
    }
}

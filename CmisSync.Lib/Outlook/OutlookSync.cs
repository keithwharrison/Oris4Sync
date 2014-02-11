using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;

namespace CmisSync.Lib.Outlook
{
    public class OutlookSync
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(OutlookSync));

        private RepoInfo repoInfo;
        private OutlookDatabase outlookDatabase;
        private string repoUrl;

        public OutlookSync(RepoInfo repoInfo)
        {
            this.repoInfo = repoInfo;

            //Database
            string dataPath = repoInfo.CmisDatabase;
            this.outlookDatabase = new OutlookDatabase(Path.Combine(Path.GetDirectoryName(dataPath),
                Path.GetFileNameWithoutExtension(dataPath) + " (outlook plugin)" +
                Path.GetExtension(dataPath)));

            //Url
            repoUrl = repoInfo.Address.GetLeftPart(UriPartial.Authority);
        }

        public void Sync()
        {
            OutlookSession outlookSession = new OutlookSession();
            Oris4RestSession restSession = new Oris4RestSession(repoUrl);

            restSession.login(repoInfo.User, repoInfo.Password.ToString());

            string defaultStoreId = outlookSession.getDefaultStoreID();

            string registeredClient = restSession.getRegisteredClient();
            Logger.Info("Client: " + registeredClient);

            if (!registeredClient.Equals(defaultStoreId))
            {
                restSession.putRegisteredClient(registeredClient);
            }

            string[] folderPaths = repoInfo.getOutlookFolders();

            foreach (string folderPath in folderPaths)
            {
                MAPIFolder pickedFolder = outlookSession.getFolderByPath(folderPath);
                if (pickedFolder == null)
                {
                    Logger.ErrorFormat("Could not find selected outlook folder: {0}", folderPath);
                    continue;
                }

                Logger.Info("Entry ID: " + pickedFolder.EntryID);
                Logger.Info("Folder Name: " + pickedFolder.Name);
                Logger.Info("Folder Path: " + pickedFolder.FolderPath);

                List<Email> emailList = new List<Email>();

                Items items = pickedFolder.Items;
                foreach (object item in items)
                {
                    if (item is MailItem)
                    {
                        MailItem mailItem = (MailItem)item;
                        emailList.Add(OutlookService.Instance.getEmail(pickedFolder, mailItem));

                        Attachments attachments = mailItem.Attachments;
                        if (attachments.Count > 0)
                        {
                            foreach (Attachment attachment in attachments)
                            {
                                string tempFilePath = OutlookService.Instance.saveAttachmentToTempFile(attachment);
                                string dataHash = Utils.Sha256File(tempFilePath);
                                Logger.InfoFormat("Attachment: {0} {1}", tempFilePath, dataHash);
                                File.Delete(tempFilePath);
                            }
                        }
                    }
                }

                List<Email> returned = restSession.insertEmail(registeredClient, "keithharrison@oris4.com", emailList);
            }
        }
    }
}

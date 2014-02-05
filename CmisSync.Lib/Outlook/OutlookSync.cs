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

        private OutlookDatabase outlookDatabase;

        public OutlookSync(string dataPath)
        {
            Logger.Info("Constructor...");
            this.outlookDatabase = new OutlookDatabase(Path.Combine(Path.GetDirectoryName(dataPath),
                Path.GetFileNameWithoutExtension(dataPath) + " (outlook plugin)" +
                Path.GetExtension(dataPath)));
        }

        public void Sync()
        {
            OutlookSession outlookSession = new OutlookSession();
            Oris4RestSession restSession = new Oris4RestSession(Config.Instance.TestUrl);


            MAPIFolder pickedFolder = outlookSession.getFolderFromID("00000000DF515440C9C02E409C10AA4D4B9BD65582800000"); //Inbox

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


            restSession.login(Config.Instance.TestUsername, Config.Instance.TestPassword);

            string defaultStoreId = outlookSession.getDefaultStoreID();

            string registeredClient = restSession.getRegisteredClient();
            Logger.Info("Client: " + registeredClient);

            if (!registeredClient.Equals(defaultStoreId))
            {
                restSession.putRegisteredClient(registeredClient);
            }

            List<Email> returned = restSession.insertEmail(registeredClient, "keithharrison@oris4.com", emailList);
        }
    }
}

using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace CmisSync.Lib.Outlook
{
    public class OutlookPlugin
    {
        protected static readonly ILog Logger = LogManager.GetLogger(typeof(OutlookPlugin));

        private Database database;

        public OutlookPlugin(string dataPath)
        {
            Logger.Info("Constructor...");
            this.database = new Database(Path.Combine(Path.GetDirectoryName(dataPath),
                Path.GetFileNameWithoutExtension(dataPath) + " (outlook plugin)" +
                Path.GetExtension(dataPath)));
        }

        public void Sync()
        {
            Application application = OutlookUtils.Instance.getApplication();
            NameSpace nameSpace = OutlookUtils.Instance.getNameSpace(application);
            MAPIFolder defaultFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            MAPIFolder pickedFolder = nameSpace.PickFolder();

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
                    Logger.Info("Mail Item: " + mailItem.Subject);

                    PropertyAccessor propertyAccessor = mailItem.PropertyAccessor;

                    Email email = new Email()
                    {
                        messageID = (string)propertyAccessor.GetProperty(OutlookUtils.PR_INTERNET_MESSAGE_ID_W),
                        receivedDate = mailItem.ReceivedTime,
                        sentDate = mailItem.SentOn,
                        subject = mailItem.Subject,
                        folderPath = pickedFolder.FolderPath,
                    };

                    string inReplyTo = (string)propertyAccessor.GetProperty(OutlookUtils.PR_IN_REPLY_TO_ID_W);
                    if (!string.IsNullOrWhiteSpace(inReplyTo))
                    {
                        email.inReplyTo = inReplyTo;
                    }

                    //string references = propertyAccessor.GetProperty(OutlookUtils.PR_INTERNET_REFERENCES_W); //TODO:: References
                    //if (!string.IsNullOrWhiteSpace(references))
                    //{
                    //    email.references = references;
                    //}

                    switch (mailItem.BodyFormat)
                    {
                        case OlBodyFormat.olFormatHTML:
                            email.body = mailItem.HTMLBody;
                            break;
                        case OlBodyFormat.olFormatRichText:
                            email.body = (string)mailItem.RTFBody;
                            break;
                        case OlBodyFormat.olFormatPlain:
                        case OlBodyFormat.olFormatUnspecified:
                        default:
                            email.body = mailItem.Body;
                            break;
                    }

                    List<EmailContact> contacts = new List<EmailContact>();
                    Recipients recipients = mailItem.Recipients;
                    foreach (Recipient recipient in recipients)
                    {
                        String emailContactType = null;
                        switch (recipient.Type)
                        {
                            case (int)OlMailRecipientType.olOriginator:
                                emailContactType = "From";
                                break;
                            case (int)OlMailRecipientType.olTo:
                                emailContactType = "To";
                                break;
                            case (int)OlMailRecipientType.olCC:
                                emailContactType = "Cc";
                                break;
                            case (int)OlMailRecipientType.olBCC:
                                emailContactType = "Bcc";
                                break;
                        }

                        if (emailContactType != null)
                        {
                            contacts.Add(new EmailContact()
                            {
                                emailAddress = recipient.Address,
                                emailContactType = emailContactType,
                            });
                        }

                        Marshal.ReleaseComObject(recipient);
                    }

                    contacts.Add(new EmailContact()
                    {
                        emailAddress = mailItem.SenderEmailAddress,
                        emailContactType = "From",
                    });

                    email.emailContacts = contacts;
                    emailList.Add(email);

                    Logger.InfoFormat("DataHash: {0}", OutlookUtils.Instance.createEmailDataHash(email));

                    foreach (Microsoft.Office.Interop.Outlook.Attachment attachment in mailItem.Attachments)
                    {
                        Logger.InfoFormat("DisplayName: {0}", attachment.DisplayName);
                        Logger.InfoFormat("FileName: {0}", attachment.FileName);
                        Logger.InfoFormat("Index: {0}", attachment.Index);
                        Logger.InfoFormat("PathName: {0}", attachment.PathName);
                        Logger.InfoFormat("Position: {0}", attachment.Position);
                        Logger.InfoFormat("Size: {0}", attachment.Size);
                        Logger.InfoFormat("Type: {0}", attachment.Type);
                        //attachment.SaveAsFile(@"C:\Users\keith\Downloads\" + attachment.FileName);
                    }

                    Marshal.ReleaseComObject(propertyAccessor);
                }

                Marshal.ReleaseComObject(item);
            }

            OAuth oAuth = Oris4RestService.login(Config.Instance.TestUrl, Config.Instance.TestUsername, Config.Instance.TestPassword);

            Store defaultStore = nameSpace.DefaultStore;
            //WebServices.Instance.putRegisteredClient(Config.Instance.TestUrl, oAuth.value, oAuth.tokenType, defaultStore.StoreID);

            string registeredClient = Oris4RestService.getRegisteredClient(Config.Instance.TestUrl, oAuth.value, oAuth.tokenType);
            Logger.Info("Client: " + registeredClient);

            List<Email> returned = Oris4RestService.insertEmail(Config.Instance.TestUrl, oAuth.value, oAuth.tokenType, "Blah", "keithharrison@oris4.com", emailList);

            Marshal.ReleaseComObject(defaultStore);
            Marshal.ReleaseComObject(pickedFolder);
            Marshal.ReleaseComObject(defaultFolder);
            Marshal.ReleaseComObject(nameSpace);
            Marshal.ReleaseComObject(application);

        }
    }
}

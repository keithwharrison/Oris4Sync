using log4net;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace CmisSync.Lib.Outlook
{
    public class OutlookUtils
    {
        protected static readonly ILog Logger = LogManager.GetLogger(typeof(OutlookUtils));

        public static string PR_INTERNET_MESSAGE_ID_W = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
        public static string PR_IN_REPLY_TO_ID_W = "http://schemas.microsoft.com/mapi/proptag/0x1042001F";
        public static string PR_INTERNET_REFERENCES_W = "http://schemas.microsoft.com/mapi/proptag/0x1039001F";

        private static OutlookUtils instance;

        public static OutlookUtils Instance
        {
            get
            {
                if (instance == null) instance = new OutlookUtils();
                return instance;
            }
        }

        private OutlookUtils()
        {
        }

        public Application getApplication()
        {
            return new Application();
        }

        public NameSpace getNameSpace(Application application)
        {
            return application.GetNamespace("MAPI");
        }

        public string getOutlookVersionString()
        {
            return (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\Outlook.Application\CurVer", null, null);
        }

        public string getOutlookVersionNumber()
        {
            string version = getOutlookVersionString();
            if (version == null)
            {
                return null;
            }
            else
            {
                if (version.StartsWith("Outlook.Application."))
                {
                    return version.Substring("Outlook.Application.".Length) + ".0";
                }
                else
                {
                    return null;
                }
            }
        }

        public bool checkForProfile()
        {
            string outlookVersionNumber = getOutlookVersionNumber();
            return outlookVersionNumber != null &&
                (Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles", null, "Key Exists") != null ||
                Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\" + outlookVersionNumber + @"\Outlook\Profiles", null, "Key Exists") != null);
        }

        public void doTest()
        {
            Logger.Info("********** TEST STARTED **********");

            Application application = null;
            NameSpace nameSpace = null;
            MAPIFolder defaultFolder = null;
            Folders folders = null;
            try
            {
                application = getApplication();
                nameSpace = getNameSpace(application);
                defaultFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                folders = nameSpace.Folders;

                foreach (MAPIFolder folder in folders)
                {
                    Folders subfolders = folder.Folders;
                    try
                    {
                        logFolder(folder);

                        foreach (MAPIFolder subfolder in subfolders)
                        {
                            Items items = subfolder.Items;
                            try
                            {

                                logFolder(subfolder);
                                if (items != null)
                                {

                                    foreach (object item in items)
                                    {
                                        try
                                        {
                                            if (item is MailItem)
                                            {
                                                logItem(subfolder, (MailItem)item);
                                            }
                                            else
                                            {
                                                Logger.Info("Item not a MailItem: " + item.GetType().ToString());
                                            }
                                        }
                                        finally
                                        {

                                            Marshal.ReleaseComObject(item);
                                        }
                                    }
                                    Logger.Info("--------------------");
                                }
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(items);
                                Marshal.ReleaseComObject(subfolder);
                            }
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(subfolders);
                        Marshal.ReleaseComObject(folder);
                    }
                }

            }
            catch (System.Exception e)
            {
                Logger.Error("An error occured: " + e.Message, e);
            }
            finally
            {
                Marshal.ReleaseComObject(folders);
                Marshal.ReleaseComObject(defaultFolder);
                Marshal.ReleaseComObject(nameSpace);
                Marshal.ReleaseComObject(application);
            }
            Logger.Info("********** TEST COMPLETED **********");
        }

        private void sendAndRecieve(Application application)
        {
            NameSpace nameSpace = null;
            SyncObjects syncObjects = null;
            try
            {
                nameSpace = getNameSpace(application);
                syncObjects = nameSpace.SyncObjects;
                foreach (SyncObject syncObject in syncObjects)
                {
                    if (syncObject != null)
                    {
                        try
                        {
                            syncObject.Start();
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(syncObject);
                        }
                    }
                }
            }
            catch (System.Exception e)
            {
                Logger.Error("Could not send/recieve outlook objects", e);
            }
            finally
            {
                if (syncObjects != null) Marshal.ReleaseComObject(syncObjects);
                if (nameSpace != null) Marshal.ReleaseComObject(nameSpace);
            }
        }

        private string itemToString(MAPIFolder folder, MailItem item)
        {
            StringBuilder sb = new StringBuilder();
            //sb.Append("Sender: ").Append(item.Sender).AppendLine();
            //sb.Append("To: ").Append(item.To).AppendLine();
            //sb.Append("Subject: ").Append(item.Subject).AppendLine();
            //sb.AppendLine("Body: ").AppendLine(item.Body).AppendLine();
            //sb.Append("AttachmentCount: ").Append(item.Attachments.Count).AppendLine();
            //sb.AppendLine();
            sb.Append("Folder: ").Append(folder.FolderPath).Append("\\").Append(item.Subject);
            return sb.ToString();
        }

        private void logFolder(MAPIFolder folder)
        {
            StringBuilder sb = new StringBuilder();
            //sb.AppendLine("Folder EntryID:").AppendLine(folder.EntryID).AppendLine();
            //sb.AppendLine("Folder StoreID:").AppendLine(folder.StoreID).AppendLine();
            //sb.AppendLine("Unread Item Count: " + folder.UnReadItemCount);
            //sb.AppendLine("Default MessageClass: " + folder.DefaultMessageClass);
            //sb.AppendLine("Current View: " + folder.CurrentView.Name);
            //sb.AppendLine("Folder Path: " + folder.FolderPath);
            sb.Append("Folder: ").Append(folder.FolderPath).Append(" [").Append(folder.ShowItemCount).Append("]");
            //sb.AppendLine();
            Logger.Info(sb.ToString());
        }

        private void logItem(MAPIFolder folder, MailItem item)
        {
            StringBuilder sb = new StringBuilder();
            //sb.Append("Sender: ").Append(item.Sender).AppendLine();
            //sb.Append("To: ").Append(item.To).AppendLine();
            //sb.Append("Subject: ").Append(item.Subject).AppendLine();
            //sb.AppendLine("Body: ").AppendLine(item.Body).AppendLine();
            //sb.Append("AttachmentCount: ").Append(item.Attachments.Count).AppendLine();
            //sb.AppendLine();
            sb.Append("Folder: ").Append(folder.FolderPath).Append("\\").Append(item.Subject);
            Logger.Info(sb.ToString());
        }

        private long toJavaDate(DateTime date)
        {
            return (long)(date.ToUniversalTime() - new DateTime(1970, 1, 1)).TotalMilliseconds;
        }

        public string createEmailDataHash(Email email)
        {
            string dataToHash = createEmailHashString(email);
            string dataHash = sha256(dataToHash);
            Logger.InfoFormat("DataHash: {0} {1}", dataToHash, dataHash);
            return dataHash;
        }

        private string createEmailHashString(Email email)
        {
            long sentDate = toJavaDate(email.sentDate);
            string messageId = email.messageID;

            string sender = null;
            List<string> recipientList = new List<string>(3);

            foreach (EmailContact emailContact in email.emailContacts)
            {
                if ("From".Equals(emailContact.emailContactType))
                {
                    sender = emailContact.emailAddress;
                }
                else
                {
                    string emailAddress = emailContact.emailAddress.ToUpper();
                    if (!recipientList.Contains(emailAddress))
                    {
                        recipientList.Add(emailAddress);
                    }
                }
            }

            return createEmailHashString(sentDate, messageId, sender, recipientList);
        }

        /// <summary>
        /// Create email hash string.
        /// </summary>
        private string createEmailHashString(long sendDate, string messageId, string sender, List<string> recipientList)
        {
            StringBuilder dataToHash = new StringBuilder(150);
            dataToHash.Append(sendDate);
            dataToHash.Append(messageId);
            dataToHash.Append(sender != null ? sender.ToUpper() : string.Empty);

            List<string> uppercaseRecipientList = new List<string>(recipientList.Count);

            foreach (string recipient in recipientList)
            {
                uppercaseRecipientList.Add(recipient.ToUpper());
            }

            uppercaseRecipientList.Sort();

            foreach (string uppercaseRecipient in uppercaseRecipientList)
            {
                dataToHash.Append(uppercaseRecipient);
            }

            return dataToHash.ToString();
        }

        private string sha256(string data)
        {
            SHA256Managed crypt = new SHA256Managed();
            string hash = String.Empty;
            byte[] crypto = crypt.ComputeHash(Encoding.UTF8.GetBytes(data), 0, Encoding.UTF8.GetByteCount(data));
            foreach (byte bit in crypto)
            {
                hash += bit.ToString("x2");
            }
            return hash;
        }
    }
}

using AddinExpress.Outlook;
using log4net;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CmisSync.Lib.Outlook
{
    public class OutlookService
    {
        protected static readonly ILog Logger = LogManager.GetLogger(typeof(OutlookService));

        public static string PR_INTERNET_MESSAGE_ID_W = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
        public static string PR_IN_REPLY_TO_ID_W = "http://schemas.microsoft.com/mapi/proptag/0x1042001F";
        public static string PR_INTERNET_REFERENCES_W = "http://schemas.microsoft.com/mapi/proptag/0x1039001F";

        public static Application getApplication()
        {
            return new Application();
        }

        public static NameSpace getNameSpace(Application application)
        {
            return application.GetNamespace("MAPI");
        }

        public static string getOutlookVersionString()
        {
            return (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\Outlook.Application\CurVer", null, null);
        }

        public static string getOutlookVersionNumber()
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

        public static bool checkForOutlookInstallation()
        {
            return getOutlookVersionString() != null;
        }

        public static bool checkForProfile()
        {
            string outlookVersionNumber = getOutlookVersionNumber();
            return outlookVersionNumber != null &&
                (Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles", null, "Key Exists") != null ||
                Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\" + outlookVersionNumber + @"\Outlook\Profiles", null, "Key Exists") != null);
        }

        public static void sendAndRecieve(NameSpace nameSpace)
        {
            try
            {
                SyncObjects syncObjects = nameSpace.SyncObjects;
                foreach (SyncObject syncObject in syncObjects)
                {
                    if (syncObject != null)
                    {
                        syncObject.Start();
                    }
                }
            }
            catch (System.Exception e)
            {
                Logger.Error("Could not send/recieve outlook objects", e);
            }
        }

        public static Email getEmail(MAPIFolder folder, MailItem mailItem)
        {
            Logger.Info("Mail Item: " + mailItem.Subject);

            Email email = new Email()
            {
                messageID = getMessageId(mailItem),
                receivedDate = mailItem.ReceivedTime,
                sentDate = mailItem.SentOn,
                subject = mailItem.Subject,
                folderPath = folder.FolderPath,
                inReplyTo = getInReplyTo(mailItem),
                references = getReferences(mailItem),
                body = getBody(mailItem),
                emailContacts = getEmailContacts(mailItem),
                entryID = mailItem.EntryID,
            };

            email.dataHash = createEmailDataHash(email);

            return email;
        }

        public static List<EmailAttachment> getEmailAttachments(MailItem mailItem, Email email)
        {
            List<EmailAttachment> emailAttachments = new List<EmailAttachment>();
            Attachments attachments = mailItem.Attachments;
            foreach (Attachment attachment in attachments)
            {
                emailAttachments.Add(getEmailAttachment(attachment, email));
            }
            return emailAttachments;
        }

        public static EmailAttachment getEmailAttachment(Attachment attachment, Email email)
        {
            string tempFilePath = saveAttachmentToTempFile(attachment);
            string dataHash = Utils.Md5File(tempFilePath);
            Logger.DebugFormat("Attachment: {0} {1}", tempFilePath, dataHash);

            return new EmailAttachment()
            {
                emailDataHash = email.dataHash,
                dataHash = dataHash,
                fileName = attachment.DisplayName,
                name = attachment.DisplayName,
                fileSize = attachment.Size,
                folderPath = email.folderPath,
                tempFilePath = tempFilePath,
            };
        }

        public static string saveAttachmentToTempFile(Attachment attachment)
        {
            string tempFilePath = Path.GetTempFileName();

            SecurityManager securityManager = new SecurityManager();
            try
            {
                securityManager.DisableOOMWarnings = true;
                attachment.SaveAsFile(tempFilePath);
            }
            finally
            {
                securityManager.DisableOOMWarnings = false;
            }
            return tempFilePath;
        }

        private static string getMessageId(MailItem mailItem)
        {
            SecurityManager securityManager = new SecurityManager();
            try
            {
                securityManager.DisableOOMWarnings = true;
                PropertyAccessor propertyAccessor = mailItem.PropertyAccessor;
                return (string)propertyAccessor.GetProperty(OutlookService.PR_INTERNET_MESSAGE_ID_W);
            }
            finally
            {
                securityManager.DisableOOMWarnings = false;
            }
        }

        private static string getInReplyTo(MailItem mailItem)
        {
            SecurityManager securityManager = new SecurityManager();
            try
            {
                securityManager.DisableOOMWarnings = true;
                PropertyAccessor propertyAccessor = mailItem.PropertyAccessor;
                string inReplyTo = (string)propertyAccessor.GetProperty(OutlookService.PR_IN_REPLY_TO_ID_W);
                if (!string.IsNullOrWhiteSpace(inReplyTo))
                {
                    return inReplyTo;
                }
                return null;
            }
            finally
            {
                securityManager.DisableOOMWarnings = false;
            }
        }

        private static string getReferences(MailItem mailItem)
        {
            /*
            SecurityManager securityManager = new SecurityManager();
            try
            {
                securityManager.DisableOOMWarnings = true;
                PropertyAccessor propertyAccessor = mailItem.PropertyAccessor;
                string references = propertyAccessor.GetProperty(OutlookService.PR_INTERNET_REFERENCES_W); //TODO:: References
                if (!string.IsNullOrWhiteSpace(references))
                {
                    return references;
                }
            */
                return null;
            /*
            }
            finally
            {
                securityManager.DisableOOMWarnings = false;
            }
            */
        }

        private static string getBody(MailItem mailItem)
        {
            SecurityManager securityManager = new SecurityManager();
            try
            {
                securityManager.DisableOOMWarnings = true;
                switch (mailItem.BodyFormat)
                {
                    case OlBodyFormat.olFormatHTML:
                        return mailItem.HTMLBody;
                    case OlBodyFormat.olFormatRichText:
                    case OlBodyFormat.olFormatPlain:
                    case OlBodyFormat.olFormatUnspecified:
                    default:
                        return mailItem.Body;
                }
            }
            finally
            {
                securityManager.DisableOOMWarnings = false;
            }
        }

        private static List<EmailContact> getEmailContacts(MailItem mailItem)
        {
            SecurityManager securityManager = new SecurityManager();
            try
            {
                securityManager.DisableOOMWarnings = true;
                List<EmailContact> contacts = new List<EmailContact>();

                //Sender...
                contacts.Add(new EmailContact()
                {
                    emailAddress = mailItem.SenderEmailAddress,
                    emailContactType = "From",
                });

                //Recipients...
                Recipients recipients = mailItem.Recipients;
                foreach (Recipient recipient in recipients)
                {
                    contacts.Add(getEmailRecipient(recipient));
                }

                return contacts;
            }
            finally
            {
                securityManager.DisableOOMWarnings = false;
            }
        }

        private static EmailContact getEmailRecipient(Recipient recipient)
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
                default:
                    Logger.WarnFormat("Invalid recipient type -> {0} - defaulting to 'To'", recipient.Type);
                    emailContactType = "To";
                    break;
            }

            return new EmailContact()
            {
                emailAddress = recipient.Address,
                emailContactType = emailContactType,
            };
        }

        private static long toJavaDate(DateTime date)
        {
            return (long)(date.ToUniversalTime() - new DateTime(1970, 1, 1)).TotalMilliseconds;
        }

        public static string createEmailDataHash(Email email)
        {
            string dataToHash = createEmailHashString(email);
            string dataHash = Utils.Sha256Data(dataToHash);
            //Logger.InfoFormat("DataHash: {0} {1}", dataToHash, dataHash);
            return dataHash;
        }

        private static string createEmailHashString(Email email)
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
        private static string createEmailHashString(long sendDate, string messageId, string sender, List<string> recipientList)
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
    }
}

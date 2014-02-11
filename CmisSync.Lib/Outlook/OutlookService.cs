using log4net;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace CmisSync.Lib.Outlook
{
    public class OutlookService
    {
        protected static readonly ILog Logger = LogManager.GetLogger(typeof(OutlookService));

        public static string PR_INTERNET_MESSAGE_ID_W = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
        public static string PR_IN_REPLY_TO_ID_W = "http://schemas.microsoft.com/mapi/proptag/0x1042001F";
        public static string PR_INTERNET_REFERENCES_W = "http://schemas.microsoft.com/mapi/proptag/0x1039001F";

        private static OutlookService instance;

        public static OutlookService Instance
        {
            get
            {
                if (instance == null) instance = new OutlookService();
                return instance;
            }
        }

        private OutlookService()
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

        public bool checkForOutlookInstallation()
        {
            return getOutlookVersionString() != null;
        }

        public bool checkForProfile()
        {
            string outlookVersionNumber = getOutlookVersionNumber();
            return outlookVersionNumber != null &&
                (Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles", null, "Key Exists") != null ||
                Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\" + outlookVersionNumber + @"\Outlook\Profiles", null, "Key Exists") != null);
        }

        public void sendAndRecieve(NameSpace nameSpace)
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

        public Email getEmail(MAPIFolder folder, MailItem mailItem)
        {
            Logger.Info("Mail Item: " + mailItem.Subject);

            PropertyAccessor propertyAccessor = mailItem.PropertyAccessor;

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
            };

            email.dataHash = createEmailDataHash(email);

            return email;
        }

        public List<EmailAttachment> getEmailAttachments(MailItem mailItem, Email email)
        {
            List<EmailAttachment> emailAttachments = new List<EmailAttachment>();
            Attachments attachments = mailItem.Attachments;
            foreach (Attachment attachment in attachments)
            {
                emailAttachments.Add(getEmailAttachment(attachment, email));
            }
            return emailAttachments;
        }

        public EmailAttachment getEmailAttachment(Attachment attachment, Email email)
        {
            string tempFilePath = saveAttachmentToTempFile(attachment);
            string dataHash = Utils.Md5File(tempFilePath);
            Logger.DebugFormat("Attachment: {0} {1}", tempFilePath, dataHash);

            return new EmailAttachment()
            {
                emailDataHash = email.dataHash,
                dataHash = dataHash,
                fileName = attachment.DisplayName,
                fileSize = attachment.Size,
                folderPath = email.folderPath,
                tempFilePath = tempFilePath,
            };
        }

        public string saveAttachmentToTempFile(Attachment attachment)
        {
            string tempFilePath = Path.GetTempFileName();
            attachment.SaveAsFile(tempFilePath);
            return tempFilePath;
        }

        private string getMessageId(MailItem mailItem)
        {
            PropertyAccessor propertyAccessor = mailItem.PropertyAccessor;
            return (string)propertyAccessor.GetProperty(OutlookService.PR_INTERNET_MESSAGE_ID_W);
        }

        private string getInReplyTo(MailItem mailItem)
        {
            PropertyAccessor propertyAccessor = mailItem.PropertyAccessor;
            string inReplyTo = (string)propertyAccessor.GetProperty(OutlookService.PR_IN_REPLY_TO_ID_W);
            if (!string.IsNullOrWhiteSpace(inReplyTo))
            {
                return inReplyTo;
            }
            return null;
        }

        private string getReferences(MailItem mailItem)
        {
            //PropertyAccessor propertyAccessor = mailItem.PropertyAccessor;
            //string references = propertyAccessor.GetProperty(OutlookUtils.PR_INTERNET_REFERENCES_W); //TODO:: References
            //if (!string.IsNullOrWhiteSpace(references))
            //{
            //    return references;
            //}
            return null;
        }

        private string getBody(MailItem mailItem)
        {
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

        private List<EmailContact> getEmailContacts(MailItem mailItem)
        {
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

        private EmailContact getEmailRecipient(Recipient recipient)
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

        private long toJavaDate(DateTime date)
        {
            return (long)(date.ToUniversalTime() - new DateTime(1970, 1, 1)).TotalMilliseconds;
        }

        public string createEmailDataHash(Email email)
        {
            string dataToHash = createEmailHashString(email);
            string dataHash = Utils.Sha256Data(dataToHash);
            //Logger.InfoFormat("DataHash: {0} {1}", dataToHash, dataHash);
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
    }
}

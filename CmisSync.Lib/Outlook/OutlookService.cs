using AddinExpress.Outlook;
using CmisSync.Lib.Cmis;
using log4net;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace CmisSync.Lib.Outlook
{
    public static class OutlookService
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(OutlookService));

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

        public static bool isOutlookInstalled()
        {
            return getOutlookVersionString() != null;
        }

        public static bool isOutlookProfileAvailable()
        {
            string outlookVersionNumber = getOutlookVersionNumber();
            return outlookVersionNumber != null &&
                (Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles", null, "Key Exists") != null ||
                Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\" + outlookVersionNumber + @"\Outlook\Profiles", null, "Key Exists") != null);
        }

        public static bool isOutlookSecurityManagerBitnessMatch()
        {
            if (!isOutlookInstalled())
            {
                return false;
            }

            if (isOutlookSecurityManager32Bit())
            {
                if (isOutlook32Bit())
                {
                    return true;
                }
                else
                {
                    Logger.Error("Outlook security manager mismatch - Security Manager: 32bit, Outlook: 64bit.  Please re-install Oris4Sync.");
                    return false;
                }
            }
            else if (isOutlookSecurityManager64Bit())
            {
                if (isOutlook64Bit())
                {
                    return true;
                }
                else
                {
                    Logger.Error("Outlook security manager mismatch - Security Manager: 64bit, Outlook: 32bit.  Please re-install Oris4Sync.");
                    return false;
                }
            }
            else
            {
                Logger.Error("Outlook installed but no security manager was registered.  Please re-install Oris4Sync.");
                return false;
            }
        }

        public static bool isOutlook32Bit()
        {
            return !isOutlook64Bit();
        }

        public static bool isOutlook64Bit()
        {
            string outlookVersionNumber = getOutlookVersionNumber();
            if (outlookVersionNumber != null)
            {
                object bitnessObject = Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Office\" + outlookVersionNumber + @"\Outlook", "Bitness", null);
                if (bitnessObject != null)
                {
                    return "x64".Equals(bitnessObject.ToString()) ? true : false;
                }
                else
                {
                    Logger.WarnFormat("Could not find bitness for Outlook version {0}, defaulting to 32bit", outlookVersionNumber);
                    return false;
                }
            }
            else
            {
                throw new BaseException("Could not determine Outlook bitness: outlook not installed.");
            }
        }

        public static bool isOutlookSecurityManager32Bit()
        {
            return Registry.GetValue(@"HKEY_CLASSES_ROOT\AppID\secman.DLL", "AppID", null) != null;
        }

        public static bool isOutlookSecurityManager64Bit()
        {
            return Registry.GetValue(@"HKEY_CLASSES_ROOT\AppID\secman64.DLL", "AppID", null) != null;
        }

        public static void checkSecurityManager(SecurityManager securityManager)
        {
            switch (securityManager.Check(osmWarningKind.osmObjectModel))
            {
                case osmResult.osmOK:
                    Logger.Info("Outlook Security Manager: OK");
                    break;
                case osmResult.osmDLLNotLoaded:
                    throw new BaseException("Unable to load Outlook Security Manager: DLL Not Loaded");
                case osmResult.osmSecurityGuardNotFound:
                    throw new BaseException("Unable to load Outlook Security Manager: Security Guard Not Found");
                case osmResult.osmUnknownOlVersion:
                    throw new BaseException("Unable to load Outlook Security Manager: Unknown Outlook Version");
                case osmResult.osmCDONotFound:
                    throw new BaseException("Unable to load Outlook Security Manager: CDO Not Found");
                default:
                    throw new BaseException("Unable to load Outlook Security Manager: Reason Unknown");
            }
        }

        public static void sendAndRecieve(NameSpace nameSpace)
        {
            try
            {
                SyncObjects syncObjects = nameSpace.SyncObjects;
                foreach (SyncObject syncObject in syncObjects)
                {
                    syncObject.Start();
                }
            }
            catch (System.Exception e)
            {
                Logger.Error("Could not send/recieve outlook objects", e);
            }
        }

        public static Email getEmail(SecurityManager securityManager, MAPIFolder folder, MailItem mailItem)
        {
            Logger.DebugFormat("Mail Item: {0}", mailItem.Subject);

            Email email = new Email()
            {
                messageID = getMessageId(securityManager, mailItem),
                receivedDate = mailItem.ReceivedTime,
                sentDate = mailItem.SentOn,
                subject = mailItem.Subject,
                folderPath = folder.FolderPath,
                inReplyTo = getInReplyTo(securityManager, mailItem),
                references = getReferences(securityManager, mailItem),
                body = getBody(securityManager, mailItem),
                emailContacts = getEmailContacts(securityManager, mailItem),
                entryID = mailItem.EntryID,
            };

            email.dataHash = createEmailDataHash(email);

            return email;
        }

        public static List<EmailAttachment> getEmailAttachments(SecurityManager securityManager, MailItem mailItem, Email email)
        {
            try
            {
                securityManager.DisableOOMWarnings = true;
                List<EmailAttachment> emailAttachments = new List<EmailAttachment>();
                Attachments attachments = mailItem.Attachments;
                foreach (Attachment attachment in attachments)
                {
                    emailAttachments.Add(new EmailAttachment()
                    {
                        emailDataHash = email.dataHash,
                        fileName = attachment.DisplayName,
                        name = attachment.DisplayName,
                        fileSize = attachment.Size,
                        folderPath = email.folderPath,
                        attachment = attachment,
                    });
                }
                return emailAttachments;
            }
            finally
            {
                securityManager.DisableOOMWarnings = false;
            }
        }

        public static EmailAttachment getEmailAttachmentWithTempFile(SecurityManager securityManager, EmailAttachment emailAttachment)
        {
            emailAttachment.tempFilePath = saveAttachmentToTempFile(securityManager, emailAttachment.attachment);
            emailAttachment.dataHash = Utils.Md5File(emailAttachment.tempFilePath);
            Logger.DebugFormat("Attachment: {0} {1}", emailAttachment.tempFilePath, emailAttachment.dataHash);

            return emailAttachment;
        }

        public static string saveAttachmentToTempFile(SecurityManager securityManager, Attachment attachment)
        {
            string tempFilePath = Path.GetTempFileName();

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

        private static string getMessageId(SecurityManager securityManager, MailItem mailItem)
        {
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

        private static string getInReplyTo(SecurityManager securityManager, MailItem mailItem)
        {
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

        private static string getReferences(SecurityManager securityManager, MailItem mailItem)
        {
            /*
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

        private static string getBody(SecurityManager securityManager, MailItem mailItem)
        {
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

        private static List<EmailContact> getEmailContacts(SecurityManager securityManager, MailItem mailItem)
        {
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

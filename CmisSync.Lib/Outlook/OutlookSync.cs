using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace CmisSync.Lib.Outlook
{
    /// <summary>
    /// Outlook Syncronization plugin.
    /// </summary>
    public class OutlookSync : IDisposable
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(OutlookSync));

        private static readonly int EMAIL_BATCH_SIZE = 50;

        /// <summary>
        /// Track whether <c>Dispose</c> has been called.
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Repository info.
        /// </summary>
        private RepoInfo repoInfo;

        /// <summary>
        /// Database object.
        /// </summary>
        private OutlookDatabase outlookDatabase;
        
        /// <summary>
        /// Repository URL.
        /// </summary>
        private string repoUrl;

        /// <summary>
        /// Sleep while suspended method handles repo pause and repo cancel actions.
        /// </summary>
        public delegate void SleepWhileSuspendedDelegate();
        private event SleepWhileSuspendedDelegate SleepWhileSuspended;

        /// <summary>
        /// Sleep while suspended method handles repo pause and repo cancel actions.
        /// </summary>
        public delegate void ProcessRecoverableExceptionDelegate(string logMessage, System.Exception exception);
        private event ProcessRecoverableExceptionDelegate ProcessRecoverableException;

        /// <summary>
        /// Constructor.
        /// </summary>
        public OutlookSync(RepoInfo repoInfo, SleepWhileSuspendedDelegate SleepWhileSuspended,
            ProcessRecoverableExceptionDelegate ProcessRecoverableException)
        {
            this.repoInfo = repoInfo;

            this.SleepWhileSuspended = SleepWhileSuspended;
            this.ProcessRecoverableException = ProcessRecoverableException;

            //Database
            this.outlookDatabase = new OutlookDatabase(GetOutlookDatabasePath(repoInfo.CmisDatabase));

            //Url
            repoUrl = repoInfo.Address.GetLeftPart(UriPartial.Authority);
        }

        /// <summary>
        /// Update settings.
        /// </summary>
        public void UpdateSettings(RepoInfo repoInfo)
        {
            this.repoInfo = repoInfo;
        }


        /// <summary>
        /// Destructor.
        /// </summary>
        ~OutlookSync()
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
        /// Dispose pattern implementation.
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    this.outlookDatabase.Dispose();
                }
                this.disposed = true;
            }
        }

        /// <summary>
        /// Get the outlook database path.
        /// </summary>
        public static string GetOutlookDatabasePath(string cmisDatabasePath)
        {
            return Path.Combine(Path.GetDirectoryName(cmisDatabasePath),
                Path.GetFileNameWithoutExtension(cmisDatabasePath) + ".outlook");
        }


        /// <summary>
        /// Dispose pattern implementation.
        /// </summary>
        public void Sync(bool fullSync)
        {
            if (!repoInfo.OutlookEnabled ||
                !fullSync || 
                !OutlookService.isOutlookInstalled() ||
                !OutlookService.isOutlookProfileAvailable() || 
                !OutlookService.isOutlookSecurityManagerBitnessMatch())
            {
                return;
            }

            Logger.InfoFormat("Outlook Sync Started: {0}", repoInfo.Name);
            try
            {
                Oris4RestSession restSession = new Oris4RestSession(repoUrl);
                using (OutlookSession outlookSession = new OutlookSession())
                {
                    restSession.login(repoInfo.User, repoInfo.Password.ToString());

                    RegisterOutlookClient(restSession);

                    //Send and recieve emails...
                    //outlookSession.sendAndRecieve();

                    string[] folderPaths = repoInfo.getOutlookFolders();

                    foreach (string folderPath in folderPaths)
                    {
                        MAPIFolder folder = outlookSession.getFolderByPath(folderPath);
                        if (folder == null)
                        {
                            Logger.ErrorFormat("Could not find selected outlook folder: {0}", folderPath);
                            continue;
                        }

                        Logger.DebugFormat("Syncing Outlook Folder: {0}", folder.FolderPath);

                        HashSet<string> allEmailsInFolder = new HashSet<string>();
                        List<Email> emailList = new List<Email>();
                        List<EmailAttachment> attachmentList = new List<EmailAttachment>();

                        Items items = folder.Items;
                        foreach (object item in items)
                        {
                            if (item is MailItem)
                            {
                                SleepWhileSuspended();

                                MailItem mailItem = (MailItem)item;
                                string entryId = mailItem.EntryID;

                                if (EmailWorthSyncing(mailItem))
                                {
                                    allEmailsInFolder.Add(entryId);

                                    if (!outlookDatabase.ContainsEmail(folderPath, entryId))
                                    {
                                        Email email = outlookSession.getEmail(folderPath, mailItem);
                                        emailList.Add(email);
                                        attachmentList.AddRange(outlookSession.getEmailAttachments(folderPath, mailItem, email));
                                    }
                                    else
                                    {
                                        //Database already contains the email
                                        int attachmentCount = outlookSession.getEmailAttachmentCount(mailItem);
                                        if (attachmentCount > 0)
                                        {
                                            int databaseAttachmentCount = outlookDatabase.CountAttachments(folderPath, entryId);
                                            if (attachmentCount > databaseAttachmentCount)
                                            {
                                                //Email has more attachments than exist in database...
                                                Email email = outlookSession.getEmail(folderPath, mailItem);
                                                attachmentList.AddRange(outlookSession.getEmailAttachments(folderPath, mailItem, email));
                                            }
                                        }
                                    }
                                }

                                if (emailList.Count >= EMAIL_BATCH_SIZE)
                                {
                                    UploadEmails(restSession, emailList);
                                    UploadAttachments(restSession, outlookSession, attachmentList);
                                }
                            }
                        }

                        if (emailList.Count > 0)
                        {
                            UploadEmails(restSession, emailList);
                            UploadAttachments(restSession, outlookSession, attachmentList);
                        }

                        DeleteObsoleteEmailsFromFolder(restSession, folderPath, allEmailsInFolder);
                    }

                    DeleteObsoleteFolders(restSession, folderPaths);
                }
            }
            finally
            {
                Logger.InfoFormat("Outlook Sync Complete: {0}", repoInfo.Name);
            }
        }

        private void RegisterOutlookClient(Oris4RestSession restSession)
        {
            SleepWhileSuspended();
            //Client registration...
            string clientId = outlookDatabase.GetClientId();
            if (string.IsNullOrWhiteSpace(clientId))
            {
                clientId = Guid.NewGuid().ToString();
            }

            string registeredClient = restSession.getRegisteredClient();
            Logger.InfoFormat("Current registered Outlook client ID: {0}", registeredClient);
            if (!registeredClient.Equals(clientId))
            {
                SleepWhileSuspended();
                //TODO: Ask user if they are sure before putting a new client ID (all emails deleted)
                Logger.InfoFormat("Registering a new Outlook client ID: {0}", clientId);
                outlookDatabase.RemoveAllEmails();
                outlookDatabase.SetClientId(clientId);
                restSession.putRegisteredClient(clientId);
            }
        }

        private bool EmailWorthSyncing(MailItem mailItem)
        {

            return true;
        }

        private bool AttachmentWorthSyncing(EmailAttachment emailAttachment)
        {

            return true;
        }

        private void UploadEmails(Oris4RestSession restSession, List<Email> emails)
        {
            SleepWhileSuspended();

            try
            {
                Logger.InfoFormat("Uploading {0} emails.", emails.Count);
                Dictionary<string, long> emailKeyMap = restSession.insertEmailBatch(emails);

                foreach (Email email in emails)
                {
                    if (emailKeyMap.ContainsKey(email.dataHash))
                    {
                        long emailKey = emailKeyMap[email.dataHash];
                        outlookDatabase.AddEmail(email.folderPath, email.entryID, email.dataHash, emailKey, DateTime.Now);
                        Logger.InfoFormat("Added email to database: {0}\\{1}", email.folderPath, email.dataHash);
                    }
                    else
                    {
                        Logger.ErrorFormat("Email was not inserted: {0}\\{1}", email.folderPath, email.dataHash);
                    }
                }
            }
            catch (System.Exception e)
            {
                ProcessRecoverableException("Problem while uploading email batch", e);
            }
            finally
            {
                emails.Clear();
            }
        }

        private void DeleteObsoleteEmailsFromFolder(Oris4RestSession restSession, string folderPath, HashSet<string> emailsInFolder)
        {
            SleepWhileSuspended();

            try
            {
                HashSet<string> emailsInDatabase = outlookDatabase.ListEntryIds(folderPath);
                List<string> obsoleteEmails = new List<string>();
                foreach (string entryId in emailsInDatabase)
                {
                    if (!emailsInFolder.Contains(entryId))
                    {
                        obsoleteEmails.Add(entryId);
                    }
                }

                if (obsoleteEmails.Count > 0)
                {
                    Logger.InfoFormat("Deleting {0} emails on server from folder: {1}", obsoleteEmails.Count, folderPath);
                    DeleteEmails(restSession, folderPath, obsoleteEmails);
                }
            }
            catch (System.Exception e)
            {
                ProcessRecoverableException("Error deleting obsolete emails from folder", e);
            }
        }

        private void DeleteObsoleteFolders(Oris4RestSession restSession, string[] folderPaths)
        {
            SleepWhileSuspended();

            try 
            {

                HashSet<string> currentFolders = new HashSet<string>(folderPaths);
                HashSet<string> foldersInDatabase = outlookDatabase.ListDistinctFolders();
                foreach (string folderPath in foldersInDatabase)
                {
                    if (!currentFolders.Contains(folderPath))
                    {
                        //Folder is obsolete
                        List<string> obsoleteEmails = new List<string>(outlookDatabase.ListEntryIds(folderPath));
                        Logger.InfoFormat("Deleting {0} emails on server from folder: {1}", obsoleteEmails.Count, folderPath);
                        DeleteEmails(restSession, folderPath, obsoleteEmails);
                    }
                }
            }
            catch (System.Exception e)
            {
                ProcessRecoverableException("Could not delete email obsolete folders.", e);
            }
        }


        private void DeleteEmails(Oris4RestSession restSession, string folderPath, List<string> emailsToDelete)
        {
            SleepWhileSuspended();

            foreach (string entryId in emailsToDelete)
            {
                DeleteEmail(restSession, folderPath, entryId);
            }
        }

        private void DeleteEmail(Oris4RestSession restSession, string folderPath, string entryId)
        {
            SleepWhileSuspended();

            try
            {
                string dataHash = outlookDatabase.GetEmailDataHash(folderPath, entryId);
                if (dataHash == null)
                {
                    Logger.WarnFormat("Could not find email in database: {0}\\{1}", folderPath, entryId);
                    return;
                }

                Logger.InfoFormat("Deleting email from server: {0}", dataHash);
                restSession.deleteEmail(dataHash);

                outlookDatabase.RemoveEmail(folderPath, entryId);
                Logger.InfoFormat("Deleted email from database: {0}\\{1}", folderPath, entryId);
            }
            catch (System.Exception e)
            {
                ProcessRecoverableException(string.Format("Could not delete email: {0}\\{1}", folderPath, entryId), e);
            }
        }

        private void UploadAttachments(Oris4RestSession restSession, OutlookSession outlookSession, List<EmailAttachment> emailAttachments)
        {
            SleepWhileSuspended();

            try
            {
                foreach (EmailAttachment emailAttachment in emailAttachments)
                {
                    if (AttachmentWorthSyncing(emailAttachment) &&
                        outlookDatabase.ContainsEmail(emailAttachment.folderPath, emailAttachment.entryID) &&
                        !outlookDatabase.ContainsAttachment(emailAttachment.folderPath, emailAttachment.entryID, emailAttachment.fileName, emailAttachment.fileSize))
                    {
                        UploadAttachment(restSession, outlookSession, emailAttachment);
                    }
                }
            }
            catch (System.Exception e)
            {
                ProcessRecoverableException("Could not upload attachment list.", e);
            }
            finally
            {
                emailAttachments.Clear();
            }
        }

        private void UploadAttachment(Oris4RestSession restSession, OutlookSession outlookSession, EmailAttachment emailAttachment)
        {
            SleepWhileSuspended();

            try
            {
                EmailAttachment emailAttachmentWithTempFile = outlookSession.getEmailAttachmentWithTempFile(emailAttachment);

                try
                {
                    Logger.InfoFormat("Uploading attachment to server: {0}\\{1}\\{2}", emailAttachment.folderPath, emailAttachment.entryID, emailAttachment.fileName);
                    string returnValue = restSession.insertAttachment(emailAttachment, File.ReadAllBytes(emailAttachment.tempFilePath));

                    outlookDatabase.AddAttachment(emailAttachment.folderPath, emailAttachment.entryID, emailAttachment.fileName, emailAttachment.fileSize,
                        emailAttachment.emailDataHash, emailAttachment.dataHash, DateTime.Now);
                    Logger.InfoFormat("Added attachment to database: {0}\\{1}\\{2}", emailAttachment.folderPath, emailAttachment.entryID, emailAttachment.fileName);
                }
                finally
                {
                    File.Delete(emailAttachment.tempFilePath);
                }
            }
            catch (System.Exception e)
            {
                ProcessRecoverableException("Could not upload attachment: ", e);
            }
        }
    }
}

using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
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
            string dataPath = repoInfo.CmisDatabase;
            this.outlookDatabase = new OutlookDatabase(Path.Combine(Path.GetDirectoryName(dataPath),
                Path.GetFileNameWithoutExtension(dataPath) + ".outlook"));

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

            Oris4RestSession restSession = new Oris4RestSession(repoUrl);
            OutlookSession outlookSession = new OutlookSession();
            try
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
                            Email email = outlookSession.getEmail(folder, mailItem);

                            if (EmailWorthSyncing(email))
                            {
                                allEmailsInFolder.Add(email.dataHash);

                                if (!outlookDatabase.ContainsEmail(email.dataHash))
                                {
                                    emailList.Add(email);
                                }

                                attachmentList.AddRange(outlookSession.getEmailAttachments(mailItem, email));
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

                    //TODO:: Delete left older folder items...
                    DeleteObsoleteEmailsFromFolder(restSession, folderPath, allEmailsInFolder);
                }

                DeleteObsoleteFolders(restSession, folderPaths);
            }
            finally
            {
                restSession = null;
                outlookSession.close();
                outlookSession = null;
                GC.Collect(); //Ensure Outlook objects are released.
                GC.WaitForPendingFinalizers();
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
                restSession.putRegisteredClient(clientId);
                outlookDatabase.RemoveAllEmails();
                outlookDatabase.SetClientId(clientId);
            }
        }

        private bool EmailWorthSyncing(Email email)
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
                        outlookDatabase.AddEmail(email.dataHash, email.folderPath, email.entryID, emailKey, DateTime.Now);
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
                HashSet<string> emailsInDatabase = outlookDatabase.ListEmailDataHashes(folderPath);
                List<string> obsoleteEmails = new List<string>();
                foreach (string dataHash in emailsInDatabase)
                {
                    if (!emailsInFolder.Contains(dataHash))
                    {
                        obsoleteEmails.Add(dataHash);
                    }
                }

                if (obsoleteEmails.Count > 0)
                {
                    Logger.InfoFormat("Deleting {0} emails on server from folder: {1}", obsoleteEmails.Count, folderPath);
                    DeleteEmails(restSession, obsoleteEmails);
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
                        List<string> obsoleteEmails = new List<string>(outlookDatabase.ListEmailDataHashes(folderPath));
                        Logger.InfoFormat("Deleting {0} emails on server from folder: {1}", obsoleteEmails.Count, folderPath);
                        DeleteEmails(restSession, obsoleteEmails);
                    }
                }
            }
            catch (System.Exception e)
            {
                ProcessRecoverableException("Could not delete email obsolete folders.", e);
            }
        }


        private void DeleteEmails(Oris4RestSession restSession, List<string> emailsToDelete)
        {
            SleepWhileSuspended();

            foreach (string dataHash in emailsToDelete)
            {
                DeleteEmail(restSession, dataHash);
            }
        }

        private void DeleteEmail(Oris4RestSession restSession, string dataHash)
        {
            SleepWhileSuspended();

            try
            {
                Logger.InfoFormat("Deleting email from server: {0}", dataHash);
                restSession.deleteEmail(dataHash);

                outlookDatabase.RemoveEmail(dataHash);
                Logger.InfoFormat("Deleted email from database: {0}", dataHash);
            }
            catch (System.Exception e)
            {
                ProcessRecoverableException("Could not delete email: " + dataHash, e);
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
                        outlookDatabase.ContainsEmail(emailAttachment.emailDataHash) &&
                        !outlookDatabase.ContainsAttachment(emailAttachment.emailDataHash, emailAttachment.fileName, emailAttachment.fileSize))
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
                    Logger.InfoFormat("Uploading attachment to server: {0}\\{1} ({2})", emailAttachment.folderPath, emailAttachment.fileName, emailAttachment.emailDataHash);
                    string returnValue = restSession.insertAttachment(emailAttachment, File.ReadAllBytes(emailAttachment.tempFilePath));

                    outlookDatabase.AddAttachment(emailAttachment.emailDataHash, emailAttachment.fileName, emailAttachment.fileSize,
                        emailAttachment.folderPath, emailAttachment.dataHash, DateTime.Now);
                    Logger.InfoFormat("Added attachment to database: {0}\\{1} ({2})", emailAttachment.folderPath, emailAttachment.fileName, emailAttachment.emailDataHash);
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

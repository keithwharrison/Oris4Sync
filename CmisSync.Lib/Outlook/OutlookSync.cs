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
                Path.GetFileNameWithoutExtension(dataPath) + " (outlook)" +
                Path.GetExtension(dataPath)));

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
            if (!repoInfo.OutlookEnabled || !fullSync || !OutlookService.isOutlookInstalled() ||
                !OutlookService.isOutlookProfileAvailable() || !OutlookService.isOutlookSecurityManagerBitnessMatch())
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
                outlookSession.sendAndRecieve();

                string[] folderPaths = repoInfo.getOutlookFolders();

                HashSet<string> allEmailsInOutlook = new HashSet<string>();

                foreach (string folderPath in folderPaths)
                {
                    MAPIFolder folder = outlookSession.getFolderByPath(folderPath);
                    if (folder == null)
                    {
                        Logger.ErrorFormat("Could not find selected outlook folder: {0}", folderPath);
                        continue;
                    }

                    Logger.InfoFormat("Syncing Outlook Folder: {0}", folder.FolderPath);

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
                                allEmailsInOutlook.Add(email.dataHash);

                                if (!outlookDatabase.ContainsEmail(email.dataHash))
                                {
                                    emailList.Add(email);
                                }

                                attachmentList.AddRange(outlookSession.getEmailAttachments(mailItem, email));
                            }

                            if (emailList.Count >= EMAIL_BATCH_SIZE)
                            {
                                UploadEmails(restSession, emailList);
                                UploadAttachments(restSession, attachmentList);
                            }
                        }
                    }

                    if (emailList.Count > 0)
                    {
                        UploadEmails(restSession, emailList);
                        UploadAttachments(restSession, attachmentList);
                    }
                }

                HashSet<string> allEmails = outlookDatabase.ListEmailDataHashes();
                List<string> emailsToDelete = new List<string>();
                foreach (string dataHash in allEmails)
                {
                    if (!allEmailsInOutlook.Contains(dataHash))
                    {
                        emailsToDelete.Add(dataHash);
                    }
                }
                if (emailsToDelete.Count > 0)
                {
                    DeleteEmails(restSession, emailsToDelete);
                }
            }
            finally
            {
                outlookSession.close();
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
                Logger.InfoFormat("Outlook Client ID not found creating a new one: {0}", clientId);
            }

            string registeredClient = restSession.getRegisteredClient();
            Logger.InfoFormat("Registered Outlook Client ID: {0}", registeredClient);
            if (!registeredClient.Equals(clientId))
            {
                SleepWhileSuspended();
                //TODO: Ask user if they are sure before putting a new client ID (all emails deleted)
                Logger.InfoFormat("Registering new client...");
                restSession.putRegisteredClient(clientId);
                outlookDatabase.RemoveAllEmails();
                outlookDatabase.SetClientId(clientId);
            }
        }

        private bool EmailWorthSyncing(Email email)
        {


            return true;
        }

        private bool AttachmentWorthSyncing()
        {

            return true;
        }

        private void UploadEmails(Oris4RestSession restSession, List<Email> emails)
        {
            SleepWhileSuspended();

            try
            {
                Dictionary<string, long> emailKeyMap = restSession.insertEmailBatch(emails);

                foreach (Email email in emails)
                {
                    if (emailKeyMap.ContainsKey(email.dataHash))
                    {
                        long emailKey = emailKeyMap[email.dataHash];
                        outlookDatabase.AddEmail(email.dataHash, email.folderPath, DateTime.Now); //TODO: add EntryID and email key into database?
                    }
                    else
                    {
                        Logger.ErrorFormat("Email was not inserted: {0}/{1}", email.folderPath, email.dataHash);
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

                restSession.deleteEmail(dataHash);

                outlookDatabase.RemoveEmail(dataHash);
            }
            catch (System.Exception e)
            {
                ProcessRecoverableException("Could not delete email: " + dataHash, e);
            }
        }

        private void UploadAttachments(Oris4RestSession restSession, List<EmailAttachment> emailAttachments)
        {
            SleepWhileSuspended();

            foreach (EmailAttachment emailAttachment in emailAttachments)
            {
                if (!outlookDatabase.ContainsAttachment(emailAttachment.emailDataHash, emailAttachment.dataHash, emailAttachment.fileName))
                {
                    UploadAttachment(restSession, emailAttachment);
                }
            }

            emailAttachments.Clear();
        }

        private void UploadAttachment(Oris4RestSession restSession, EmailAttachment emailAttachment)
        {
            SleepWhileSuspended();

            try
            {
                string returnValue = restSession.insertAttachment(emailAttachment, File.ReadAllBytes(emailAttachment.tempFilePath));

                //Todo: check return value...

                outlookDatabase.AddAttachment(emailAttachment.emailDataHash, emailAttachment.dataHash, emailAttachment.fileName,
                    emailAttachment.folderPath, DateTime.Now);

                File.Delete(emailAttachment.tempFilePath);
            }
            catch (System.Exception e)
            {
                ProcessRecoverableException("Could not upload attachment: ", e);
            }
        }
    }
}

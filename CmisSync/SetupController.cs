//   CmisSync, a collaboration and sharing tool.
//   Copyright (C) 2010  Hylke Bons <hylkebons@gmail.com>
//
//   This program is free software: you can redistribute it and/or modify
//   it under the terms of the GNU General Public License as published by
//   the Free Software Foundation, either version 3 of the License, or
//   (at your option) any later version.
//
//   This program is distributed in the hope that it will be useful,
//   but WITHOUT ANY WARRANTY; without even the implied warranty of
//   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
//   GNU General Public License for more details.
//
//   You should have received a copy of the GNU General Public License
//   along with this program. If not, see <http://www.gnu.org/licenses/>.


using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;

using CmisSync.Lib;
using CmisSync.Lib.Cmis;
using log4net;
using System.ComponentModel;
using CmisSync.Lib.Outlook;

namespace CmisSync
{

    /// <summary>
    /// Kind of pages that are used in the folder addition wizards.
    /// </summary>
    public enum PageType
    {
        /// <summary>
        /// No page.
        /// </summary>
        None,
        /// <summary>
        /// Setup page (add to startup items).
        /// </summary>
        Setup,
        /// <summary>
        /// Add repository page.
        /// </summary>
        Add1,
        /// <summary>
        /// Select remote folder.
        /// </summary>
        Add2,
        /// <summary>
        /// Select name/local folder.
        /// </summary>
        Customize,
        /// <summary>
        /// Configure outlook.
        /// </summary>
        Outlook,
        /// <summary>
        /// Add complete.
        /// </summary>
        Finished,
        /// <summary>
        /// Settings page.
        /// </summary>
        Settings,
    }

    /// <summary>
    /// MVC controller for the two wizards:
    /// - wizard to add a new remote folder.
    /// </summary>
    public class SetupController
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(SetupController));

        private static readonly String MY_FILES = "My Files";
        private static readonly String ROOT_PATH = "/";

        /// <summary>
        /// Show window event.
        /// </summary>
        public event Action ShowWindowEvent = delegate { };
        
        /// <summary>
        /// Hide window event.
        /// </summary>
        public event Action HideWindowEvent = delegate { };

        /// <summary>
        /// Change page event.
        /// </summary>
        public event ChangePageEventHandler ChangePageEvent = delegate { };
        
        /// <summary>
        /// Change page event.
        /// </summary>
        public delegate void ChangePageEventHandler(PageType page);

        /// <summary>
        /// Update setup continue button.
        /// </summary>
        public event UpdateSetupContinueButtonEventHandler UpdateSetupContinueButtonEvent = delegate { };

        /// <summary>
        /// Update setup continue button.
        /// </summary>
        public delegate void UpdateSetupContinueButtonEventHandler(bool button_enabled);

        /// <summary>
        /// Update add project button event.
        /// </summary>
        public event UpdateAddProjectButtonEventHandler UpdateAddProjectButtonEvent = delegate { };

        /// <summary>
        /// Update add project button event.
        /// </summary>
        public delegate void UpdateAddProjectButtonEventHandler(bool button_enabled);

        /// <summary>
        /// Change address field event.
        /// </summary>
        public event ChangeAddressFieldEventHandler ChangeAddressFieldEvent = delegate { };
        
        /// <summary>
        /// Change address field event.
        /// </summary>
        public delegate void ChangeAddressFieldEventHandler(string text, string example_text);

        /// <summary>
        /// Change repository field event.
        /// </summary>
        public event ChangeRepositoryFieldEventHandler ChangeRepositoryFieldEvent = delegate { };
        /// <summary>
        /// Change repository field event.
        /// </summary>
        public delegate void ChangeRepositoryFieldEventHandler(string text, string example_text);

        /// <summary>
        /// Change path field event.
        /// </summary>
        public event ChangePathFieldEventHandler ChangePathFieldEvent = delegate { };
        /// <summary>
        /// Change path field event.
        /// </summary>
        public delegate void ChangePathFieldEventHandler(string text, string example_text);

        /// <summary>
        /// Change user field.
        /// </summary>
        public event ChangeUserFieldEventHandler ChangeUserFieldEvent = delegate { };
        /// <summary>
        /// Change user field.
        /// </summary>
        public delegate void ChangeUserFieldEventHandler(string text, string example_text);

        /// <summary>
        /// Change password event.
        /// </summary>
        public event ChangePasswordFieldEventHandler ChangePasswordFieldEvent = delegate { };
        /// <summary>
        /// Change password event.
        /// </summary>
        public delegate void ChangePasswordFieldEventHandler(string text, string example_text);

        /// <summary>
        /// Whether the window is currently open.
        /// </summary>
        public bool WindowIsOpen { get; private set; }

        /// <summary>
        /// Current step of the remote folder addition wizard.
        /// </summary>
        private PageType FolderAdditionWizardCurrentPage;

        /// <summary>
        /// Previous address.
        /// </summary>
        public Uri PreviousAddress { get; private set; }

        /// <summary>
        /// Event.
        /// </summary>
        public string PreviousPath { get; private set; }

        /// <summary>
        /// Previous repository.
        /// </summary>
        public string PreviousRepository { get; private set; }

        /// <summary>
        /// Syncing repository name.
        /// </summary>
        public string SyncingReponame { get; private set; }

        /// <summary>
        /// Default repository path..
        /// </summary>
        public string DefaultRepoPath { get; private set; }

        /// <summary>
        /// Progress bar percentage.
        /// </summary>
        public double ProgressBarPercentage { get; private set; }

        /// <summary>
        /// Saved address.
        /// </summary>
        public Uri saved_address = null;

        /// <summary>
        /// Saved remote path.
        /// </summary>
        public string saved_remote_path = "";

        /// <summary>
        /// Saved user.
        /// </summary>
        public string saved_user = "";

        /// <summary>
        /// Saved password.
        /// </summary>
        public string saved_password = "";

        /// <summary>
        /// Saved repository.
        /// </summary>
        public string saved_repository = "";

        /// <summary>
        /// Saved local repository directory.
        /// </summary>
        public string saved_local_path = "";

        /// <summary>
        /// Saved sync interval.
        /// </summary>
        public int saved_sync_interval = 15;

        /// <summary>
        /// Ignored paths.
        /// </summary>
        public List<string> ignoredPaths = new List<string>();

        /// <summary>
        /// List of the CMIS repositories at the chosen URL.
        /// </summary>
        public Dictionary<string, string> repositories;

        /// <summary>
        /// Whether CmisSync should be started automatically at login.
        /// </summary>
        private bool create_startup_item = true;

        /// <summary>
        /// Whether or not outlook is inabled.
        /// </summary>
        public bool saved_outlook_enabled = false;

        /// <summary>
        /// Selected outlook folders.
        /// </summary>
        public List<string> saved_outlook_folders = new List<string>();

        /// <summary>
        /// Load repositories information from a CMIS endpoint.
        /// </summary>
        static public Tuple<CmisServer, Exception> GetRepositoriesFuzzy(string url, string user, string password)
        {
            Uri uri;
            try
            {
                uri = new Uri(url);
                return CmisUtils.GetRepositoriesFuzzy(uri, user, password);
            }
            catch (Exception e)
            {
                return new Tuple<CmisServer, Exception>(null, e);
            }

        }


        /// <summary>
        /// Get the list of subfolders contained in a CMIS folder.
        /// </summary>
        static public string[] GetSubfolders(string repositoryId, string path,
            string address, string user, string password)
        {
            return CmisUtils.GetSubfolders(repositoryId, path, address, user, password);
        }

        /// <summary>
        /// Regex to check an HTTP/HTTPS URL.
        /// </summary>
        private Regex UrlRegex = new Regex(@"^" +
                    "(https?)://" +                                                 // protocol
                    "(([a-z\\d$_\\.\\+!\\*'\\(\\),;\\?&=-]|%[\\da-f]{2})+" +        // username
                    "(:([a-z\\d$_\\.\\+!\\*'\\(\\),;\\?&=-]|%[\\da-f]{2})+)?" +     // password
                    "@)?(?#" +                                                      // auth delimiter
                    ")((([a-z\\d]\\.|[a-z\\d][a-z\\d-]*[a-z\\d]\\.)*" +             // domain segments AND
                    "[a-z][a-z\\d-]*[a-z\\d]" +                                     // top level domain OR
                    "|((\\d|\\d\\d|1\\d{2}|2[0-4]\\d|25[0-5])\\.){3}" +             // IP address
                    "(\\d|[1-9]\\d|1\\d{2}|2[0-4]\\d|25[0-5])" +                    //
                    ")(:\\d+)?" +                                                   // port
                    ")(.*)" +                                                       // path
                    "$", RegexOptions.IgnoreCase | RegexOptions.Compiled);


        /// <summary>
        /// Regex to check a CmisSync repository local folder name.
        /// Basically, it should be a valid local filesystem folder name.
        /// </summary>
        Regex RepositoryRegex = new Regex(@"^([a-zA-Z0-9][^*/><?\|:]*)$");


        /// <summary>
        /// Constructor.
        /// </summary>
        public SetupController()
        {
            Logger.Debug("Entering constructor.");

            PreviousAddress = null;
            PreviousPath = "";
            SyncingReponame = "";
            DefaultRepoPath = Program.Controller.FoldersPath;

            // Actions.

            ChangePageEvent += delegate(PageType page)
            {
                this.FolderAdditionWizardCurrentPage = page;
            };

            Program.Controller.ShowSetupWindowEvent += delegate(PageType page)
            {
                if (this.FolderAdditionWizardCurrentPage == PageType.Finished)
                {
                    ShowWindowEvent();
                    return;
                }

                if (page == PageType.Add1)
                {
                    if (WindowIsOpen)
                    {
                        if (this.FolderAdditionWizardCurrentPage == PageType.Finished ||
                            this.FolderAdditionWizardCurrentPage == PageType.None)
                        {

                            ChangePageEvent(PageType.Add1);
                        }

                        ShowWindowEvent();

                    }
                    else
                    {
                        WindowIsOpen = true;
                        ChangePageEvent(PageType.Add1);
                        ShowWindowEvent();
                    }
                    return;
                }

                WindowIsOpen = true;
                ChangePageEvent(page);
                ShowWindowEvent();
            };
            Logger.Debug("Exiting constructor.");
        }


        /// <summary>
        /// User pressed the "Cancel" button, hide window.
        /// </summary>
        public void PageCancelled()
        {
            PreviousAddress = null;
            PreviousRepository = "";
            PreviousPath = "";
            ignoredPaths.Clear();

            WindowIsOpen = false;
            HideWindowEvent();
        }


        /// <summary>
        /// Check setup page.
        /// </summary>
        public void CheckSetupPage()
        {
            UpdateSetupContinueButtonEvent(true);
        }


        /// <summary>
        /// First-time wizard has been cancelled, so quit CmisSync.
        /// </summary>
        public void SetupPageCancelled()
        {
            Program.Controller.Quit();
        }


        /// <summary>
        /// Move to the add repository page...
        /// </summary>
        public void SetupPageCompleted()
        {
            // If requested, add CmisSync to the list of programs to be started up when the user logs into Windows.
            if (this.create_startup_item)
                new Thread(() => Program.Controller.CreateStartupItem()).Start();

            ChangePageEvent(PageType.Add1);
        }


        /// <summary>
        /// Checkbox to add CmisSync to the list of programs to be started up when the user logs into Windows.
        /// </summary>
        public void StartupItemChanged(bool create_startup_item)
        {
            this.create_startup_item = create_startup_item;
        }


        /// <summary>
        /// Check whether the address is syntaxically valid.
        /// If OK, enable button to next step.
        /// </summary>
        /// <param name="address">URL to check</param>
        /// <returns>validity error, or empty string if valid</returns>
        public string CheckAddPage(string address)
        {
            address = address.Trim();


            bool emptyAddress = string.IsNullOrEmpty(address);
            bool rejexMatch = this.UrlRegex.IsMatch(address);
            // Check address validity.
            if (!emptyAddress && rejexMatch)
            {
                try
                {
                    this.saved_address = new Uri(address);
                }
                catch (Exception ex)
                {
                    Logger.Debug("Error creating URI: " + ex.Message, ex);
                    rejexMatch = false;
                }
            }
            // Enable button to next step.
            UpdateAddProjectButtonEvent(!emptyAddress && rejexMatch);

            // Return validity error, or empty string if valid.
            if (emptyAddress)
            {
                return "EmptyURLNotAllowed";
            }
            if (!rejexMatch)
            {
                return "InvalidURL";
            }
            return String.Empty;
        }

        /// <summary>
        /// Check local repository path and repo name.
        /// </summary>
        /// <param name="localpath"></param>
        /// <param name="reponame"></param>
        /// <returns>validity error, or empty string if valid</returns>
        public string CheckRepoPathAndName(string localpath, string reponame)
        {
            // Check whether foldername is already in use
            bool folderAlreadyExists = (Program.Controller.Folders.FindIndex(x => x.Equals(reponame, StringComparison.OrdinalIgnoreCase)) != -1);

            // Check whether folde rname contains invalid characters.
            bool valid = (RepositoryRegex.IsMatch(reponame) && (!folderAlreadyExists));

            if (!valid)
            {
                // Disable button to next step.
                UpdateAddProjectButtonEvent(false);
            }
            // Return validity error, or continue validating.
            if (folderAlreadyExists) return "FolderAlreadyExist";
            if (!RepositoryRegex.IsMatch(reponame)) return "InvalidFolderName";

            // Validate localpath
            folderAlreadyExists = Directory.Exists(localpath);

            valid = !folderAlreadyExists;

            // Enable button to next step.
            UpdateAddProjectButtonEvent(valid);

            // Return validity error, or empty string if valid.
            if (folderAlreadyExists) return "LocalDirectoryExist";
            return String.Empty;
        }

        /// <summary>
        /// Return the default name of the selected repository.
        /// </summary>
        public string getSelectedRepositoryDefaultName()
        {
            string localfoldername = ""; 
            foreach (KeyValuePair<String, String> repository in repositories)
            {
                if (repository.Key == saved_repository)
                {
                    localfoldername = repository.Value;
                    break;
                }
            }
            return /*Controller.saved_address.Host.ToString() + "\\" + */localfoldername;
        }

        /// <summary>
        /// First step of remote folder addition wizard is complete, switch to second step
        /// </summary>
        public void Add1PageCompleted(Uri address, string user, string password)
        {
            saved_address = address;
            saved_user = user;
            saved_password = password;

            //Automatic repository selection
            String repositorySelection = null;
            foreach (KeyValuePair<String, String> repository in this.repositories)
            {
                if (repository.Key == MY_FILES)
                {
                    repositorySelection = repository.Key;
                    break;
                }
            }

            if (repositorySelection != null)
            {
                this.saved_repository = repositorySelection;
                this.saved_remote_path = ROOT_PATH; //Automatic selection selects root folder
                Add2PageCompleted(this.saved_repository, this.saved_remote_path);
            }
            else
            {
                ChangePageEvent(PageType.Add2);
            }
        }


        /// <summary>
        /// Switch back from second to first step, presumably to change server or user.
        /// </summary>
        public void BackToPage1()
        {
            PreviousAddress = saved_address;
            PreviousPath = saved_user;
            ChangePageEvent(PageType.Add1);
        }


        /// <summary>
        /// Second step of remote folder addition wizard is complete, switch to customization step.
        /// </summary>
        public void Add2PageCompleted(string repository, string remote_path, string[] ignoredPaths, string[] selectedFolder)
        {
            SyncingReponame = Path.GetFileName(remote_path);
            ProgressBarPercentage = 1.0;

            Uri address = saved_address;
            repository = repository.Trim();
            remote_path = remote_path.Trim();

            PreviousAddress = address;
            PreviousRepository = repository;
            PreviousPath = remote_path;

            foreach (string ignore in ignoredPaths)
                this.ignoredPaths.Add(ignore);

            //Automatically Select Default location...
            String localRepoPath = Path.Combine(DefaultRepoPath, repository);
            String error = CheckRepoPathAndName(localRepoPath, repository);
            if (String.IsNullOrEmpty(error))
            {
                CustomizePageCompleted(repository, localRepoPath);
            }
            else
            {
                ChangePageEvent(PageType.Customize);
            }
        }

        /// <summary>
        /// Second step of remote folder addition wizard is complete, switch to customization step.
        /// </summary>
        public void Add2PageCompleted(string repository, string remote_path)
        {
            Add2PageCompleted(repository, remote_path, new string[] { }, new string[] { });
        }

        /// <summary>
        /// Determine if outlook integration is available.
        /// </summary>
        public bool isOutlookIntegrationAvailable()
        {
            //TODO: Check server outlook compatibility?
            return OutlookService.Instance.checkForOutlookInstallation() &&
                OutlookService.Instance.checkForProfile();
        }

        /// <summary>
        /// Customization step of remote folder addition wizard is complete, start CmisSync.
        /// </summary>
        public void CustomizePageCompleted(String repoName, String localrepopath)
        {
            SyncingReponame = repoName;
            saved_local_path = localrepopath;

            if (isOutlookIntegrationAvailable())
            {
                ChangePageEvent(PageType.Outlook);
            }
            else
            {
                Finish();
            }
        }

        /// <summary>
        /// Outlook configuration page completed.
        /// </summary>
        public void OutlookPageCompleted(bool outlookEnabled, List<string> outlookFolders)
        {
            this.saved_outlook_enabled = outlookEnabled;
            this.saved_outlook_folders = outlookFolders;

            Finish();
        }

        /// <summary>
        /// Wizard finished: create the repository and show the finished screen.
        /// </summary>
        public void Finish()
        {
            // Add the remote folder to the configuration and start syncing.
            try
            {
                Program.Controller.CreateRepository(
                    SyncingReponame,
                    saved_address,
                    saved_user.TrimEnd(),
                    saved_password.TrimEnd(),
                    PreviousRepository,
                    PreviousPath,
                    saved_local_path,
                    ignoredPaths,
                    saved_outlook_enabled,
                    saved_outlook_folders);
            }
            catch (Exception e)
            {
                Logger.Fatal("Could not create repository.", e);
                Program.Controller.ShowAlert(Properties_Resources.Error, String.Format(Properties_Resources.SyncError, SyncingReponame, e.Message));
                FinishPageCompleted();
            }

            ChangePageEvent(PageType.Finished);
        }


        /// <summary>
        /// Switch back from customization to step 2 of the remote folder addition wizard.
        /// </summary>
        public void BackToPage2()
        {
            ignoredPaths.Clear();

            if (saved_repository == MY_FILES)
            {
                BackToPage1();
            }
            else
            {
                ChangePageEvent(PageType.Add2);
            }
        }

        /// <summary>
        /// Switch back from outlook to customize page.
        /// </summary>
        public void BackToCustomize()
        {
            string defaultRepoName = getSelectedRepositoryDefaultName();
            if (SyncingReponame.Equals(defaultRepoName) &&
                saved_local_path.Equals(Path.Combine(DefaultRepoPath, defaultRepoName))) 
            {
                BackToPage2();
            }
            else
            {
                ChangePageEvent(PageType.Customize);
            }
        }

        /// <summary>
        /// User clicked on the button to open the newly-created synchronized folder in the local file explorer.
        /// </summary>
        public void OpenFolderClicked()
        {
            Program.Controller.OpenCmisSyncFolder(SyncingReponame);
            SyncingReponame = String.Empty;
            FinishPageCompleted();
        }


        /// <summary>
        /// Folder addition wizard is over, reset it for next use.
        /// </summary>
        public void FinishPageCompleted()
        {
            PreviousAddress = null;
            PreviousPath = "";

            this.FolderAdditionWizardCurrentPage = PageType.None;
            HideWindowEvent();
        }

        /// <summary>
        /// Repository settings page.
        /// </summary>
        public void SettingsPageCompleted(string password, int pollInterval, bool outlookEnabled, List<string> outlookFolders)
        {
            //Run this in background so as not to free the UI...
            BackgroundWorker worker = new BackgroundWorker();
            worker.DoWork += new DoWorkEventHandler(
                delegate(Object o, DoWorkEventArgs args)
                {
                    Program.Controller.UpdateRepositorySettings(saved_repository, password, pollInterval, outlookEnabled, outlookFolders.ToArray());
                }
            );
            worker.RunWorkerAsync();

            FinishPageCompleted();
        }
    }
}

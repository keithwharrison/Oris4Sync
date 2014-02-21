//   CmisSync, a collaboration and sharing tool.
//   Copyright (C) 2010  Hylke Bons <hylkebons@gmail.com>
//
//   This program is free software: you can redistribute it and/or modify
//   it under the terms of the GNU General private License as published by
//   the Free Software Foundation, either version 3 of the License, or
//   (at your option) any later version.
//
//   This program is distributed in the hope that it will be useful,
//   but WITHOUT ANY WARRANTY; without even the implied warranty of
//   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
//   GNU General private License for more details.
//
//   You should have received a copy of the GNU General private License
//   along with this program. If not, see <http://www.gnu.org/licenses/>.


using CmisSync.CmisTree;
using CmisSync.Lib.Cmis;
using CmisSync.Lib.Outlook;
using log4net;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Media;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Shell;

namespace CmisSync
{
    /// <summary>
    /// Dialog for the tutorial, and for the wizard to add a new remote folder.
    /// </summary>
    public class Setup : SetupWindow
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(Setup));

        /// <summary>
        /// MVC controller.
        /// </summary>
        public SetupController Controller = new SetupController();

        delegate Tuple<CmisServer, Exception> GetRepositoriesFuzzyDelegate(Uri url, string user, string password);

        delegate string[] GetSubfoldersDelegate(string repositoryId, string path,
            string address, string user, string password);

        delegate void CheckRepoPathAndNameDelegate();

        delegate List<OutlookFolder> GetOutlookFolderTreeDelegate();

        private EventHandler windowActivatedEventHandler = null;

        /// <summary>
        /// Constructor.
        /// </summary>
        public Setup()
        {
            Logger.Debug("Entering constructor.");

            // Defines how to show the setup window.
            Controller.ShowWindowEvent += delegate
            {
                Dispatcher.BeginInvoke((Action)delegate
                {
                    Logger.Debug("Entering ShowWindowEvent.");
                    Show();
                    Activate();
                    BringIntoView();
                    Logger.Debug("Exiting ShowWindowEvent.");
                });
            };

            // Defines how to hide the setup windows.
            Controller.HideWindowEvent += delegate
            {
                Dispatcher.BeginInvoke((Action)delegate
                {
                    Hide();
                });
            };

            // Defines what to do when changing page.
            // The remote folder addition wizard has several steps.
            Controller.ChangePageEvent += delegate(PageType type)
            {
                Dispatcher.BeginInvoke((Action)delegate
                {
                    Logger.Debug("Entering ChangePageEvent.");
                    Reset();

                    PageTitle = CmisSync.Properties_Resources.Oris4Sync;

                    //Remove window activated event handler if one exists...
                    if (windowActivatedEventHandler != null)
                    {
                        Logger.Debug("Removing window activated event handler");
                        Activated -= windowActivatedEventHandler;
                        windowActivatedEventHandler = null;
                    }

                    // Show appropriate setup page.
                    switch (type)
                    {
                        // Welcome page that shows up at first run.
                        #region Page Setup
                        case PageType.Setup:
                            {
                                // UI elements.

                                Header = Properties_Resources.Welcome;
                                Description = Properties_Resources.Intro;

                                Button cancel_button = new Button()
                                {
                                    Content = Properties_Resources.Cancel
                                };

                                Button continue_button = new Button()
                                {
                                    Content = Properties_Resources.Continue,
                                    IsEnabled = false
                                };

                                CheckBox check_box = new CheckBox()
                                {
                                    Content = Properties_Resources.Startup,
                                    IsChecked = true
                                };

                                ContentCanvas.Children.Add(check_box);
                                Canvas.SetLeft(check_box, 185);
                                Canvas.SetBottom(check_box, 12);

                                Buttons.Add(continue_button);
                                Buttons.Add(cancel_button);

                                continue_button.Focus();

                                // Actions.

                                Controller.UpdateSetupContinueButtonEvent += delegate(bool enabled)
                                {
                                    Dispatcher.BeginInvoke((Action)delegate
                                    {
                                        continue_button.IsEnabled = enabled;
                                    });
                                };

                                check_box.Click += delegate
                                {
                                    Controller.StartupItemChanged(check_box.IsChecked.Value);
                                };

                                cancel_button.Click += delegate
                                {
                                    Dispatcher.BeginInvoke((Action)delegate
                                    {
                                        Program.UI.StatusIcon.Dispose();
                                        Controller.SetupPageCancelled();
                                    });
                                };

                                continue_button.Click += delegate
                                {
                                    Controller.SetupPageCompleted();
                                };

                                Controller.CheckSetupPage();

                                break;
                            }
                        #endregion

                        // First step of the remote folder addition dialog: Specifying the server.
                        #region Page Add1
                        case PageType.Add1:
                            {
                                // UI elements.

                                Header = Properties_Resources.Connect;

                                // Address input UI.
                                TextBlock address_label = new TextBlock()
                                {
                                    Text = Properties_Resources.EnterWebAddress,
                                    FontWeight = FontWeights.Bold
                                };

                                TextBox address_box = new TextBox()
                                {
                                    Width = 420,
                                    Text = (Controller.PreviousAddress != null) ? Controller.PreviousAddress.ToString() : ""
                                };

                                TextBlock address_help_label = new TextBlock()
                                {
                                    Text = Properties_Resources.Hint + ": " + Properties_Resources.ExampleUrl,
                                    FontSize = 11,
                                    Foreground = new SolidColorBrush(Color.FromRgb(128, 128, 128))
                                };

                                // Rather than a TextBlock, we use a textBox so that users can copy/paste the error message and Google it.
                                TextBox address_error_label = new TextBox()
                                {
                                    FontSize = 11,
                                    Foreground = new SolidColorBrush(Color.FromRgb(255, 128, 128)),
                                    Visibility = Visibility.Hidden,
                                    IsReadOnly = true,
                                    Background = Brushes.Transparent,
                                    BorderThickness = new Thickness(0),
                                    TextWrapping = TextWrapping.Wrap,
                                    MaxWidth = 420
                                };

                                // User input UI.
                                TextBlock user_label = new TextBlock()
                                {
                                    Text = Properties_Resources.Username + ":",
                                    FontWeight = FontWeights.Bold,
                                    Width = 200
                                };

                                TextBox user_box = new TextBox()
                                {
                                    Width = 200
                                };
                                if (Controller.saved_user == String.Empty || Controller.saved_user == null)
                                {
                                    user_box.Text = Environment.UserName;
                                }
                                else
                                {
                                    user_box.Text = Controller.saved_user;
                                }

                                TextBlock user_help_label = new TextBlock()
                                {
                                    FontSize = 11,
                                    Width = 200,
                                    Foreground = new SolidColorBrush(Color.FromRgb(128, 128, 128))
                                };

                                // Password input UI.
                                TextBlock password_label = new TextBlock()
                                {
                                    Text = Properties_Resources.Password + ":",
                                    FontWeight = FontWeights.Bold,
                                    Width = 200
                                };

                                PasswordBox password_box = new PasswordBox()
                                {
                                    Width = 200
                                };

                                TextBlock password_help_label = new TextBlock()
                                {
                                    FontSize = 11,
                                    Width = 200,
                                    Foreground = new SolidColorBrush(Color.FromRgb(128, 128, 128))
                                };

                                // Buttons.
                                Button cancel_button = new Button()
                                {
                                    Content = Properties_Resources.Cancel
                                };

                                Button continue_button = new Button()
                                {
                                    Content = Properties_Resources.Continue
                                };

                                Buttons.Add(continue_button);
                                Buttons.Add(cancel_button);

                                // Address
                                ContentCanvas.Children.Add(address_label);
                                Canvas.SetTop(address_label, 160);
                                Canvas.SetLeft(address_label, 185);

                                ContentCanvas.Children.Add(address_box);
                                Canvas.SetTop(address_box, 180);
                                Canvas.SetLeft(address_box, 185);

                                ContentCanvas.Children.Add(address_help_label);
                                Canvas.SetTop(address_help_label, 205);
                                Canvas.SetLeft(address_help_label, 185);

                                ContentCanvas.Children.Add(address_error_label);
                                Canvas.SetTop(address_error_label, 235);
                                Canvas.SetLeft(address_error_label, 185);

                                // User
                                ContentCanvas.Children.Add(user_label);
                                Canvas.SetTop(user_label, 300);
                                Canvas.SetLeft(user_label, 185);

                                ContentCanvas.Children.Add(user_box);
                                Canvas.SetTop(user_box, 320);
                                Canvas.SetLeft(user_box, 185);

                                ContentCanvas.Children.Add(user_help_label);
                                Canvas.SetTop(user_help_label, 345);
                                Canvas.SetLeft(user_help_label, 185);

                                // Password
                                ContentCanvas.Children.Add(password_label);
                                Canvas.SetTop(password_label, 300);
                                Canvas.SetRight(password_label, 30);

                                ContentCanvas.Children.Add(password_box);
                                Canvas.SetTop(password_box, 320);
                                Canvas.SetRight(password_box, 30);

                                ContentCanvas.Children.Add(password_help_label);
                                Canvas.SetTop(password_help_label, 345);
                                Canvas.SetRight(password_help_label, 30);

                                TaskbarItemInfo.ProgressValue = 0.0;
                                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;

                                if (Controller.PreviousAddress == null || Controller.PreviousAddress.ToString() == String.Empty)
                                    address_box.Text = "https://";
                                else
                                    address_box.Text = Controller.PreviousAddress.ToString();
                                address_box.Focus();
                                address_box.Select(address_box.Text.Length, 0);

                                // Actions.

                                Controller.ChangeAddressFieldEvent += delegate(string text,
                                    string example_text)
                                {
                                    Dispatcher.BeginInvoke((Action)delegate
                                    {
                                        address_box.Text = text;
                                        address_help_label.Text = example_text;
                                    });
                                };

                                Controller.ChangeUserFieldEvent += delegate(string text,
                                    string example_text)
                                {
                                    Dispatcher.BeginInvoke((Action)delegate
                                    {
                                        user_box.Text = text;
                                        user_help_label.Text = example_text;
                                    });
                                };

                                Controller.ChangePasswordFieldEvent += delegate(string text,
                                    string example_text)
                                {
                                    Dispatcher.BeginInvoke((Action)delegate
                                    {
                                        password_box.Password = text;
                                        password_help_label.Text = example_text;
                                    });
                                };

                                Controller.UpdateAddProjectButtonEvent += delegate(bool button_enabled)
                                {
                                    Dispatcher.BeginInvoke((Action)delegate
                                    {
                                        continue_button.IsEnabled = button_enabled;
                                    });
                                };

                                Controller.CheckAddPage(address_box.Text);

                                address_box.TextChanged += delegate
                                {
                                    string error = Controller.CheckAddPage(address_box.Text);
                                    if (!String.IsNullOrEmpty(error))
                                    {
                                        address_error_label.Text = Properties_Resources.ResourceManager.GetString(error, CultureInfo.CurrentCulture);
                                        address_error_label.Visibility = Visibility.Visible;
                                    }
                                    else address_error_label.Visibility = Visibility.Hidden;
                                };

                                cancel_button.Click += delegate
                                {
                                    Controller.PageCancelled();
                                };

                                continue_button.Click += delegate
                                {
                                    // Show wait cursor
                                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

                                    // Try to find the CMIS server (asynchronously)
                                    GetRepositoriesFuzzyDelegate dlgt =
                                        new GetRepositoriesFuzzyDelegate(CmisUtils.GetRepositoriesFuzzy); //TODO:: Check for outlook capabilities here.
                                    IAsyncResult ar = dlgt.BeginInvoke(new Uri(address_box.Text), user_box.Text,
                                        password_box.Password, null, null);
                                    while (!ar.AsyncWaitHandle.WaitOne(100))
                                    {
                                        System.Windows.Forms.Application.DoEvents();
                                    }
                                    Tuple<CmisServer, Exception> result = dlgt.EndInvoke(ar);
                                    CmisServer cmisServer = result.Item1;

                                    Controller.repositories = cmisServer != null ? cmisServer.Repositories : null;

                                    address_box.Text = cmisServer.Url.ToString();

                                    // Hide wait cursor
                                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;

                                    if (Controller.repositories == null)
                                    {
                                        // Could not retrieve repositories list from server, show warning.
                                        string warning = "";
                                        string message = result.Item2.Message;
                                        Exception e = result.Item2;
                                        if (e is PermissionDeniedException)
                                        {
                                            warning = Properties_Resources.LoginFailedForbidden;
                                        }
                                        else if (e is ServerNotFoundException)
                                        {
                                            warning = Properties_Resources.ConnectFailure;
                                        }
                                        else if (e.Message == "SendFailure" && cmisServer.Url.Scheme.StartsWith("https"))
                                        {
                                            warning = Properties_Resources.SendFailureHttps;
                                        }
                                        else if (e.Message == "TrustFailure")
                                        {
                                            warning = Properties_Resources.TrustFailure;
                                        }
                                        else
                                        {
                                            warning = message + Environment.NewLine + Properties_Resources.Sorry;
                                        }
                                        address_error_label.Text = warning;
                                        address_error_label.Visibility = Visibility.Visible;
                                    }
                                    else
                                    {
                                        // Continue to next step, which is choosing a particular folder.
                                        Controller.Add1PageCompleted(
                                            new Uri(address_box.Text), user_box.Text, password_box.Password);
                                    }
                                };
                                break;
                            }
                        #endregion

                        // Second step of the remote folder addition dialog: choosing the folder.
                        #region Page Add2
                        case PageType.Add2:
                            {
                                // UI elements.

                                Header = Properties_Resources.Which;

                                // A tree allowing the user to browse CMIS repositories/folders.
                                /*if(TODO check if OpenDataSpace, and further separate code below)
                                {
                                    System.Uri resourceLocater = new System.Uri("/CmisSync;component/TreeView.xaml", System.UriKind.Relative);
                                    System.Windows.Controls.TreeView treeView = System.Windows.Application.LoadComponent(resourceLocater) as TreeView;
                                    ObservableCollection<CmisRepo> repos = new ObservableCollection<CmisRepo>();
                                */
                                System.Windows.Controls.TreeView treeView = new System.Windows.Controls.TreeView();
                                treeView.Width = 410;
                                treeView.Height = 267;

                                // Some CMIS servers hold several repositories (ex:Nuxeo). Show one root per repository.
                                foreach (KeyValuePair<String, String> repository in Controller.repositories)
                                {
                                    System.Windows.Controls.TreeViewItem item = new System.Windows.Controls.TreeViewItem();
                                    item.Tag = new SelectionTreeItem(repository.Key, "/");
                                    item.Header = repository.Value;
                                    treeView.Items.Add(item);
                                }

                                ContentCanvas.Children.Add(treeView);
                                Canvas.SetTop(treeView, 70);
                                Canvas.SetLeft(treeView, 185);

                                // Action: when an element in the tree is clicked, loads its children and show them.
                                treeView.SelectedItemChanged += delegate
                                {
                                    // Identify the selected remote path.
                                    TreeViewItem item = (TreeViewItem)treeView.SelectedItem;
                                    Controller.saved_remote_path = ((SelectionTreeItem)item.Tag).fullPath;

                                    // Identify the selected repository.
                                    object cursor = item;
                                    while (cursor is TreeViewItem)
                                    {
                                        TreeViewItem treeViewItem = (TreeViewItem)cursor;
                                        cursor = treeViewItem.Parent;
                                        if (!(cursor is TreeViewItem))
                                        {
                                            Controller.saved_repository = ((SelectionTreeItem)treeViewItem.Tag).repository;
                                        }
                                    }

                                    // Load sub-folders if it has not been done already.
                                    // We use each item's Tag to store metadata: whether this item's subfolders have been loaded or not.
                                    if (((SelectionTreeItem)item.Tag).childrenLoaded == false)
                                    {
                                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

                                        // Get list of subfolders (asynchronously)
                                        GetSubfoldersDelegate dlgt = new GetSubfoldersDelegate(CmisUtils.GetSubfolders);
                                        IAsyncResult ar = dlgt.BeginInvoke(Controller.saved_repository,
                                            Controller.saved_remote_path, Controller.saved_address.ToString(),
                                            Controller.saved_user, Controller.saved_password, null, null);
                                        while (!ar.AsyncWaitHandle.WaitOne(100))
                                        {
                                            System.Windows.Forms.Application.DoEvents();
                                        }
                                        string[] subfolders = dlgt.EndInvoke(ar);

                                        // Create a sub-item for each subfolder
                                        foreach (string subfolder in subfolders)
                                        {
                                            System.Windows.Controls.TreeViewItem subItem =
                                                new System.Windows.Controls.TreeViewItem();
                                            subItem.Tag = new SelectionTreeItem(null, subfolder);
                                            subItem.Header = Path.GetFileName(subfolder);
                                            item.Items.Add(subItem);
                                        }
                                        ((SelectionTreeItem)item.Tag).childrenLoaded = true;
                                        item.ExpandSubtree();
                                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                                    }
                                };

                                Button cancel_button = new Button()
                                {
                                    Content = Properties_Resources.Cancel
                                };

                                Button continue_button = new Button()
                                {
                                    Content = CmisSync.Properties_Resources.ResourceManager.GetString("Continue", CultureInfo.CurrentCulture)
                                };

                                Button back_button = new Button()
                                {
                                    Content = Properties_Resources.Back,
                                    IsDefault = false
                                };

                                Buttons.Add(back_button);
                                Buttons.Add(continue_button);
                                Buttons.Add(cancel_button);

                                continue_button.Focus();

                                cancel_button.Click += delegate
                                {
                                    Controller.PageCancelled();
                                };

                                continue_button.Click += delegate
                                {
                                    Controller.Add2PageCompleted(
                                        Controller.saved_repository, Controller.saved_remote_path);
                                };

                                back_button.Click += delegate
                                {
                                    Controller.BackToPage1();
                                };
                                break;
                            }
                        #endregion

                        // Third step of the remote folder addition dialog: Customizing the local folder.
                        #region Page Customize
                        case PageType.Customize:
                            {
                                string parentFolder = Controller.DefaultRepoPath;

                                // UI elements.

                                Header = Properties_Resources.Customize;

                                // Customize local folder name
                                TextBlock localfolder_label = new TextBlock()
                                {
                                    Text = Properties_Resources.EnterLocalFolderName,
                                    FontWeight = FontWeights.Bold,
                                    TextWrapping = TextWrapping.Wrap,
                                    Width = 420
                                };

                                string localfoldername = Controller.getSelectedRepositoryDefaultName();
                                TextBox localfolder_box = new TextBox()
                                {
                                    Width = 420,
                                    Text = localfoldername
                                };

                                TextBlock localrepopath_label = new TextBlock()
                                {
                                    Text = Properties_Resources.ChangeRepoPath,
                                    FontWeight = FontWeights.Bold
                                };

                                TextBox localrepopath_box = new TextBox()
                                {
                                    Width = 375,
                                    Text = Path.Combine(parentFolder, localfolder_box.Text)
                                };

                                Button choose_folder_button = new Button()
                                {
                                    Width = 40,
                                    Content = "..."
                                };

                                TextBlock localfolder_error_label = new TextBlock()
                                {
                                    FontSize = 11,
                                    Foreground = new SolidColorBrush(Color.FromRgb(255, 128, 128)),
                                    Visibility = Visibility.Hidden,
                                    TextWrapping = TextWrapping.Wrap
                                };

                                Button cancel_button = new Button()
                                {
                                    Content = Properties_Resources.Cancel
                                };

                                Button add_button = new Button()
                                {
                                    Content = Properties_Resources.Continue,
                                    IsDefault = true
                                };

                                Button back_button = new Button()
                                {
                                    Content = Properties_Resources.Back
                                };

                                Buttons.Add(back_button);
                                Buttons.Add(add_button);
                                Buttons.Add(cancel_button);

                                // Local Folder Name
                                ContentCanvas.Children.Add(localfolder_label);
                                Canvas.SetTop(localfolder_label, 160);
                                Canvas.SetLeft(localfolder_label, 185);

                                ContentCanvas.Children.Add(localfolder_box);
                                Canvas.SetTop(localfolder_box, 180);
                                Canvas.SetLeft(localfolder_box, 185);

                                ContentCanvas.Children.Add(localrepopath_label);
                                Canvas.SetTop(localrepopath_label, 200);
                                Canvas.SetLeft(localrepopath_label, 185);

                                ContentCanvas.Children.Add(localrepopath_box);
                                Canvas.SetTop(localrepopath_box, 220);
                                Canvas.SetLeft(localrepopath_box, 185);

                                ContentCanvas.Children.Add(choose_folder_button);
                                Canvas.SetTop(choose_folder_button, 220);
                                Canvas.SetLeft(choose_folder_button, 565);

                                ContentCanvas.Children.Add(localfolder_error_label);
                                Canvas.SetTop(localfolder_error_label, 275);
                                Canvas.SetLeft(localfolder_error_label, 185);

                                TaskbarItemInfo.ProgressValue = 0.0;
                                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;

                                localfolder_box.Focus();
                                localfolder_box.Select(localfolder_box.Text.Length, 0);

                                // Actions.

                                Controller.UpdateAddProjectButtonEvent += delegate(bool button_enabled)
                                {
                                    Dispatcher.BeginInvoke((Action)delegate
                                    {
                                        if (add_button.IsEnabled != button_enabled)
                                        {
                                            add_button.IsEnabled = button_enabled;
                                            if (button_enabled)
                                            {
                                                add_button.IsDefault = true;
                                                back_button.IsDefault = false;
                                            }
                                        }
                                    });
                                };

                                CheckRepoPathAndNameDelegate checkRepoPathAndNameDelegate = delegate()
                                {
                                    string error = Controller.CheckRepoPathAndName(localrepopath_box.Text, localfolder_box.Text);

                                    if (!String.IsNullOrEmpty(error))
                                    {
                                        localfolder_error_label.Text = Properties_Resources.ResourceManager.GetString(error, CultureInfo.CurrentCulture);
                                        localfolder_error_label.Visibility = Visibility.Visible;
                                    }
                                    else
                                    {
                                        localfolder_error_label.Visibility = Visibility.Hidden;
                                    }
                                };

                                //execute the check on first run...
                                checkRepoPathAndNameDelegate();

                                localfolder_box.TextChanged += delegate
                                {
                                    localrepopath_box.Text = Path.Combine(parentFolder, localfolder_box.Text);
                                };

                                localrepopath_box.TextChanged += delegate
                                {
                                    checkRepoPathAndNameDelegate();
                                };

                                windowActivatedEventHandler = new EventHandler( delegate(object sender, EventArgs e)
                                {
                                    checkRepoPathAndNameDelegate();
                                });

                                Activated += windowActivatedEventHandler;

                                // Choose a folder.
                                choose_folder_button.Click += delegate
                                {
                                    System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
                                    if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                    {
                                        parentFolder = folderBrowserDialog1.SelectedPath;
                                        localrepopath_box.Text = Path.Combine(parentFolder, localfolder_box.Text);
                                    }
                                };

                                // Other actions.

                                cancel_button.Click += delegate
                                {
                                    Controller.PageCancelled();
                                };

                                back_button.Click += delegate
                                {
                                    Controller.BackToPage2();
                                };

                                add_button.Click += delegate
                                {
                                    Controller.CustomizePageCompleted(localfolder_box.Text, localrepopath_box.Text);
                                };
                                break;
                            }
                        #endregion

                        // Outlook integration.
                        #region Page Outlook
                        case PageType.Outlook:
                            {
                                // UI elements.

                                Header = Properties_Resources.OutlookSyncWithOris4;
                                

                                CheckBox outlookCheckbox = new CheckBox()
                                {
                                    Content = Properties_Resources.OutlookEnable,
                                    IsChecked = false,
                                };

                                TextBlock outlookTreeViewLabel = new TextBlock()
                                {
                                    Text = Properties_Resources.OutlookSelectFolders,
                                    Visibility = Visibility.Hidden,
                                };
                                
                                TreeView outlookTreeView = new TreeView()
                                {
                                    Width = 420,
                                    Height = 235,
                                    Visibility = Visibility.Hidden,
                                };

                                Button cancel_button = new Button()
                                {
                                    Content = Properties_Resources.Cancel
                                };

                                Button continue_button = new Button()
                                {
                                    Content = CmisSync.Properties_Resources.Continue,
                                    IsDefault = true
                                };

                                Button back_button = new Button()
                                {
                                    Content = Properties_Resources.Back,
                                };

                                Buttons.Add(back_button);
                                Buttons.Add(continue_button);
                                Buttons.Add(cancel_button);

                                // Place items
                                ContentCanvas.Children.Add(outlookCheckbox);
                                Canvas.SetTop(outlookCheckbox, 60);
                                Canvas.SetLeft(outlookCheckbox, 185);

                                ContentCanvas.Children.Add(outlookTreeViewLabel);
                                Canvas.SetTop(outlookTreeViewLabel, 90);
                                Canvas.SetLeft(outlookTreeViewLabel, 185);
                                 
                                ContentCanvas.Children.Add(outlookTreeView);
                                Canvas.SetTop(outlookTreeView, 110);
                                Canvas.SetLeft(outlookTreeView, 185);

                                outlookCheckbox.Focus();

                                outlookCheckbox.Click += delegate
                                {
                                    if (outlookCheckbox.IsChecked.Value && outlookTreeView.Items.Count <= 0)
                                    {
                                        //Populate tree...
                                        populateOutlookTreeView(outlookTreeView);
                                    }
                                    outlookTreeViewLabel.Visibility = outlookCheckbox.IsChecked.Value ? Visibility.Visible : Visibility.Hidden;
                                    outlookTreeView.Visibility = outlookCheckbox.IsChecked.Value ? Visibility.Visible : Visibility.Hidden;
                                };

                                cancel_button.Click += delegate
                                {
                                    Controller.PageCancelled();
                                };

                                continue_button.Click += delegate
                                {
                                    bool outlookEnabled = outlookCheckbox.IsChecked.Value;
                                    List<string> outlookFolders = new List<string>();
                                    if (outlookEnabled)
                                    {
                                        getSelectedOutlookFolders(outlookTreeView.Items, outlookFolders);
                                        if (outlookFolders.Count <= 0)
                                        {
                                            outlookEnabled = false;
                                        }
                                    }
                                    Controller.OutlookPageCompleted(outlookEnabled, outlookFolders);
                                };

                                back_button.Click += delegate
                                {
                                    Controller.BackToCustomize();
                                };
                                break;
                            }
                        #endregion
                        
                        // Final page of the remote folder addition dialog: end of the addition wizard.
                        #region Page Finished
                        case PageType.Finished:
                            {
                                // UI elements.

                                Header = Properties_Resources.Ready;
                                Description = Properties_Resources.YouCanFind;

                                Button finish_button = new Button()
                                {
                                    Content = Properties_Resources.Finish
                                };

                                Button open_folder_button = new Button()
                                {
                                    Content = Properties_Resources.OpenFolder
                                };

                                TaskbarItemInfo.ProgressValue = 0.0;
                                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;

                                Buttons.Add(open_folder_button);
                                Buttons.Add(finish_button);

                                // Actions.

                                finish_button.Click += delegate
                                {
                                    Controller.FinishPageCompleted();
                                };

                                open_folder_button.Click += delegate
                                {
                                    Controller.OpenFolderClicked();
                                };


                                SystemSounds.Exclamation.Play();

                                break;
                            }
                        #endregion

                        // Settings Dialog...
                        #region Page Settings
                        case PageType.Settings:
                            {
                                // UI elements.
                                PageTitle = Properties_Resources.Settings;

                                const string OUTLOOK = "outlook";

                                // Address input UI.
                                TextBlock address_label = new TextBlock()
                                {
                                    Text = Properties_Resources.WebAddress,
                                    FontWeight = FontWeights.Bold
                                };

                                TextBox address_box = new TextBox()
                                {
                                    Width = 400,
                                    Text = Controller.saved_address.ToString(),
                                    IsEnabled = false,
                                };
                                
                                // User input UI.
                                TextBlock user_label = new TextBlock()
                                {
                                    Text = Properties_Resources.Username + ":",
                                    FontWeight = FontWeights.Bold,
                                    Width = 200
                                };

                                TextBox user_box = new TextBox()
                                {
                                    Width = 200,
                                    Text = Controller.saved_user,
                                    IsEnabled = false,
                                };

                                // Password input UI.
                                TextBlock password_label = new TextBlock()
                                {
                                    Text = Properties_Resources.Password + ":",
                                    FontWeight = FontWeights.Bold,
                                    Width = 200
                                };

                                PasswordBox password_box = new PasswordBox()
                                {
                                    Width = 200
                                };

                                // Rather than a TextBlock, we use a textBox so that users can copy/paste the error message and Google it.
                                TextBox authentication_error_label = new TextBox()
                                {
                                    FontSize = 11,
                                    Foreground = new SolidColorBrush(Color.FromRgb(255, 128, 128)),
                                    Visibility = Visibility.Hidden,
                                    IsReadOnly = true,
                                    Background = Brushes.Transparent,
                                    BorderThickness = new Thickness(0),
                                    TextWrapping = TextWrapping.Wrap,
                                    MaxWidth = 400
                                };

                                // Sync duration input UI.
                                TextBlock slider_label = new TextBlock()
                                {
                                    Text = Properties_Resources.SyncInterval + ":",
                                    FontWeight = FontWeights.Bold,
                                    Width = 200,
                                };

                                PollIntervalSlider slider = new PollIntervalSlider()
                                {
                                    Width = 400,
                                    PollInterval = Controller.saved_sync_interval,
                                };

                                TextBlock slider_min_label = new TextBlock()
                                {
                                    Text = slider.FormattedMinimum(),
                                    Width = 200,
                                };

                                TextBlock slider_max_label = new TextBlock()
                                {
                                    Text = slider.FormattedMaximum(),
                                    Width = 200,
                                    TextAlignment = TextAlignment.Right,
                                };

                                //Outlook
                                CheckBox outlookCheckbox = new CheckBox()
                                {
                                    Content = Properties_Resources.OutlookEnable,
                                    IsChecked = Controller.saved_outlook_enabled,
                                };

                                TextBlock outlookTreeViewLabel = new TextBlock()
                                {
                                    Text = Properties_Resources.OutlookSelectFolders,
                                    Visibility = Visibility.Hidden,
                                };

                                TreeView outlookTreeView = new TreeView()
                                {
                                    Width = 400,
                                    Height = 235,
                                    Visibility = Visibility.Hidden,
                                };

                                // Buttons.
                                Button cancel_button = new Button()
                                {
                                    Content = Properties_Resources.Cancel,
                                };

                                Button save_button = new Button()
                                {
                                    Content = Properties_Resources.Save,
                                    IsDefault = true,
                                };

                                Buttons.Add(save_button);
                                Buttons.Add(cancel_button);

                                //General settings tab...
                                Canvas generalCanvas = new Canvas();
                                
                                // Address
                                generalCanvas.Children.Add(address_label);
                                Canvas.SetTop(address_label, 10);
                                Canvas.SetLeft(address_label, 10);

                                generalCanvas.Children.Add(address_box);
                                Canvas.SetTop(address_box, 30);
                                Canvas.SetLeft(address_box, 10);

                                // User
                                generalCanvas.Children.Add(user_label);
                                Canvas.SetTop(user_label, 70);
                                Canvas.SetLeft(user_label, 10);

                                generalCanvas.Children.Add(user_box);
                                Canvas.SetTop(user_box, 90);
                                Canvas.SetLeft(user_box, 10);

                                // Password
                                generalCanvas.Children.Add(password_label);
                                Canvas.SetTop(password_label, 130);
                                Canvas.SetLeft(password_label, 10);

                                generalCanvas.Children.Add(password_box);
                                Canvas.SetTop(password_box, 150);
                                Canvas.SetLeft(password_box, 10);

                                // Error label
                                generalCanvas.Children.Add(authentication_error_label);
                                Canvas.SetTop(authentication_error_label, 175);
                                Canvas.SetLeft(authentication_error_label, 10);

                                // Sync Interval
                                generalCanvas.Children.Add(slider_label);
                                Canvas.SetTop(slider_label, 210);
                                Canvas.SetLeft(slider_label, 10);

                                generalCanvas.Children.Add(slider);
                                Canvas.SetTop(slider, 230);
                                Canvas.SetLeft(slider, 10);

                                generalCanvas.Children.Add(slider_min_label);
                                Canvas.SetTop(slider_min_label, 260);
                                Canvas.SetLeft(slider_min_label, 10);

                                generalCanvas.Children.Add(slider_max_label);
                                Canvas.SetTop(slider_max_label, 260);
                                Canvas.SetLeft(slider_max_label, 210);
                                
                                //Outlook canvas...
                                Canvas outlookCanvas = new Canvas();

                                outlookCanvas.Children.Add(outlookCheckbox);
                                Canvas.SetTop(outlookCheckbox, 10);
                                Canvas.SetLeft(outlookCheckbox, 10);

                                outlookCanvas.Children.Add(outlookTreeViewLabel);
                                Canvas.SetTop(outlookTreeViewLabel, 40);
                                Canvas.SetLeft(outlookTreeViewLabel, 10);

                                outlookCanvas.Children.Add(outlookTreeView);
                                Canvas.SetTop(outlookTreeView, 60);
                                Canvas.SetLeft(outlookTreeView, 10);


                                //Tab control...
                                TabControl tabControl = new TabControl()
                                {
                                    Width = 435,
                                    Height = 340,
                                };

                                tabControl.SelectionChanged += delegate(object sender, SelectionChangedEventArgs e)
                                {
                                    if (OUTLOOK.Equals(((TabItem)tabControl.SelectedItem).Name) &&
                                        outlookCheckbox.IsChecked.Value &&
                                        outlookTreeView.Items.Count <= 0)
                                    {
                                        populateOutlookTreeView(outlookTreeView, Controller.saved_outlook_folders);
                                    }
                                    outlookTreeViewLabel.Visibility = outlookCheckbox.IsChecked.Value ? Visibility.Visible : Visibility.Hidden;
                                    outlookTreeView.Visibility = outlookCheckbox.IsChecked.Value ? Visibility.Visible : Visibility.Hidden;
                                };

                                outlookCheckbox.Click += delegate
                                {
                                    if (outlookCheckbox.IsChecked.Value && outlookTreeView.Items.Count <= 0)
                                    {
                                        //Populate tree...
                                        populateOutlookTreeView(outlookTreeView, Controller.saved_outlook_folders);
                                    }
                                    outlookTreeViewLabel.Visibility = outlookCheckbox.IsChecked.Value ? Visibility.Visible : Visibility.Hidden;
                                    outlookTreeView.Visibility = outlookCheckbox.IsChecked.Value ? Visibility.Visible : Visibility.Hidden;
                                };

                                if (Controller.isOutlookIntegrationAvailable())
                                {
                                    tabControl.Items.Add(new TabItem()
                                    {
                                        Header = Properties_Resources.General,
                                        Content = generalCanvas,
                                    });

                                    tabControl.Items.Add(new TabItem()
                                    {
                                        Name = OUTLOOK,
                                        Header = Properties_Resources.Outlook,
                                        Content = outlookCanvas,
                                    });

                                    //Display tabs when outlook is available
                                    ContentCanvas.Children.Add(tabControl);
                                    Canvas.SetTop(tabControl, 10);
                                    Canvas.SetLeft(tabControl, 170);
                                }
                                else
                                {
                                    //No tabs when outlook is not available
                                    Header = Properties_Resources.Settings;

                                    ContentCanvas.Children.Add(generalCanvas);
                                    Canvas.SetTop(generalCanvas, 60);
                                    Canvas.SetLeft(generalCanvas, 175);
                                }

                                // Actions.
                                Controller.UpdateAddProjectButtonEvent += delegate(bool button_enabled)
                                {
                                    Dispatcher.BeginInvoke((Action)delegate
                                    {
                                        save_button.IsEnabled = button_enabled;
                                    });
                                };

                                cancel_button.Click += delegate
                                {
                                    Controller.PageCancelled();
                                };

                                save_button.Click += delegate
                                {
                                    bool outlookEnabled = false;
                                    List<string> outlookFolders = new List<string>();
                                    if (Controller.isOutlookIntegrationAvailable())
                                    {
                                        outlookEnabled = outlookCheckbox.IsChecked.Value;
                                        if (outlookTreeView.Items.Count <= 0)
                                        {
                                            //TreeView was never initialized so just keep the old values
                                            outlookFolders.AddRange(Controller.saved_outlook_folders);
                                        }
                                        else
                                        {
                                            getSelectedOutlookFolders(outlookTreeView.Items, outlookFolders);
                                        }

                                        if (outlookFolders.Count <= 0)
                                        {
                                            outlookEnabled = false;
                                        }
                                    }

                                    if (!String.IsNullOrEmpty(password_box.Password))
                                    {
                                        // Show wait cursor
                                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

                                        // Try to find the CMIS server (asynchronously)
                                        GetRepositoriesFuzzyDelegate dlgt =
                                            new GetRepositoriesFuzzyDelegate(CmisUtils.GetRepositoriesFuzzy);
                                        IAsyncResult ar = dlgt.BeginInvoke(Controller.saved_address, Controller.saved_user,
                                            password_box.Password, null, null);
                                        while (!ar.AsyncWaitHandle.WaitOne(100))
                                        {
                                            System.Windows.Forms.Application.DoEvents();
                                        }
                                        Tuple<CmisServer, Exception> result = dlgt.EndInvoke(ar);
                                        CmisServer cmisServer = result.Item1;

                                        Controller.repositories = cmisServer != null ? cmisServer.Repositories : null;

                                        // Hide wait cursor
                                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;

                                        if (Controller.repositories == null)
                                        {
                                            // Could not retrieve repositories list from server, show warning.
                                            string warning = "";
                                            string message = result.Item2.Message;
                                            Exception e = result.Item2;
                                            if (e is PermissionDeniedException)
                                            {
                                                warning = Properties_Resources.LoginFailedForbidden;
                                            }
                                            else if (e is ServerNotFoundException)
                                            {
                                                warning = Properties_Resources.ConnectFailure;
                                            }
                                            else if (e.Message == "SendFailure" && cmisServer.Url.Scheme.StartsWith("https"))
                                            {
                                                warning = Properties_Resources.SendFailureHttps;
                                            }
                                            else if (e.Message == "TrustFailure")
                                            {
                                                warning = Properties_Resources.TrustFailure;
                                            }
                                            else
                                            {
                                                warning = message + Environment.NewLine + Properties_Resources.Sorry;
                                            }
                                            authentication_error_label.Text = warning;
                                            authentication_error_label.Visibility = Visibility.Visible;
                                        }
                                        else
                                        {
                                            // Continue to next step, which is choosing a particular folder.
                                            Controller.SettingsPageCompleted(password_box.Password, slider.PollInterval, outlookEnabled, outlookFolders);
                                        }

                                    }
                                    else
                                    {
                                        Controller.SettingsPageCompleted(null, slider.PollInterval, outlookEnabled, outlookFolders);
                                    }

                                };
                                break;
                            }
                        #endregion
                    
                    }

                    ShowAll();
                    Logger.Debug("Exiting ChangePageEvent.");
                });
            };
            this.Closing += delegate
            {
                Controller.PageCancelled();
            };

            Controller.PageCancelled();
            Logger.Debug("Exiting constructor.");
        }

        private static List<OutlookFolder> getOutlookFolderTree()
        {
            OutlookSession outlookSession = new OutlookSession();
            return outlookSession.getFolderTree();
        }

        private static void populateOutlookTreeView(TreeView outlookTreeView)
        {
            populateOutlookTreeView(outlookTreeView, new List<string>(0));
        }

        private static void populateOutlookTreeView(TreeView outlookTreeView, List<string> selectedFolders)
        {
            // Show wait cursor
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

            // Get outlook folders (asynchronously)
            GetOutlookFolderTreeDelegate outlookFolderTreeDelegate =
                new GetOutlookFolderTreeDelegate(getOutlookFolderTree);
            IAsyncResult asyncResult = outlookFolderTreeDelegate.BeginInvoke(null, null);
            while (!asyncResult.AsyncWaitHandle.WaitOne()) //TODO: Do we want a timeout here or wait indefinately
            {
                System.Windows.Forms.Application.DoEvents();
            }
            List<OutlookFolder> outlookFolderTree = outlookFolderTreeDelegate.EndInvoke(asyncResult);
            populateOutlookTreeView(outlookTreeView.Items, outlookFolderTree, selectedFolders);

            // Hide wait cursor
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        }

        private static void populateOutlookTreeView(ItemCollection treeViewItems, List<OutlookFolder> folderTree, List<string> selectedFolders)
        {
            foreach(OutlookFolder outlookFolder in folderTree)
            {
                CheckBox checkBox = new CheckBox()
                {
                     Content = outlookFolder.name,
                     IsChecked = selectedFolders.Contains(outlookFolder.folderPath),
                };

                TreeViewItem treeViewItem = new TreeViewItem()
                {
                    Tag = outlookFolder.folderPath,
                    Header = checkBox,
                    IsExpanded = true,
                };

                checkBox.Click += delegate
                {
                    checkOutlookTreeViewItems(treeViewItem.Items, checkBox.IsChecked.Value);
                };

                populateOutlookTreeView(treeViewItem.Items, outlookFolder.children, selectedFolders);
                treeViewItems.Add(treeViewItem); 
            }
        }

        private static void checkOutlookTreeViewItems(ItemCollection treeViewItems, bool isChecked)
        {
            foreach (object treeViewItemObject in treeViewItems)
            {
                TreeViewItem treeViewItem = (TreeViewItem)treeViewItemObject;
                CheckBox checkBox = (CheckBox)treeViewItem.Header;
                checkBox.IsChecked = isChecked;
                 
                //Recursive...
                checkOutlookTreeViewItems(treeViewItem.Items, isChecked);
            }
        }

        private static void getSelectedOutlookFolders(ItemCollection treeViewItems, List<string> outlookFolders)
        {
            foreach (object treeViewItemObject in treeViewItems)
            {
                TreeViewItem treeViewItem = (TreeViewItem)treeViewItemObject;
                CheckBox checkBox = (CheckBox)treeViewItem.Header;
                if (checkBox.IsChecked.Value)
                {
                    outlookFolders.Add((string)treeViewItem.Tag);
                }

                //Recursive...
                getSelectedOutlookFolders(treeViewItem.Items, outlookFolders);
            }
        }

        private List<string> getIgnoredFolder(Folder f)
        {
            List<string> result = new List<string>();
            if (f.IsIgnored)
            {
                result.Add(f.Path);
            }
            else
            {
                foreach (Folder child in f.SubFolder)
                    result.AddRange(getIgnoredFolder(child));
            }
            return result;
        }

        private List<string> getSelectedFolder(Folder f)
        {
            List<string> result = new List<string>();
            if (f.Selected == true)
                result.Add(f.Path);
            else
            {
                foreach (Folder child in f.SubFolder)
                    result.AddRange(getSelectedFolder(child));
            }
            return result;
        }
    }

    /// <summary>
    /// Stores the metadata of an item in the folder selection dialog.
    /// </summary>
    public class SelectionTreeItem
    {
        /// <summary>
        /// Whether this item's children have been loaded yet.
        /// </summary>
        public bool childrenLoaded = false;

        /// <summary>
        /// Address of the repository.
        /// Only necessary for repository root nodes.
        /// </summary>
        public string repository;

        /// <summary>
        /// Full path to the item.
        /// </summary>
        public string fullPath;

        /// <summary>
        /// Constructor.
        /// </summary>
        public SelectionTreeItem(string repository, string fullPath)
        {
            this.repository = repository;
            this.fullPath = fullPath;
        }
    }
}

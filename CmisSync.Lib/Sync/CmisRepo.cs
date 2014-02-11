//   CmisSync, a CMIS synchronization tool.
//   Copyright (C) 2012  Nicolas Raoul <nicolas.raoul@aegif.jp>
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

namespace CmisSync.Lib.Sync
{

    public partial class CmisRepo : RepoBase
    {
        /// <summary>
        /// Remote folder to synchronize.
        /// </summary>
        private SynchronizedFolder synchronizedFolder;


        /// <summary>
        /// Track whether <c>Dispose</c> has been called.
        /// </summary>
        private bool disposed = false;


        /// <summary>
        /// Constructor.
        /// </summary>
        public CmisRepo(RepoInfo repoInfo, IActivityListener activityListener)
            : base(repoInfo, activityListener)
        {
            this.synchronizedFolder = new SynchronizedFolder(repoInfo, this);
            Logger.Info(synchronizedFolder);
        }


        /// <summary>
        /// Dispose pattern implementation.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    this.synchronizedFolder.Dispose();
                }
                this.disposed = true;
            }
            base.Dispose(disposing);
        }


        /// <summary>
        /// Synchronize.
        /// The synchronization is performed in the background, so that the UI stays usable.
        /// </summary>
        public override void SyncInBackground(bool syncFull)
        {
            if (this.synchronizedFolder != null)// Because it is sometimes called before the object's constructor has completed.
            {
                if (this.Status == SyncStatus.Idle)
                {
                    this.synchronizedFolder.SyncInBackground(syncFull);
                }
                else
                {
                    Logger.Info("Sync skipped. Status=" + this.Status);
                }
            }
        }


        /// <summary>
        /// Synchronize.
        /// The synchronization is performed in the background, so that the UI stays usable.
        /// </summary>
        public override void SyncInBackground()
        {
            SyncInBackground(true);
        }

        /// <summary>
        /// Update repository settings.
        /// </summary>
        public override void UpdateSettings(string password, int pollInterval, bool outlookEnabled, string[] outlookFolders)
        {
            base.UpdateSettings(password, pollInterval, outlookEnabled, outlookFolders);
            synchronizedFolder.CancelSync();
            synchronizedFolder.Dispose();
            synchronizedFolder = new SynchronizedFolder(RepoInfo, this);
            Logger.Info("Updating sync settings. Restarting sync.");
            SyncInBackground();
        }

        /// <summary>
        /// Cancel the currently running sync.  Returns once the current blocking operation completes.
        /// </summary>
        public override void CancelSync()
        {
            synchronizedFolder.CancelSync();
        }

        /// <summary>
        /// Size of the synchronized folder in bytes.
        /// Obtained by adding the individual sizes of all files, recursively.
        /// </summary>
        public override double Size
        {
            get
            {
                return 1234567; // TODO
            }
        }
    }
}

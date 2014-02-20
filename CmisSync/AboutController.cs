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
using System.Net;
using System.Threading;

using CmisSync.Lib;
using System.Reflection;

namespace CmisSync {

    /// <summary>
    /// Controller for the About dialog.
    /// </summary>
    public class AboutController {

        /// <summary>
        /// Show window event.
        /// </summary>
        public event Action ShowWindowEvent = delegate { };

        /// <summary>
        /// HIde window event.
        /// </summary>
        public event Action HideWindowEvent = delegate { };

        /// <summary>
        /// Website URL.
        /// </summary>
        public readonly string WebsiteLinkAddress       = "http://www.oris4.com";

        /// <summary>
        /// Credits link URL.
        /// </summary>
        public readonly string CreditsLinkAddress = "https://raw.github.com/keithwharrison/Oris4Sync/master/legal/AUTHORS.txt";

        /// <summary>
        /// Report problem URL.
        /// </summary>
        public readonly string ReportProblemLinkAddress = "http://www.github.com/keithwharrison/Oris4Sync/issues";


        /// <summary>
        /// Constructor.
        /// </summary>
        public AboutController()
        {
            Program.Controller.ShowAboutWindowEvent += delegate
            {
                ShowWindowEvent();
            };
        }


        /// <summary>
        /// Get the CmisSync version.
        /// </summary>
        public string RunningVersion {
            get {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        /// <summary>
        /// Closing the dialog.
        /// </summary>
        public void WindowClosed ()
        {
            HideWindowEvent ();
        }

    }
}

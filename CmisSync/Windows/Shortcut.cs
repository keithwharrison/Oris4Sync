using log4net;
using System;
using System.Runtime.InteropServices;

namespace CmisSync
{

    /// <summary>
    /// Create a Windows shortcut for CmisSync.
    /// </summary>
    public static class Shortcut
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(Shortcut));

        /// <summary>
        /// Create shortcut.
        /// </summary>
        public static void Create(string target_path, string file_path)
        {
            try
            {
                Type t = Type.GetTypeFromCLSID(new Guid("72C24DD5-D70A-438B-8A42-98424B88AFB8")); //Windows Script Host Shell Object
                dynamic shell = Activator.CreateInstance(t);
                try
                {
                    var lnk = shell.CreateShortcut(file_path);
                    try
                    {
                        lnk.TargetPath = target_path;
                        lnk.Save();
                    }
                    finally
                    {
                        Marshal.FinalReleaseComObject(lnk);
                    }
                }
                finally
                {
                    Marshal.FinalReleaseComObject(shell);
                }
            }
            catch (Exception e)
            {
                Logger.Error(String.Format("Could not create shortcut: {0} -> {1}", file_path, target_path), e);
            }
        }

        /// <summary>
        /// Create shortcut.
        /// </summary>
        public static void Create(string target_path, string file_path, string icofile, int icoidx)
        {
            try
            {
                Type t = Type.GetTypeFromCLSID(new Guid("72C24DD5-D70A-438B-8A42-98424B88AFB8")); //Windows Script Host Shell Object
                dynamic shell = Activator.CreateInstance(t);
                try
                {
                    var lnk = shell.CreateShortcut(file_path);
                    try
                    {
                        lnk.TargetPath = target_path;
                        lnk.IconLocation = icofile + ", " + icoidx;
                        lnk.Save();
                    }
                    finally
                    {
                        Marshal.FinalReleaseComObject(lnk);
                    }
                }
                finally
                {
                    Marshal.FinalReleaseComObject(shell);
                }
            }
            catch (Exception e)
            {
                Logger.Error(String.Format("Could not create shortcut: {0} -> {1} (icon: {2}, {3})", file_path, target_path, icofile, icoidx), e);
            }
        }
    }
}

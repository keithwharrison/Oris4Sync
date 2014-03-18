using DotCMIS;
using DotCMIS.Client;
using DotCMIS.Client.Impl;
using DotCMIS.Exceptions;
using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CmisSync.Lib.Cmis
{
    /// <summary>
    /// Data object representing a CMIS server.
    /// </summary>
    public class CmisServer
    {
        /// <summary>
        /// URL of the CMIS server.
        /// </summary>
        public Uri Url { get; private set; }

        /// <summary>
        /// Repositories contained in the CMIS server.
        /// </summary>
        public Dictionary<string, string> Repositories { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        public CmisServer(Uri url, Dictionary<string, string> repositories)
        {
            Url = url;
            Repositories = repositories;
        }
    }


    /// <summary>
    /// Useful CMIS methods.
    /// </summary>
    public static class CmisUtils
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(CmisUtils));

        
        /// <summary>
        /// Try to find the CMIS server associated to any URL.
        /// Users can provide the URL of the web interface, and we have to return the CMIS URL
        /// Returns the list of repositories as well.
        /// </summary>
        static public Tuple<CmisServer, Exception> GetRepositoriesFuzzy(Uri url, string user, string password)
        {
            Dictionary<string, string> repositories = null;
            Exception firstException = null;

            // Try the given URL, maybe user directly entered the CMIS AtomPub endpoint URL.
            try
            {
                repositories = GetRepositories(url, user, password);
            }
            catch (CmisRuntimeException e)
            {
                if (e.Message == "ConnectFailure")
                    return new Tuple<CmisServer, Exception>(new CmisServer(url, null), new ServerNotFoundException(e.Message, e));
                firstException = e;

            }
            catch (Exception e)
            {
                // Save first Exception and try other possibilities.
                firstException = e;
            }
            if (repositories != null)
            {
                // Found!
                return new Tuple<CmisServer, Exception>(new CmisServer(url, repositories), null);
            }

            // Extract protocol and server name or IP address
            string prefix = url.GetLeftPart(UriPartial.Authority);

            // See https://github.com/nicolas-raoul/CmisSync/wiki/What-address for the list of ECM products prefixes
            // Please send us requests to support more CMIS servers: https://github.com/nicolas-raoul/CmisSync/issues
            string[] suffixes = {
                "/cmis/atom",
                /* We don't need all these for oris4
                "/alfresco/cmisatom",
                "/alfresco/service/cmis",
                "/cmis/resources/",
                "/emc-cmis-ea/resources/",
                "/emc-cmis-weblogic/resources/",
                "/emc-cmis-wls/resources/",
                "/emc-cmis-was61/resources/",
                "/emc-cmis-wls1030/resources/",
                "/xcmis/rest/cmisatom",
                "/files/basic/cmis/my/servicedoc",
                "/p8cmis/resources/Service",
                "/_vti_bin/cmis/rest?getRepositories",
                "/Nemaki/atom/bedroom",
                "/nuxeo/atom/cmis"
                */
            };
            string bestUrl = null;
            // Try all suffixes
            for (int i=0; i < suffixes.Length; i++)
            {
                string fuzzyUrl = prefix + suffixes[i];
                Logger.Info("Sync | Trying with " + fuzzyUrl);
                try
                {
                    repositories = GetRepositories(new Uri(fuzzyUrl), user, password);
                }
                catch (CmisPermissionDeniedException e)
                {
                    firstException = new PermissionDeniedException(e.Message, e);
                    bestUrl = fuzzyUrl;
                }
                catch (Exception e)
                {
                    // Do nothing, try other possibilities.
                    Logger.Info(e.Message);
                }
                if (repositories != null)
                {
                    // Found!
                    return new Tuple<CmisServer, Exception>( new CmisServer(new Uri(fuzzyUrl), repositories), null);
                }
            }

            // Not found. Return also the first exception to inform the user correctly
            return new Tuple<CmisServer,Exception>(new CmisServer(bestUrl==null?url:new Uri(bestUrl), null), firstException);
        }


        /// <summary>
        /// Get the list of repositories of a CMIS server
        /// Each item contains id + 
        /// </summary>
        /// <returns>The list of repositories. Each item contains the identifier and the human-readable name of the repository.</returns>
        static public Dictionary<string,string> GetRepositories(Uri url, string user, string password)
        {
            Dictionary<string,string> result = new Dictionary<string,string>();

            // If no URL was provided, return empty result.
            if (null == url)
            {
                return result;
            }

            IList<IRepository> repositories;
            try
            {
                repositories = Auth.Auth.GetCmisRepositories(url, user, password);
            }
            catch (CmisPermissionDeniedException e)
            {
                Logger.Error("CMIS server found, but permission denied. Please check username/password. ", e);
                throw;
            }
            catch (CmisRuntimeException e)
            {
                Logger.Error("No CMIS server at this address, or no connection. ", e);
                throw;
            }
            catch (CmisObjectNotFoundException e)
            {
                Logger.Error("No CMIS server at this address, or no connection. ", e);
                throw;
            }
            catch (CmisConnectionException e)
            {
                Logger.Error("No CMIS server at this address, or no connection. ", e);
                throw;
            }
            catch (CmisInvalidArgumentException e)
            {
                Logger.Error("Invalid URL, maybe Alfresco Cloud? ", e);
                throw;
            }

            // Populate the result list with identifier and name of each repository.
            foreach (IRepository repo in repositories)
            {
                result.Add(repo.Id, repo.Name);
            }
            
            return result;
        }


        /// <summary>
        /// Get the sub-folders of a particular CMIS folder.
        /// </summary>
        /// <returns>Full path of each sub-folder, including leading slash.</returns>
        static public string[] GetSubfolders(string repositoryId, string path,
            string url, string user, string password)
        {
            List<string> result = new List<string>();

            // Connect to the CMIS repository.
            ISession session = Auth.Auth.GetCmisSession(url, user, password, repositoryId);

            // Get the folder.
            IFolder folder;
            try
            {
                folder = (IFolder)session.GetObjectByPath(path);
            }
            catch (Exception ex)
            {
                Logger.Warn(String.Format("CmisUtils | exception when session GetObjectByPath for {0}", path), ex);
                return result.ToArray();
            }

            // Debug the properties count, which allows to check whether a particular CMIS implementation is compliant or not.
            // For instance, IBM Connections is known to send an illegal count.
            Logger.Info("CmisUtils | folder.Properties.Count:" + folder.Properties.Count.ToString());

            // Get the folder's sub-folders.
            IItemEnumerable<ICmisObject> children = folder.GetChildren();

            // Return the full path of each of the sub-folders.
            foreach (var subfolder in children.OfType<IFolder>())
            {
                result.Add(subfolder.Path);
            }
            return result.ToArray();
        }


        /// <summary>
        /// Folder tree.
        /// </summary>
        public class FolderTree
        {
            /// <summary>
            /// Children.
            /// </summary>
            public List<FolderTree> children = new List<FolderTree>();
            /// <summary>
            /// Folder path.
            /// </summary>
            public string path;
            /// <summary>
            /// Folder name.
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// Constructor.
            /// </summary>
            public FolderTree(IList<ITree<IFileableCmisObject>> trees, IFolder folder)
            {
                this.path = folder.Path;
                this.Name = folder.Name;
                if(trees != null)
                    foreach (ITree<IFileableCmisObject> tree in trees)
                    {
                        Folder f = tree.Item as Folder;
                        if(f!=null)
                            this.children.Add(new FolderTree(tree.Children, f));
                    }
            }
        }

        /// <summary>
        /// Get the sub-folders of a particular CMIS folder.
        /// </summary>
        /// <returns>Full path of each sub-folder, including leading slash.</returns>
        static public FolderTree GetSubfolderTree(string repositoryId, string path,
            string url, string user, string password, int depth)
        {
            // Connect to the CMIS repository.
            ISession session = Auth.Auth.GetCmisSession(url, user, password, repositoryId);

            // Get the folder.
            IFolder folder;
            try
            {
                folder = (IFolder)session.GetObjectByPath(path);
            }
            catch (Exception ex)
            {
                Logger.Warn(String.Format("CmisUtils | exception when session GetObjectByPath for {0}", path), ex);
                throw;
            }

            // Debug the properties count, which allows to check whether a particular CMIS implementation is compliant or not.
            // For instance, IBM Connections is known to send an illegal count.
            Logger.Info("CmisUtils | folder.Properties.Count:" + folder.Properties.Count.ToString());
            try
            {
                IList<ITree<IFileableCmisObject>> trees = folder.GetFolderTree(depth);
                return new FolderTree(trees, folder);
            }
            catch (Exception e)
            {
                Logger.Info("CmisUtils getSubFolderTree | Exception " + e.Message, e);
                throw;
            }
        }


        /// <summary>
        /// Guess the web address where files can be seen using a browser.
        /// Not bulletproof. It depends on the server, and on some servers there is no web UI at all.
        /// </summary>
        static public string GetBrowsableURL(RepoInfo repo)
        {
            if (null == repo)
            {
                throw new ArgumentNullException("repo");
            }
            
            // Case of Alfresco.
            string suffix1 = "alfresco/cmisatom";
            string suffix2 = "alfresco/service/cmis";
            if (repo.Address.AbsoluteUri.EndsWith(suffix1) || repo.Address.AbsoluteUri.EndsWith(suffix2))
            {
                // Detect suffix length.
                int suffixLength = 0;
                if (repo.Address.AbsoluteUri.EndsWith(suffix1))
                    suffixLength = suffix1.Length;
                if (repo.Address.AbsoluteUri.EndsWith(suffix2))
                    suffixLength = suffix2.Length;

                string root = repo.Address.AbsoluteUri.Substring(0, repo.Address.AbsoluteUri.Length - suffixLength);
                if (repo.RemotePath.StartsWith("/Sites"))
                {
                    // Case of Alfresco Share.

                    // Example RemotePath: /Sites/thesite
                    // Result: http://server/share/page/site/thesite/documentlibrary
                    // Example RemotePath: /Sites/thesite/documentLibrary/somefolder/anotherfolder
                    // Result: http://server/share/page/site/thesite/documentlibrary#filter=path|%2Fsomefolder%2Fanotherfolder
                    // Example RemotePath: /Sites/s1/documentLibrary/éß和ệ
                    // Result: http://server/share/page/site/s1/documentlibrary#filter=path|%2F%25E9%25DF%25u548C%25u1EC7
                    // Example RemotePath: /Sites/s1/documentLibrary/a#bc/éß和ệ
                    // Result: http://server/share/page/site/thesite/documentlibrary#filter=path%7C%2Fa%2523bc%2F%25E9%25DF%25u548C%25u1EC7%7C

                    string path = repo.RemotePath.Substring("/Sites/".Length);
                    if (path.Contains("documentLibrary"))
                    {
                        int firstSlashPosition = path.IndexOf('/');
                        string siteName = path.Substring(0, firstSlashPosition);
                        string pathWithinSite = path.Substring(firstSlashPosition + "/documentLibrary".Length);
                        string escapedPathWithinSite = HttpUtility.UrlEncode(pathWithinSite);
                        string reescapedPathWithinSite = HttpUtility.UrlEncode(escapedPathWithinSite);
                        string sharePath = reescapedPathWithinSite.Replace("%252f", "%2F");
                        return root + "share/page/site/" + siteName + "/documentlibrary#filter=path|" + sharePath;
                    }
                    else
                    {
                        // Site name only.
                        return root + "share/page/site/" + path + "/documentlibrary";
                    }
                }
                else
                {
                    // Case of Alfresco Web Client. Difficult to build a direct URL, so return root.
                    return root;
                }
            }
            else
            {
                // Another server was detected, try to open the thinclient url, otherwise try to open the repo path

                try
                {
                    // Connect to the CMIS repository.
                    ISession session = Auth.Auth.GetCmisSession(repo.Address.ToString(), repo.User, repo.Password.ToString(), repo.RepoID);

                    if (session.RepositoryInfo.ThinClientUri == null
                        || String.IsNullOrEmpty(session.RepositoryInfo.ThinClientUri.ToString()))
                    {
                        Logger.Error("CmisUtils GetBrowsableURL | Repository does not implement ThinClientUri: " + repo.Address.AbsoluteUri);
                        return repo.Address.AbsoluteUri + repo.RemotePath;
                    }
                    else
                    {
                        // Return CmisServer-provided thin URL.
                        return session.RepositoryInfo.ThinClientUri.ToString();
                    }
                }
                catch (Exception e)
                {
                    Logger.Error("CmisUtils GetBrowsableURL | Exception " + e.Message, e);
                    // Server down or authentication problem, no way to know the right URL, so just open server.
                    return repo.Address.AbsoluteUri + repo.RemotePath;
                }
            }
        }
    }
}

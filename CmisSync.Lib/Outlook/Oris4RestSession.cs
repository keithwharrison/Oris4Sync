using CmisSync.Lib.Cmis;
using RestSharp;
using System.Collections.Generic;

namespace CmisSync.Lib.Outlook
{
    public class Oris4RestSession
    {
        private RestClient client = null;
        private OAuth oAuth = null;
        private string emailAddress = null;
        private string registeredClient = null;

        public Oris4RestSession(string baseUrl)
            : this(baseUrl, null)
        {

        }

        public Oris4RestSession(string baseUrl, OAuth oAuth)
        {
            this.client = new RestClient(baseUrl);
            this.oAuth = oAuth;

            if (oAuth != null)
            {
                client.Authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(oAuth.value, oAuth.tokenType);
            }
        }

        public void login(string username, string password)
        {
            oAuth = Oris4RestService.login(client, username, password);
            if (oAuth != null)
            {
                client.Authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(oAuth.value, oAuth.tokenType);
                this.emailAddress = username;
            }
        }

        public Email getEmail(long emailKey, bool linkedEntities, int offset, int pageSize)
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            return Oris4RestService.getEmail(client, emailKey, linkedEntities, offset, pageSize);
        }

        public void deleteEmail(string emailHash)
        {
            if (oAuth == null || string.IsNullOrWhiteSpace(emailAddress))
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            if (string.IsNullOrWhiteSpace(registeredClient))
            {
                throw new PermissionDeniedException("You must register outlook before performing this action");
            }

            Oris4RestService.deleteEmail(client, registeredClient, emailAddress, emailHash);
        }

        public List<Email> listEmail(long folderKey, int offset, int pageSize)
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            return Oris4RestService.listEmail(client, folderKey, offset, pageSize);
        }

        public void putRegisteredClient(string accountId)
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            Oris4RestService.putRegisteredClient(client, accountId);
            this.registeredClient = accountId;
        }

        public string getRegisteredClient()
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            registeredClient = Oris4RestService.getRegisteredClient(client);
            return registeredClient;
        }

        public List<Email> insertEmail(List<Email> emailList)
        {
            if (oAuth == null || string.IsNullOrWhiteSpace(emailAddress))
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            if (string.IsNullOrWhiteSpace(registeredClient))
            {
                throw new PermissionDeniedException("You must register outlook before performing this action");
            }

            return Oris4RestService.insertEmail(client, registeredClient, emailAddress, emailList);
        }

        public Dictionary<string, long> insertEmailBatch(List<Email> emailList)
        {
            if (oAuth == null || string.IsNullOrWhiteSpace(emailAddress))
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            if (string.IsNullOrWhiteSpace(registeredClient))
            {
                throw new PermissionDeniedException("You must register outlook before performing this action");
            }

            return Oris4RestService.insertEmailBatch(client, registeredClient, emailAddress, emailList);
        }

        public string insertAttachment(EmailAttachment emailAttachment, byte[] data)
        {
            if (oAuth == null || string.IsNullOrWhiteSpace(emailAddress))
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            if (string.IsNullOrWhiteSpace(registeredClient))
            {
                throw new PermissionDeniedException("You must register outlook before performing this action");
            }

            return Oris4RestService.insertAttachment(client, registeredClient, emailAddress, emailAttachment, data);
        }

        public Oris4Folder getRootOutlookFolder()
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            List<Oris4Folder> rootFolderList = Oris4RestService.listRootFolders(client);

            Oris4Folder rootOutlookFolder = null;
            foreach (Oris4Folder folder in rootFolderList)
            {
                if ("emailingestion".Equals(folder.integrationId))
                {
                    rootOutlookFolder = folder;
                }
            }

            return rootOutlookFolder;
        }

        public List<Oris4Folder> listSubFolders(long folderKey)
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            return Oris4RestService.listSubFolders(client, folderKey);
        }
    }
}

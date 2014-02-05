using CmisSync.Lib.Cmis;
using RestSharp;
using System.Collections.Generic;

namespace CmisSync.Lib.Outlook
{
    public class Oris4RestSession
    {
        private RestClient client = null;
        private OAuth oAuth = null;

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
            oAuth = Oris4RestService.Instance.login(client, username, password);
            if (oAuth != null)
            {
                client.Authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(oAuth.value, oAuth.tokenType);
            }
        }

        public Email getEmail(int emailKey, bool linkedEntities, int offset, int pageSize)
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            return Oris4RestService.Instance.getEmail(client, emailKey, linkedEntities, offset, pageSize);
        }

        public void deleteEmail(string accountId, string emailAddress, string emailHash)
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            Oris4RestService.Instance.deleteEmail(client, accountId, emailAddress, emailHash);
        }

        public List<Email> listEmail(int folderKey, int offset, int pageSize)
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            return Oris4RestService.Instance.listEmail(client, folderKey, offset, pageSize);
        }

        public void putRegisteredClient(string accountId)
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            Oris4RestService.Instance.putRegisteredClient(client, accountId);
        }

        public string getRegisteredClient()
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            return Oris4RestService.Instance.getRegisteredClient(client);
        }

        public List<Email> insertEmail(string accountId, string emailAddress, List<Email> emailList)
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            return Oris4RestService.Instance.insertEmail(client, accountId, emailAddress, emailList);
        }

        public string insertAttachment(string accountId, string emailAddress, EmailAttachment emailAttachment,
            byte[] data, string contentType)
        {
            if (oAuth == null)
            {
                throw new PermissionDeniedException("You must login before performing this action");
            }

            return Oris4RestService.Instance.insertAttachment(client, accountId, emailAddress, emailAttachment, data, contentType);
        }
    }
}

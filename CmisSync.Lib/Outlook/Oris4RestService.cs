using log4net;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Net;

namespace CmisSync.Lib.Outlook
{
    public class Oris4RestService
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(Oris4RestService));

        private static readonly string URI_OAUTH_LOGIN = "oauth/token";

        private static readonly string URI_EMAIL_GET = "webServices/rest/v3/email/{key}";
        private static readonly string URI_EMAIL_DELETE = "webServices/rest/v3/email/{emailHash}";
        private static readonly string URI_EMAIL_LIST_GET = "webServices/rest/v3/email/list";
        private static readonly string URI_EMAIL_REGISTERED_CLIENT_GET = "webServices/rest/v3/email/registeredClient";
        private static readonly string URI_EMAIL_REGISTERED_CLIENT_PUT = "webServices/rest/v3/email/registeredClient";
        private static readonly string URI_EMAIL_POST = "webServices/rest/v3/email/";
        private static readonly string URI_EMAIL_ATTACHMENT_POST = "webServices/rest/v3/email/attachment";

        private static readonly string CLIENT_TYPE_OUTLOOK = "outlook";

        private static Oris4RestService instance;

        public static Oris4RestService Instance
        {
            get
            {
                if (instance == null) instance = new Oris4RestService();
                return instance;
            }
        }

        private Oris4RestService()
        {
        }

        private IRestRequest getRestRequest(string uri, Method method)
        {
            IRestRequest request = new RestRequest(uri, method);
            request.JsonSerializer = new JsonSerializer();
            return request;
        }

        public OAuth login(RestClient client, string username, string password)
        {
            string consumerKey = Config.Instance.ConsumerKey;
            string consumerSecret = Config.Instance.ConsumerSecret;
            string grantType = Config.Instance.GrantType;

            IRestRequest request = getRestRequest(URI_OAUTH_LOGIN, Method.POST);
            request.AddParameter("client_id", consumerKey);
            request.AddParameter("client_secret", consumerSecret);
            request.AddParameter("grant_type", grantType);
            request.AddParameter("username", username);
            request.AddParameter("password", password);

            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<OAuth> response = client.Execute<OAuth>(request);

            checkResponseStatus(response);

            return response.Data;
        }

        public Email getEmail(RestClient client, int emailKey, bool linkedEntities, int offset, int pageSize)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_GET, Method.GET);
            request.AddUrlSegment("key", emailKey.ToString());

            request.AddParameter("getLinkedEntities", linkedEntities.ToString());
            request.AddParameter("offset", offset.ToString());
            request.AddParameter("pageSize", pageSize.ToString());


            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<Email> response = client.Execute<Email>(request);

            checkResponseStatus(response);

            return response.Data;
        }

        public void deleteEmail(RestClient client, string accountId, string emailAddress, string emailHash)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_DELETE, Method.DELETE);
            request.AddHeader("Client-GUID", accountId);
            request.AddHeader("Email-Address", emailAddress);

            request.AddUrlSegment("emailHash", emailHash);

            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<Email> response = client.Execute<Email>(request);

            checkResponseStatus(response, HttpStatusCode.NoContent);
        }

        public List<Email> listEmail(RestClient client, int folderKey, int offset, int pageSize)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_LIST_GET, Method.GET);

            request.AddParameter("folderKey", folderKey.ToString());
            request.AddParameter("offset", offset.ToString());
            request.AddParameter("pageSize", pageSize.ToString());


            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<List<Email>> response = client.Execute<List<Email>>(request);

            checkResponseStatus(response);

            return response.Data;
        }

        public void putRegisteredClient(RestClient client, string accountId)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_REGISTERED_CLIENT_PUT, Method.PUT);
            request.AddHeader("Client-Type", CLIENT_TYPE_OUTLOOK);
            request.AddHeader("Client-GUID", accountId);

            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse response = client.Execute(request);

            checkResponseStatus(response, HttpStatusCode.NoContent);
        }

        public string getRegisteredClient(RestClient client)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_REGISTERED_CLIENT_GET, Method.GET);
            request.AddHeader("Client-Type", CLIENT_TYPE_OUTLOOK);

            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse response = client.Execute(request);

            checkResponseStatus(response, new List<HttpStatusCode>() { HttpStatusCode.OK, HttpStatusCode.NoContent });

            string accountId = response.Content;
            if (string.IsNullOrWhiteSpace(accountId))
            {
                accountId = string.Empty;
            }

            return accountId.Trim('"');
        }

        public List<Email> insertEmail(RestClient client, string accountId, string emailAddress, List<Email> emailList)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_POST, Method.POST);
            request.AddHeader("Client-GUID", accountId);
            request.AddHeader("Email-Address", emailAddress);

            request.RequestFormat = DataFormat.Json;
            request.AddBody(emailList);

            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<List<Email>> response = client.Execute<List<Email>>(request);

            checkResponseStatus(response, new List<HttpStatusCode>() { HttpStatusCode.OK, HttpStatusCode.ExpectationFailed });

            switch (response.StatusCode)
            {
                case HttpStatusCode.OK:
                    Logger.Info("Emails were created...");
                    break;
                case HttpStatusCode.ExpectationFailed:
                    Logger.Info("Some emails created...");
                    break;
            }

            return response.Data;
        }

        public string insertAttachment(RestClient client, string accountId, string emailAddress, EmailAttachment emailAttachment, byte[] data)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_ATTACHMENT_POST, Method.POST);
            request.AddHeader("Client-GUID", accountId);
            request.AddHeader("Email-Address", emailAddress);

            request.AddParameter("jsonEmailAttachment", request.JsonSerializer.Serialize(emailAttachment));

            request.AddFile("data", data, emailAttachment.fileName);

            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse response = client.Execute(request);

            checkResponseStatus(response, new List<HttpStatusCode>() { HttpStatusCode.OK, HttpStatusCode.Created });

            switch (response.StatusCode)
            {
                case HttpStatusCode.OK:
                    Logger.Info("Attachment already exists.");
                    break;
                case HttpStatusCode.Created:
                    Logger.Info("Attachment created.");
                    break;
            }

            return response.Content;
        }

        private void checkResponseStatus(IRestResponse restResponse)
        {
            checkResponseStatus(restResponse, HttpStatusCode.OK);
        }

        private void checkResponseStatus(IRestResponse restResponse, HttpStatusCode statusCode)
        {
            checkResponseStatus(restResponse, new List<HttpStatusCode>() { statusCode });
        }

        private void checkResponseStatus(IRestResponse restResponse, List<HttpStatusCode> expectedStatus)
        {
            //Logger.DebugFormat("StatusCode: {0} {1}", restResponse.StatusCode, restResponse.StatusDescription);
            //Logger.DebugFormat("ResponseStatus: {0}", restResponse.ResponseStatus);
            //Logger.DebugFormat("Content: {0}", restResponse.Content);

            if (expectedStatus.Contains(restResponse.StatusCode))
            {
                return;
            }

            switch ((int)restResponse.StatusCode)
            {
                case (int)HttpStatusCode.Unauthorized:
                case (int)HttpStatusCode.Forbidden:
                case (int)HttpStatusCode.NotFound:
                    throw new CmisSync.Lib.Cmis.PermissionDeniedException("Authentication Failed.");

                case 420:
                    throw new CmisSync.Lib.Cmis.ServerBusyException("Server was busy, try again later.");

                case 423:
                    throw new CmisSync.Lib.Cmis.AccountLockedException("User account is locked.");

                case (int)HttpStatusCode.BadRequest:
                case (int)HttpStatusCode.InternalServerError:
                default:
                    throw new CmisSync.Lib.Cmis.BaseException(String.Format("Error: {0} {1}", restResponse.StatusDescription, restResponse.ResponseStatus));
            }

        }
    }
}

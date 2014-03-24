using CmisSync.Lib.Cmis;
using log4net;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Net;

namespace CmisSync.Lib.Outlook
{
    public static class Oris4RestService
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(Oris4RestService));

        private static readonly string URI_OAUTH_LOGIN = "oauth/token";

        private static readonly string URI_EMAIL_GET = "webServices/rest/v3/email/{key}";
        private static readonly string URI_EMAIL_DELETE = "webServices/rest/v3/email/{emailHash}";
        private static readonly string URI_EMAIL_LIST_GET = "webServices/rest/v3/email/list";
        private static readonly string URI_EMAIL_REGISTERED_CLIENT_GET = "webServices/rest/v3/email/registeredClient";
        private static readonly string URI_EMAIL_REGISTERED_CLIENT_PUT = "webServices/rest/v3/email/registeredClient";
        private static readonly string URI_EMAIL_POST = "webServices/rest/v3/email/";
        private static readonly string URI_EMAIL_BATCH_POST = "webServices/rest/v3/email/batch";
        private static readonly string URI_EMAIL_ATTACHMENT_POST = "webServices/rest/v3/email/attachment";
        private static readonly string URI_FOLDER_ROOT_GET = "webServices/rest/v3/folder/root";
        private static readonly string URI_FOLDER_SUBFOLDERS_GET = "webServices/rest/v3/folder/subfolders/{folderKey}";

        private static readonly string CLIENT_TYPE_OUTLOOK = "outlook";

        private static IRestRequest getRestRequest(string uri, Method method)
        {
            IRestRequest request = new RestRequest(uri, method);
            request.JsonSerializer = new JsonSerializer();
            return request;
        }

        public static OAuth login(RestClient client, string username, string password)
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

            Logger.DebugFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<OAuth> response = client.Execute<OAuth>(request);

            checkResponseStatus(response);

            return response.Data;
        }

        public static Email getEmail(RestClient client, long emailKey, bool linkedEntities, int offset, int pageSize)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_GET, Method.GET);
            request.AddUrlSegment("key", emailKey.ToString());

            request.AddParameter("getLinkedEntities", linkedEntities.ToString());
            request.AddParameter("offset", offset.ToString());
            request.AddParameter("pageSize", pageSize.ToString());


            Logger.DebugFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<Email> response = client.Execute<Email>(request);

            checkResponseStatus(response);

            return response.Data;
        }

        public static void deleteEmail(RestClient client, string accountId, string emailAddress, string emailHash)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_DELETE, Method.DELETE);
            request.AddHeader("Client-GUID", accountId);
            request.AddHeader("Email-Address", emailAddress);

            request.AddUrlSegment("emailHash", emailHash);

            Logger.DebugFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<Email> response = client.Execute<Email>(request);

            checkResponseStatus(response, HttpStatusCode.NoContent);
        }

        public static List<Email> listEmail(RestClient client, long folderKey, int offset, int pageSize)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_LIST_GET, Method.GET);

            request.AddParameter("folderKey", folderKey.ToString());
            request.AddParameter("offset", offset.ToString());
            request.AddParameter("pageSize", pageSize.ToString());


            Logger.DebugFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<List<Email>> response = client.Execute<List<Email>>(request);

            checkResponseStatus(response);

            return response.Data;
        }

        public static void putRegisteredClient(RestClient client, string accountId)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_REGISTERED_CLIENT_PUT, Method.PUT);
            request.AddHeader("Client-Type", CLIENT_TYPE_OUTLOOK);
            request.AddHeader("Client-GUID", accountId);

            Logger.DebugFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse response = client.Execute(request);

            checkResponseStatus(response, HttpStatusCode.NoContent);
        }

        public static string getRegisteredClient(RestClient client)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_REGISTERED_CLIENT_GET, Method.GET);
            request.AddHeader("Client-Type", CLIENT_TYPE_OUTLOOK);

            Logger.DebugFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse response = client.Execute(request);

            checkResponseStatus(response, new List<HttpStatusCode>() { HttpStatusCode.OK, HttpStatusCode.NoContent });

            string accountId = response.Content;
            if (string.IsNullOrWhiteSpace(accountId))
            {
                accountId = string.Empty;
            }

            return accountId.Trim('"');
        }

        public static List<Email> insertEmail(RestClient client, string accountId, string emailAddress, List<Email> emailList)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_POST, Method.POST);
            request.AddHeader("Client-GUID", accountId);
            request.AddHeader("Email-Address", emailAddress);

            request.RequestFormat = DataFormat.Json;
            request.AddBody(emailList);

            Logger.DebugFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<List<Email>> response = client.Execute<List<Email>>(request);

            checkResponseStatus(response, new List<HttpStatusCode>() { HttpStatusCode.OK, HttpStatusCode.ExpectationFailed });

            switch (response.StatusCode)
            {
                case HttpStatusCode.OK:
                    Logger.Debug("Emails were created...");
                    break;
                case HttpStatusCode.ExpectationFailed:
                    Logger.Debug("Some emails created...");
                    break;
            }

            return response.Data;
        }

        public static Dictionary<string, long> insertEmailBatch(RestClient client, string accountId, string emailAddress, List<Email> emailList)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_BATCH_POST, Method.POST);
            request.AddHeader("Client-GUID", accountId);
            request.AddHeader("Email-Address", emailAddress);

            request.RequestFormat = DataFormat.Json;
            request.AddBody(emailList);

            Logger.DebugFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<Dictionary<string, long>> response = client.Execute<Dictionary<string, long>>(request);

            checkResponseStatus(response, HttpStatusCode.OK);

            return response.Data;
        }

        public static string insertAttachment(RestClient client, string accountId, string emailAddress, EmailAttachment emailAttachment, byte[] data)
        {
            IRestRequest request = getRestRequest(URI_EMAIL_ATTACHMENT_POST, Method.POST);
            request.AddHeader("Client-GUID", accountId);
            request.AddHeader("Email-Address", emailAddress);

            request.AddParameter("jsonEmailAttachment", request.JsonSerializer.Serialize(emailAttachment));

            request.AddFile("data", data, emailAttachment.fileName);

            Logger.DebugFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse response = client.Execute(request);

            checkResponseStatus(response, new List<HttpStatusCode>() { HttpStatusCode.OK, HttpStatusCode.Created });

            switch (response.StatusCode)
            {
                case HttpStatusCode.OK:
                    Logger.Debug("Attachment already exists.");
                    break;
                case HttpStatusCode.Created:
                    Logger.Debug("Attachment created.");
                    break;
            }

            return response.Content;
        }

        public static List<Oris4Folder> listRootFolders(RestClient client)
        {
            IRestRequest request = getRestRequest(URI_FOLDER_ROOT_GET, Method.GET);

            request.AddParameter("withSubFolders", true);

            Logger.DebugFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<List<Oris4Folder>> response = client.Execute<List<Oris4Folder>>(request);

            checkResponseStatus(response);

            return response.Data;
        }

        public static List<Oris4Folder> listSubFolders(RestClient client, long folderKey)
        {
            IRestRequest request = getRestRequest(URI_FOLDER_SUBFOLDERS_GET, Method.GET);

            request.AddUrlSegment("folderKey", folderKey.ToString());

            Logger.DebugFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<List<Oris4Folder>> response = client.Execute<List<Oris4Folder>>(request);

            checkResponseStatus(response);

            return response.Data;
        }

        private static void checkResponseStatus(IRestResponse restResponse)
        {
            checkResponseStatus(restResponse, HttpStatusCode.OK);
        }

        private static void checkResponseStatus(IRestResponse restResponse, HttpStatusCode statusCode)
        {
            checkResponseStatus(restResponse, new List<HttpStatusCode>() { statusCode });
        }

        private static void checkResponseStatus(IRestResponse restResponse, List<HttpStatusCode> expectedStatus)
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
                    throw new PermissionDeniedException("Authentication Failed.");

                case 420:
                    throw new ServerBusyException("Server was busy, try again later.");

                case 423:
                    throw new AccountLockedException("User account is locked.");

                case (int)HttpStatusCode.BadRequest:
                case (int)HttpStatusCode.InternalServerError:
                default:
                    throw new BaseException(String.Format("Error: {0} {1}", restResponse.StatusDescription, restResponse.ResponseStatus));
            }

        }
    }
}

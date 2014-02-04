using log4net;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

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

        public static OAuth login(string baseUrl, string username, string password)
        {
            string consumerKey = Config.Instance.ConsumerKey;
            string consumerSecret = Config.Instance.ConsumerSecret;
            string grantType = Config.Instance.GrantType;

            RestClient client = new RestClient(baseUrl);
            IRestRequest request = new RestRequest(URI_OAUTH_LOGIN, Method.POST);
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

        public static void getEmail(string baseUrl, string oauthToken, string oauthTokenType, int emailKey, bool linkedEntities, int offset, int pageSize)
        {
            RestClient client = new RestClient(baseUrl);
            client.Authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(oauthToken, oauthTokenType);
            RestRequest request = new RestRequest(URI_EMAIL_GET, Method.GET);
            request.AddUrlSegment("key", emailKey.ToString());

            request.AddParameter("getLinkedEntities", linkedEntities.ToString());
            request.AddParameter("offset", offset.ToString());
            request.AddParameter("pageSize", pageSize.ToString());


            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<Email> response = client.Execute<Email>(request);

            checkResponseStatus(response);
        }

        public static void deleteEmail(string baseUrl, string oauthToken, string oauthTokenType, string accountId, string emailAddress, string emailHash)
        {
            RestClient client = new RestClient(baseUrl);
            client.Authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(oauthToken, oauthTokenType);
            RestRequest request = new RestRequest(URI_EMAIL_DELETE, Method.DELETE);
            request.AddHeader("Client-GUID", accountId);
            request.AddHeader("Email-Address", emailAddress);

            request.AddUrlSegment("emailHash", emailHash);

            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<Email> response = client.Execute<Email>(request);

            checkResponseStatus(response, HttpStatusCode.NoContent);
        }

        public static void listEmail(string baseUrl, string oauthToken, string oauthTokenType, int folderKey, int offset, int pageSize)
        {

        }

        public static void putRegisteredClient(string baseUrl, string oauthToken, string oauthTokenType, string accountId)
        {
            RestClient client = new RestClient(baseUrl);
            client.Authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(oauthToken, oauthTokenType);
            RestRequest request = new RestRequest(URI_EMAIL_REGISTERED_CLIENT_PUT, Method.PUT);
            request.AddHeader("Client-Type", CLIENT_TYPE_OUTLOOK);
            request.AddHeader("Client-GUID", accountId);

            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse response = client.Execute(request);
            Logger.Info("StatusCode: " + response.StatusCode);
            Logger.Info("ResponseStatus: " + response.ResponseStatus);
            Logger.Info("Content: " + response.Content);

            //TODO: Check response codes...
        }

        public static string getRegisteredClient(string baseUrl, string oauthToken, string oauthTokenType)
        {
            RestClient client = new RestClient(baseUrl);
            client.Authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(oauthToken, oauthTokenType);
            RestRequest request = new RestRequest(URI_EMAIL_REGISTERED_CLIENT_GET, Method.GET);
            request.AddHeader("Client-Type", CLIENT_TYPE_OUTLOOK);

            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse response = client.Execute(request);
            Logger.Info("StatusCode: " + response.StatusCode);
            Logger.Info("ResponseStatus: " + response.ResponseStatus);
            Logger.Info("Content: " + response.Content);

            //TODO: Check response codes...

            string accountId = response.Content;
            if (string.IsNullOrWhiteSpace(accountId))
            {
                throw new Exception("Blah!");
            }

            return accountId.Trim('"');
        }

        public static List<Email> insertEmail(string baseUrl, string oauthToken, string oauthTokenType, string accountId, string emailAddress, List<Email> emailList)
        {
            RestClient client = new RestClient(baseUrl);
            client.Authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(oauthToken, oauthTokenType);
            RestRequest request = new RestRequest(URI_EMAIL_POST, Method.POST);
            request.AddHeader("Client-GUID", accountId);
            request.AddHeader("Email-Address", emailAddress);

            request.RequestFormat = DataFormat.Json;
            request.AddBody(emailList);

            Logger.InfoFormat("Request: {0} {1}", request.Method, request.Resource);
            IRestResponse<List<Email>> response = client.Execute<List<Email>>(request);
            Logger.Info("ResponseUri: " + response.ResponseUri);
            Logger.Info("StatusCode: " + response.StatusCode);
            Logger.Info("ResponseStatus: " + response.ResponseStatus);
            Logger.Info("Content: " + response.Content);

            //TODO: Check response codes...
            return response.Data;
        }

        public static void insertAttachment(string baseUrl, string oauthToken, string oauthTokenType, string accountId, string emailAddress/*, Attachment attachment */)
        {

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
            Logger.DebugFormat("StatusCode: {0} {1}", restResponse.StatusCode, restResponse.StatusDescription);
            Logger.DebugFormat("ResponseStatus: {0}", restResponse.ResponseStatus);
            Logger.DebugFormat("Content: {0}", restResponse.Content);
            
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

        public static void doTest()
        {
            Logger.Info("RestSharpTest");

            OAuth oAuth = login(Config.Instance.TestUrl, Config.Instance.TestUsername, Config.Instance.TestPassword);

            putRegisteredClient(Config.Instance.TestUrl, oAuth.value, oAuth.tokenType, "THISISANOTHERMUFUKINTEST");

            string registeredClient = getRegisteredClient(Config.Instance.TestUrl, oAuth.value, oAuth.tokenType);
            Logger.Info("Client: " + registeredClient);

            //getIntegrationFolder(Config.Instance.TestUrl, oAuth.value, oAuth.tokenType);
        }

    }

}

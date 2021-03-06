using Microsoft.SharePoint.Client;
//using PnP.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
//using System.Net;
using System.Net.Http;
using System.Web.Http;
using Microsoft.ApplicationInsights;
using System.Globalization;
using System.Net;
using SN_DEV_JIW_TEST.Entities;

namespace SN_DEV_JIW_TEST.Controllers
{
    public class ValuesController : ApiController
    {
        // GET api/values
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        [HttpGet]
        [Route("api/TestSharePointIntegration/GetContextTest")]
        public IHttpActionResult TestAccessToken()
        {
            try
            {
                //string siteUrl = "https://jciadqa.sharepoint.com/sites/SFProjInteg";
                //string keyVaultEndpoint = "https://kv-sn-dev-001.vault.azure.net/";
                //string userAgent = "NONISV|JohnsonControls|SelNav.Integration.Web/1.0";
                ////string clientId = "8d8c127f-6a3c-48a5-9538-0b329f5eb5fe";
                ////tring clientSecret = "ERNLnW6K11vlwf21wZ4i8KbiiKCZmUNOrJgmWHlY/wU=";
                //telemetryClient.TrackTrace("Calling Access Token Method");
                //var kv = new KeyVaultHelper(keyVaultEndpoint);
                //var cliendID = kv.RetrieveSecret($"spoClientId-JIW");
                //var cliendSecret = kv.RetrieveSecret($"spoClientSecret-JIW");

                //telemetryClient.TrackTrace($"DATA FROM KEYVAULT :: {cliendID} :: {cliendSecret}");

                //Uri siteUri = new Uri(siteUrl);
                ////Get the realm for the URL
                //string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
                //string HostedAppHostName = WebConfigurationManager.AppSettings.Get("HostedAppHostName");
                //string resource = GetFormattedPrincipal(TokenHelper.SharePointPrincipal, siteUri.Authority, realm);
                //string clientId = GetFormattedPrincipal(cliendID, HostedAppHostName, realm);

                //telemetryClient.TrackTrace($"DATA2:Resource: {resource} :clientId: {clientId} :HostedAppHostName: {HostedAppHostName} :HostedAppHostName: {realm}");

                //// tc.TrackEvent($"Nikhil : realm: {realm}");
                ////Get the access token for the URL.  
                ////   Requires this app to be registered with the tenant
                //string accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;

                //ClientContext clientContext = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), accessToken, userAgent);

                ////var oWebsite = clientContext.Web;
                ////ListCollection collList = oWebsite.Lists;
                ////clientContext.Load(collList);
                ////clientContext.ExecuteQuery();

                ////var count = collList.Count();
                ///

                //TESTING IN MY OWN TENANT

                //Lisandro Tenant
                string clientId = "c8df37d8-749d-46d0-8eb6-f85b9c786657";
                string clientSecret = "SxyiDr4EldeK1d54qouPvjrMNbxXk588DVYVUfOfvv0=";
                string provisioningUrl = "https://lisandrorossi444.sharepoint.com/";

                //using (var clientContext = new AuthenticationManager().GetACSAppOnlyContext(provisioningUrl, clientId, clientSecret))
                //{
                //    var oWebsite = clientContext.Web;
                //    ListCollection collList = oWebsite.Lists;
                //    List spList = clientContext.Web.Lists.GetByTitle("TestList");

                //    clientContext.Load(collList);
                //    clientContext.Load(spList);

                //    ListItemCollection items = spList.GetItems(CamlQuery.CreateAllItemsQuery());
                //    clientContext.Load(items); // loading all the fields
                //    clientContext.ExecuteQuery();

                //    foreach (var item in items)
                //    {

                //        item["Title"] = "modificado por la azure app service de prueba";

                //        item.Update();
                //    }
                //    clientContext.ExecuteQuery(); // important, commit changes to the server


                //    clientContext.ExecuteQuery();

                //    var count = collList.Count();
                //};

                return Ok("Success");
            }
            catch (Exception e)
            {
                string message = $"EXEPTION FROM KEYVAULT";
                //telemetryClient.TrackException(new ExceptionTelemetry() { Message = e.Message });
                //telemetryClient.TrackTrace($"Error :: {e.StackTrace} :: {e.InnerException}");

                //return Content(HttpStatusCode.BadRequest, "Error Message:" + e.Message + "\n" + e.StackTrace);

                return Ok("Error Message:" + e.Message + "\n" + e.StackTrace);
            }
        }

        [HttpGet]
        [Route("api/TestSharePointIntegration/GetContextTestOld")]
        public IHttpActionResult TestAccessTokenOld()
        {
            try
            {
                //string siteUrl = "https://jciadqa.sharepoint.com/sites/SFProjInteg";
                //string keyVaultEndpoint = "https://kv-sn-dev-001.vault.azure.net/";
                //string userAgent = "NONISV|JohnsonControls|SelNav.Integration.Web/1.0";
                ////string clientId = "8d8c127f-6a3c-48a5-9538-0b329f5eb5fe";
                ////tring clientSecret = "ERNLnW6K11vlwf21wZ4i8KbiiKCZmUNOrJgmWHlY/wU=";
                var telemetryClient = new TelemetryClient();
                //Lisandro Tenant
                string clientId = "c8df37d8-749d-46d0-8eb6-f85b9c786657";
                string clientSecret = "SxyiDr4EldeK1d54qouPvjrMNbxXk588DVYVUfOfvv0=";
                string provisioningUrl = "https://jciadqa.sharepoint.com/sites/SFProjInteg"; //"https://lisandrorossi444.sharepoint.com/";

                telemetryClient.TrackTrace("Calling Access Token Method");
                //var kv = new KeyVaultHelper(keyVaultEndpoint);
                //var cliendID = kv.RetrieveSecret($"spoClientId-JIW");
                //var cliendSecret = kv.RetrieveSecret($"spoClientSecret-JIW");

                Uri siteUri = new Uri(provisioningUrl);
                //Get the realm for the URL
                string realm = "9d89ee4f-c590-47c4-be02-c3fc82611185"; //string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
                string HostedAppHostName = "JIW"; // WebConfigurationManager.AppSettings.Get("HostedAppHostName");
                string resource = GetFormattedPrincipal(TokenHelper.SharePointPrincipal, siteUri.Authority, realm);
                string clientIdFormated = GetFormattedPrincipal(clientId, HostedAppHostName, realm);

                telemetryClient.TrackTrace($"DATA2:Resource: {resource} :clientId: {clientId} :HostedAppHostName: {HostedAppHostName} :HostedAppHostName: {realm}");


                // tc.TrackEvent($"Nikhil : realm: {realm}");
                //Get the access token for the URL.  
                //   Requires this app to be registered with the tenant

                TokenHelper.ClientId = "8d8c127f-6a3c-48a5-9538-0b329f5eb5fe";
                TokenHelper.ClientSecret = "ERNLnW6K11vlwf21wZ4i8KbiiKCZmUNOrJgmWHlY/wU=";
                string accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;

                string userAgent = "NONISV|JohnsonControls|SelNav.Integration.Web/1.0";

                ClientContext clientContext = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), accessToken, userAgent);

                //var oWebsite = clientContext.Web;
                //ListCollection collList = oWebsite.Lists;
                //List spList = clientContext.Web.Lists.GetByTitle("TestList");

                //clientContext.Load(collList);
                //clientContext.Load(spList);

                //ListItemCollection items = spList.GetItems(CamlQuery.CreateAllItemsQuery());
                //clientContext.Load(items); // loading all the fields
                //clientContext.ExecuteQuery();

                //foreach (var item in items)
                //{

                //    item["Title"] = "modificado por la azure app service de prueba";

                //    item.Update();
                //}
                //clientContext.ExecuteQuery(); // important, commit changes to the server

                return Ok("Success");
            }
            catch (Exception e)
            {
                string message = $"EXEPTION FROM KEYVAULT";
                //telemetryClient.TrackException(new ExceptionTelemetry() { Message = e.Message });
                //telemetryClient.TrackTrace($"Error :: {e.StackTrace} :: {e.InnerException}");

                //return Content(HttpStatusCode.BadRequest, "Error Message:" + e.Message + "\n" + e.StackTrace);

                return Ok("Error Message:" + e.Message + "\n" + e.StackTrace);
            }
        }

        [HttpPost]
        [Route("api/TestSharePointIntegration/AddUserToSPGroup")]
        public IHttpActionResult AddUserToSPGroup([FromBody] AddUserModelRequest parameters)
        {
            try
            {
          
                var telemetryClient = new TelemetryClient();
                //Lisandro Tenant
                string clientId = "c8df37d8-749d-46d0-8eb6-f85b9c786657";
                string clientSecret = "SxyiDr4EldeK1d54qouPvjrMNbxXk588DVYVUfOfvv0=";
                string provisioningUrl = "https://lisandrorossi444.sharepoint.com/";

                telemetryClient.TrackTrace("Calling Access Token Method");
                Uri siteUri = new Uri(provisioningUrl);
                //Get the realm for the URL
                // this should works after update rpoject to framework 4.7.2 or change the run time in the webconfig
                string realm = "629fd4e8-9d26-4da5-85ff-cc01ca1948c4"; //string realm = TokenHelper.GetRealmFromTargetUrl(siteUri); 
                string HostedAppHostName = "JIW"; // WebConfigurationManager.AppSettings.Get("HostedAppHostName");
                string resource = GetFormattedPrincipal(TokenHelper.SharePointPrincipal, siteUri.Authority, realm);
                string clientIdFormated = GetFormattedPrincipal(clientId, HostedAppHostName, realm);

                telemetryClient.TrackTrace($"DATA2:Resource: {resource} :clientId: {clientId} :HostedAppHostName: {HostedAppHostName} :HostedAppHostName: {realm}");


                TokenHelper.ClientId = clientId;
                TokenHelper.ClientSecret = clientSecret;
                string accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;

                string userAgent = "NONISV|JohnsonControls|SelNav.Integration.Web/1.0";

                ClientContext clientContext = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), accessToken, userAgent);

                var oWebsite = clientContext.Web;


                // Get group object using group name
                Group oGroup = oWebsite.SiteGroups.GetByName(parameters.groupName);

                // Get user using Logon name
                User oUser = oWebsite.EnsureUser(parameters.userName);

                // Add user to the group
                oGroup.Users.AddUser(oUser);

                clientContext.ExecuteQuery();


                //ListCollection collList = oWebsite.Lists;
                //List spList = clientContext.Web.Lists.GetByTitle("TestList");

                //clientContext.Load(collList);
                //clientContext.Load(spList);

                //ListItemCollection items = spList.GetItems(CamlQuery.CreateAllItemsQuery());
                //clientContext.Load(items); // loading all the fields
                //clientContext.ExecuteQuery();

                //foreach (var item in items)
                //{

                //    item["Title"] = "modificado por la azure app service de prueba";

                //    item.Update();
                //}
                //clientContext.ExecuteQuery(); // important, commit changes to the server

                return Ok("Success");
            }
            catch (Exception e)
            {
                string message = $"EXEPTION FROM KEYVAULT";
                //telemetryClient.TrackException(new ExceptionTelemetry() { Message = e.Message });
                //telemetryClient.TrackTrace($"Error :: {e.StackTrace} :: {e.InnerException}");

                //return Content(HttpStatusCode.BadRequest, "Error Message:" + e.Message + "\n" + e.StackTrace);

                return Ok("Error Message:" + e.Message + "\n" + e.StackTrace);
            }
        }

        [HttpPost]
        [Route("api/TestSharePointIntegration/GetRealm")]
        public static string GetRealmFromTargetUrl(string provisioningUrl)
        {
            Uri siteUri = new Uri(provisioningUrl);
            WebRequest request = WebRequest.Create(siteUri + "/_vti_bin/client.svc");
            request.Headers.Add("Authorization: Bearer ");

            try
            {
                using (request.GetResponse())
                {
                }
            }
            catch (WebException e)
            {
                if (e.Response == null)
                {
                    return null;
                }

                string bearerResponseHeader = e.Response.Headers["WWW-Authenticate"];
                if (string.IsNullOrEmpty(bearerResponseHeader))
                {
                    return null;
                }

                const string bearer = "Bearer realm=\"";
                int bearerIndex = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal);
                if (bearerIndex < 0)
                {
                    return null;
                }

                int realmIndex = bearerIndex + bearer.Length;

                if (bearerResponseHeader.Length >= realmIndex + 36)
                {
                    string targetRealm = bearerResponseHeader.Substring(realmIndex, 36);

                    Guid realmGuid;

                    if (Guid.TryParse(targetRealm, out realmGuid))
                    {
                        return targetRealm;
                    }
                }
            }
            return null;
        }

        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        public void Post([FromBody] string value)
        {
        }

        // PUT api/values/5
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }

        #region Helpers
        private static string GetFormattedPrincipal(string principalName, string hostName, string realm)
        {
            if (!String.IsNullOrEmpty(hostName))
            {
                return String.Format(CultureInfo.InvariantCulture, "{0}/{1}@{2}", principalName, hostName, realm);
            }

            return String.Format(CultureInfo.InvariantCulture, "{0}@{1}", principalName, realm);
        }
        #endregion
    }
}
